import os
import time
import datetime
import re
import xmlrpclib

from subprocess import Popen, PIPE, STDOUT
from threading import Thread, Event

from win32api import OpenProcess, TerminateProcess
from win32event import WaitForSingleObject
import win32pdh
import pywintypes

import win32serviceutil

from ftplib import FTP

from SimpleXMLRPCServer import SimpleXMLRPCServer
from csv import writer as csvWriter

from optparse import OptionParser
from getopt import GetoptError

# TODO funcs to:
#   getClock changeClock


def startProcess(img, args, flag=None):
    if not flag:
        flag='P_DETACH'
    flag = getattr(os, flag)
    img = os.path.normpath(img)
    print img
    print args
    return os.spawnv(flag, img, (r'"%s"' % img,) + tuple(args))


def stopProcess(imagename, username=None):
    PROCESS_TERMINATE = 1

    pid = _getPid(imagename.lower(), username)
    if not pid:
        return -1
    try:
        handle = OpenProcess(PROCESS_TERMINATE, False, pid)
    except pywintypes.error:        # case pid already exited
        return -1

    try:
        try:
            TerminateProcess(handle, -1)
            WaitForSingleObject(handle, 0)
        except pywintypes.error:    # case process is already dead
            return -1
    finally:
        handle.close()
    return 0

def startService(service):
    if win32serviceutil.QueryServiceStatus(service)[1] == 4:
        pass
    else:
        win32serviceutil.StartService(service)
    return 0

def stopService(service):
    if win32serviceutil.QueryServiceStatus(service)[1] == 4:
        win32serviceutil.StopService(service)
    return 0

def downloadFileFromFTP(from_path, to_path, user='sample-user-name', password='sample-password'):
    server    = from_path.strip().split('/')[2]
    directorys = from_path.strip().split('/')[3:-1]
    thefile   = from_path.strip().split('/')[-1]
    fileh = open(to_path, 'wb')
    ftp = FTP(server)
    ftp.login(user, password)
    for adir in directorys:
        ftp.cwd(adir)
    ftp.retrbinary("RETR " + thefile, fileh.write)
    fileh.close()
    ftp.quit()
    return 0


def _getPid(imagename, username=None):
    img = imagename.lower()
    if username is not None:
        username = username.lower()
        cmd = r'tasklist /FI "imagename eq %s.exe" /FI '
        cmd += r'"username eq %s" /FO csv'
        cmd %= (img, username)
    else:
        cmd = r'tasklist /FI "imagename eq %s.exe" /FO csv"' % img

    out = Popen(cmd, stdout=PIPE, stderr=STDOUT).communicate()[0]

    if out == '' or out.startswith('INFO:'):
        return None

    out = out.strip().split('\n')[1:]
    if len(out) > 1:
        raise AgentError('multiple %s processes' % img)

    pid = int(out[0].split(',')[1].strip('"'))
    return pid


class IxNetPerfMon(object):
    __thread = None
    __stopEvent = Event()
    __csvFile = None

    def startIxNetPerfMon(self, seconds):
        if self.__class__.__thread is not None:
            self._stopPMon()
            self.__class__.__csvFile.close()
            self.__class__.__csvFile = None

        self.__class__.__csvFile = os.tmpfile()
        th = Thread(target=self.__ixNPerfMon, args=(self.__class__.__csvFile,
                                        seconds))
        self.__class__.__thread = th

        th.start()
        return 0

    def stopIxNetPerfMon(self):
        self._stopPMon()
        fh = self.__class__.__csvFile

        fh.seek(0)
        res = xmlrpclib.Binary(fh.read())
        self.__class__.__csvFile = None
        fh.close()

        return res

    @classmethod
    def _stopPMon(cls):
        if cls.__thread is not None:
            cls.__stopEvent.set()
            cls.__thread.join()
            cls.__thread = None
            cls.__stopEvent.clear()

    def __ixNPerfMon(self, fh, seconds):
        hQuery = win32pdh.OpenQuery(None, 0)

        pathList = [win32pdh.MakeCounterPath(p) for p in \
            [(None, '.NET CLR Memory', 'ixLoad', None, -1,
                    '# Bytes in all Heaps'),
            (None, '.NET CLR Memory', 'IxNetwork', None, -1,
                    '# Bytes in all Heaps'),
            (None, 'Process', 'ixLoad', None, -1, '% Processor Time'),
            (None, 'Process', 'ixLoad', None, -1, 'Handle Count'),
            (None, 'Process', 'ixLoad', None, -1, 'Private Bytes'),
            (None, 'Process', 'ixLoad', None, -1, 'Virtual Bytes'),
            (None, 'Process', 'IxNetwork', None, -1, '% Processor Time'),
            (None, 'Process', 'IxNetwork', None, -1, 'Handle Count'),
            (None, 'Process', 'IxNetwork', None, -1, 'Private Bytes'),
            (None, 'Process', 'IxNetwork', None, -1, 'Virtual Bytes')]]

        pTipe = re.compile(r'.*Processor.*')

        hCntList = []
        for p in pathList:
            if pTipe.match(p):
                fmt = win32pdh.PDH_FMT_DOUBLE
            else:
                fmt = win32pdh.PDH_FMT_LONG
            hCntList.append((win32pdh.AddCounter(hQuery, p, 0), fmt))

        csv = csvWriter(fh)
        csv.dialect.lineterminator = '\n'
        csv.dialect.quoting = 1

        hdr = ['Time', 'Elapsed Time']
        hdr.extend(pathList)
        csv.writerow(hdr)

        elapsed = 0
        try:
            while True:
                if self.__class__.__stopEvent.isSet():
                    break

                row = [str(datetime.datetime.now()), elapsed]
                try:
                    win32pdh.CollectQueryData(hQuery)
                except pywintypes.error:
                    time.sleep(seconds)
                    continue
                for hd in hCntList:
                    try:
                        val = win32pdh.GetFormattedCounterValue(*hd)[1]
                    except pywintypes.error:
                        val = None
                    row.append(val)

                csv.writerow(row)
                elapsed += seconds
                time.sleep(seconds)
        finally:
            win32pdh.CloseQuery(hQuery)


class AgentError(Exception):
    pass

if __name__ == '__main__':
    parser = OptionParser()

    parser.add_option("-p", "--port", dest='port', type=int,
                    default=8000, help="Port to use.")

    try:
        options, args = parser.parse_args()
    except GetoptError:
        usage()
        sys.exit(2)

    server = SimpleXMLRPCServer(("", options.port))
    print "PyThar Agent listening on port %d..." % options.port
    server.register_multicall_functions()
    server.register_function(startProcess, 'startProcess')
    server.register_function(stopProcess, 'stopProcess')
    server.register_function(startService, 'startService')
    server.register_function(stopService, 'stopService')
    server.register_function(downloadFileFromFTP, 'downloadFileFromFTP')
    server.register_instance(IxNetPerfMon())
    server.serve_forever()
