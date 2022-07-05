from asyncio.windows_events import NULL
import datetime
import os
import sys
def WriteEx(append = True):
    if os.path.exists('OutputLogs.txt') :
        if append :
            log = open('OutputLogs.txt', 'a')
        else :
            log = open('OutputLogs.txt', 'w')
    else :
        log = open('OutputLogs.txt', 'w')
    now = datetime.datetime.now()
    exc_type, exc_obj, exc_tb = sys.exc_info()
    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
    current_time = now.strftime('%Y-%m-%d %H:%M:%S')
    ##
    log.write(f'[{current_time}]:\nException {exc_type} in {fname} line {exc_tb.tb_lineno}.\n\n')

def WriteError(ErrData, file, append = True):
    if os.path.exists('OutputLogs.txt') :
        if append :
            log = open('OutputLogs.txt', 'a')
        else :
            log = open('OutputLogs.txt', 'w')
    else :
        log = open('OutputLogs.txt', 'w')
    now = datetime.datetime.now()
    current_time = now.strftime('%Y-%m-%d %H:%M:%S')
    ##
    log.write(f'[{current_time}]:\nError {ErrData} in {file}.\n\n')

def ExceptionExit(msg):
    print(msg)
    os.system('pause')
    sys.exit()

def ReadFile(filename, scriptname) : 
    if os.path.exists(filename) :
        return open(filename, 'r')
    else :
        WriteError(f'{filename} File Not Found.', scriptname)
        ExceptionExit(f'{filename} File Not Found.')
