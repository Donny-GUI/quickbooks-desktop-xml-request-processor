import sys
import ctypes
import winreg
import win32api
import win32con
import win32event
import pythoncom
from win32com.client import constants as C
from win32com.client import Dispatch, DispatchEx, gencache
from win32com.server import util
from win32com.server.exception import COMException



# Constants
dwTimeOut = 5000  # time for EXE to be idle before shutting down
dwPause = 1000  # time to wait for threads to finish up

# GUIDs and other data types
CLSID_RequestProcessor2 = "{E64B99E6-25F1-11D4-8CAE-0080C792E5D8}"
LIBID_QBXMLRP2ELib = "{A056F9FC-75D9-11D3-A184-0080C792E5D8}"
LPCTSTR = ctypes.c_wchar_p
LPTSTR = ctypes.c_wchar_p
DWORD = ctypes.c_ulong
HANDLE = ctypes.c_void_p
LPDWORD = ctypes.POINTER(DWORD)
LPVOID = ctypes.c_void_p
PVOID = ctypes.c_void_p
ULONG_PTR = ctypes.c_ulonglong

# Windows API functions and structs
kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
user32 = ctypes.WinDLL('user32', use_last_error=True)
comctl32 = ctypes.WinDLL('comctl32', use_last_error=True)
advapi32 = ctypes.WinDLL('advapi32', use_last_error=True)
ole32 = ctypes.WinDLL('ole32', use_last_error=True)

LPSECURITY_ATTRIBUTES = ctypes.POINTER(ctypes.c_void_p)
LPSTARTUPINFO = ctypes.POINTER(ctypes.c_void_p)
LPPROCESS_INFORMATION = ctypes.POINTER(ctypes.c_void_p)
LPCTSTR = ctypes.c_wchar_p

# Resource IDs
IDR_QBXMLRPE = 101


class QBServerUtil:
    @staticmethod
    def DeleteRegValue(mainKey, keyPath, keyName):
        try:
            with winreg.OpenKey(mainKey, keyPath, 0, winreg.KEY_WRITE) as key:
                winreg.DeleteValue(key, keyName)
                return True
        except OSError:
            return False
    
    @staticmethod
    def SetRegValue(mainKey, keyPath, keyName, value):
        try:
            with winreg.OpenKey(mainKey, keyPath, 0, winreg.KEY_WRITE) as key:
                winreg.SetValueEx(key, keyName, 0, winreg.REG_SZ, value)
                return True
        except OSError:
            return False
    
    @staticmethod
    def IsRegValueExist(mainKey, keyPath, keyName):
        try:
            with winreg.OpenKey(mainKey, keyPath, 0, winreg.KEY_READ) as key:
                winreg.QueryValueEx(key, keyName)
                return True
        except OSError:
            return False
    
    @staticmethod
    def Register9xService(bRegister):
        # implementation omitted
        pass
    
    @staticmethod
    def IsNT():
        # implementation omitted
        pass
    
    def __init__(self):
        pass
    
    def __del__(self):
        pass


class CRequestProcessor:
    _reg_clsid_ = '{FBE5B340-8AEE-11D1-00A0-C90600000000}'
    _reg_progid_ = 'QBXMLRP2.RequestProcessor2'
    _public_methods_ = ['get_QBXMLVersionsForSession',
                        'get_ReleaseNumber', 'get_ReleaseLevel', 'get_MinorVersion',
                        'get_MajorVersion', 'GetCurrentCompanyFileName', 'EndSession',
                        'BeginSession', 'CloseConnection', 'ProcessRequest', 'OpenConnection',
                        'get_ConnectionType', 'ProcessSubscription', 'get_QBXMLVersionsForSubscription',
                        'OpenConnection2', 'get_AuthPreferences']
    _public_attrs_ = []
    _readonly_attrs_ = []

    def __init__(self):
        self.qbXMLRPPtr = None
        self.authPrefsEPtr = None
        self.VerifyQBXMLRP()

    def get_QBXMLVersionsForSession(self, ticket):
        return self.qbXMLRPPtr.get_QBXMLVersionsForSession(ticket)

    def get_ReleaseNumber(self):
        return self.qbXMLRPPtr.get_ReleaseNumber()

    def get_ReleaseLevel(self):
        return self.qbXMLRPPtr.get_ReleaseLevel()

    def get_MinorVersion(self):
        return self.qbXMLRPPtr.get_MinorVersion()

    def get_MajorVersion(self):
        return self.qbXMLRPPtr.get_MajorVersion()

    def GetCurrentCompanyFileName(self, ticket):
        return self.qbXMLRPPtr.GetCurrentCompanyFileName(ticket)

    def EndSession(self, ticket):
        self.qbXMLRPPtr.EndSession(ticket)

    def BeginSession(self, qbFileName, reqFileMode):
        return self.qbXMLRPPtr.BeginSession(qbFileName, reqFileMode)

    def CloseConnection(self):
        self.qbXMLRPPtr.CloseConnection()

    def ProcessRequest(self, ticket, inputRequest):
        return self.qbXMLRPPtr.ProcessRequest(ticket, inputRequest)

    def OpenConnection(self, appID, appName):
        self.qbXMLRPPtr.OpenConnection(appID, appName)

    def get_ConnectionType(self):
        return self.qbXMLRPPtr.get_ConnectionType()

    def ProcessSubscription(self, inputRequest):
        return self.qbXMLRPPtr.ProcessSubscription(inputRequest)

    def get_QBXMLVersionsForSubscription(self):
        return self.qbXMLRPPtr.get_QBXMLVersionsForSubscription()

    def OpenConnection2(self, appID, appName, connPref):
        self.qbXMLRPPtr.OpenConnection2(appID, appName, connPref)

    def get_AuthPreferences(self):
        return self.authPrefsEPtr

    def VerifyQBXMLRP(self):
        try:
            self.qbXMLRPPtr = gencache.EnsureDispatch('QBXMLRP2.RequestProcessor4')
            self.authPrefsEPtr = self.qbXMLRPPtr.get_AuthPreferences()
        except:
            raise Exception("Unable to initialize QBXMLRP2.RequestProcessor4")



# Passed to CreateThread to monitor the shutdown event
def MonitorProc(p):
    p.MonitorShutdown()

class CExeModule(util.Handle):
    _public_methods_ = ['StartMonitor']
    _reg_clsctx_ = pythoncom.CLSCTX_LOCAL_SERVER
    _create_thread_ = True
    _objects_in_threads_ = False
    _use_apartment_ = pythoncom.COINIT_APARTMENTTHREADED
    _register_ = True

    def __init__(self):
        self.hEventShutdown = win32event.CreateEvent(None, False, False, None)
        self.dwThreadID = win32api.GetCurrentThreadId()
        self.m_nLockCnt = 0
        self.bActivity = False

    def Unlock(self):
        l = util.Handle.Unlock(self)

        # if this is Win95/98/ME
        # do not let this server unload if lock count gets to zero
        if QBServerUtil.IsNT():
            if l == 0:
                self.bActivity = True
                win32event.SetEvent(self.hEventShutdown) # tell monitor that we transitioned to zero

        return l

    # Monitors the shutdown event
    def MonitorShutdown(self):
        while 1:
            rc = win32event.WaitForSingleObject(self.hEventShutdown, win32event.INFINITE)
            dwWait = 0
            while 1:
                self.bActivity = False
                dwWait = win32event.WaitForSingleObject(self.hEventShutdown, dwTimeOut)
                if dwWait != win32event.WAIT_OBJECT_0:
                    break
            # timed out
            if not self.bActivity and self.m_nLockCnt == 0: # if no activity let's really bail
                break

        win32event.CloseHandle(self.hEventShutdown)
        win32api.PostThreadMessage(self.dwThreadID, win32con.WM_QUIT, 0, 0)

    def StartMonitor(self):
        h = win32api.CreateThread(None, 0, MonitorProc, self, 0, None)
        return h != 0

_Module = CExeModule()

# This registers our object, and is called by regsvr32
def DllRegisterServer():
    try:
        import winerror
    except ImportError:
        winerror = None

    print("Registring server...")
    from win32com.server.register import UseCommandLine
    UseCommandLine(_Module)

# This unregisters our object, and is called by regsvr32 /u
def DllUnregisterServer():
    try:
        import winerror
    except ImportError:
        winerror = None

    print("Unregistring server...")
    from win32com.server.register import UseCommandLine
    UseCommandLine(_Module, "/Unregister")
    

class CExeModule:
    """CExeModule class."""
    def __init__(self):
        self.hEventShutdown = None
        self.dwThreadID = None
        self.m_nLockCnt = 0
        self.bActivity = False

    def Unlock(self):
        l = comctl32.CComModule.Unlock(ctypes.byref("CExeModule"))
        if not QBServerUtil.IsNT() and l == 0:
            self.bActivity = True
            kernel32.SetEvent(self.hEventShutdown)
        return l

    def MonitorShutdown(self):
        while True:
            kernel32.WaitForSingleObject(self.hEventShutdown, kernel32.INFINITE)
            dwWait = 0
            while True:
                self.bActivity = False
                dwWait = kernel32.WaitForSingleObject(self.hEventShutdown, dwTimeOut)
                if dwWait != kernel32.WAIT_OBJECT_0:
                    break
            if not self.bActivity and self.m_nLockCnt == 0:
                if sys.getwindowsversion().major >= 4 and hasattr(comctl32, 'CoSuspendClassObjects'):
                    comctl32.CoSuspendClassObjects()
                    if not self.bActivity and self.m_nLockCnt == 0:
                        break
                else:
                    break
        kernel32.CloseHandle(self.hEventShutdown)
        user32.PostThreadMessageW(self.dwThreadID, win32con.WM_QUIT, 0, 0)

    def StartMonitor(self):
        self.hEventShutdown = kernel32.CreateEventW(None, False, False, None)
        if not self.hEventShutdown:
            return False
        dwThreadID = DWORD()
        hThread = kernel32.CreateThread(None, 0, ctypes.cast("MonitorProc", PVOID), ctypes.byref(self), 0, ctypes.byref(dwThreadID))
        return bool
