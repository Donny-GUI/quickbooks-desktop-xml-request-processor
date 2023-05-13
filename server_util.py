import winreg


class CQBServerUtil:
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
