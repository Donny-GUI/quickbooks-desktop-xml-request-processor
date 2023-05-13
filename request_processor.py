
import win32com.client
import ctypes


class SAFEARRAY:
    def __init__(self, data, vt_type):
        self._data = data
        self._vt_type = vt_type
        self._array = win32com.client.VARIANT(win32com.client.VT_ARRAY | vt_type, data)
        
    def __len__(self):
        return len(self._data)
    
    def __getitem__(self, index):
        return self._array[index]
    
    def __setitem__(self, index, value):
        self._array[index] = value
    
    def tolist(self):
        return list(self._data)
    
    
class BSTR(ctypes.c_wchar_p):
    def __init__(self, value=None):
        if value is not None:
            self.value = value
        else:
            self.value = ''

    def __repr__(self):
        return f'BSTR({self.value!r})'

    def __str__(self):
        return self.value

    def __len__(self):
        return ctypes.c_int(len(self.value))

    def __eq__(self, other):
        if isinstance(other, BSTR):
            return self.value == other.value
        elif isinstance(other, str):
            return self.value == other
        return False

    def __ne__(self, other):
        return not self.__eq__(other)

    def __add__(self, other):
        if isinstance(other, BSTR):
            return BSTR(self.value + other.value)
        elif isinstance(other, str):
            return BSTR(self.value + other)
        else:
            raise TypeError(f'Cannot concatenate BSTR with type {type(other).__name__}')

    def __radd__(self, other):
        if isinstance(other, str):
            return BSTR(other + self.value)
        else:
            raise TypeError(f'Cannot concatenate {type(other).__name__} with BSTR')

    def __getitem__(self, index):
        return self.value[index]

    def __setitem__(self, index, value):
        self.value = self.value[:index] + value + self.value[index+1:]

    def __delitem__(self, index):
        self.value = self.value[:index] + self.value[index+1:]

    def __contains__(self, item):
        return item in self.value

    def __hash__(self):
        return hash(self.value)

    def to_string(self):
        return ctypes.cast(self, ctypes.c_char_p)

    @classmethod
    def from_param(cls, value):
        if value is None or isinstance(value, BSTR):
            return value
        elif isinstance(value, str):
            return cls(value)
        else:
            raise TypeError(f'Cannot convert {type(value).__name__} to BSTR')

class QBFileModeE:
    pass 

class Result:
    def __init__(self, success: bool, message: str):
        self.success = success
        self.message = message


class RequestProcessor:
    handle = "QBXMLRP2.RequestProcessor2"
    dispatcher = "win32com.client.Dispatch"
    
    def __init__(self):
        self.qbXMLRPPtr = None

    def verify(self) -> Result:
        """ Verify that the quickbooks extensive markup language request processor is working

        Returns:
            Result: Result of the quickbooks com object fetch
        """
        if self.qbXMLRPPtr is None:
            try:
                self.qbXMLRPPtr = win32com.client.Dispatch("QBXMLRP2.RequestProcessor2")
            except:
                return Result(success=-1, message=f"Failed to verify {self.handle} with {self.dispatcher}")
        return Result(success=0, message=f"{self.handle} verified with {self.dispatcher}")

    def open_connection(self, app_id: str, app_name: str):
        
        hr = self.verify()
        if hr != 0:
            return hr
        return self.qbXMLRPPtr.OpenConnection(app_id, app_name)

    def process_request(self, ticket: str, input_request: str) -> str:
        if input_request is None:
            return ''
        if self.qbXMLRPPtr is None:
            return ''
        output_response = self.qbXMLRPPtr.ProcessRequest(ticket, input_request)
        return output_response

    def close_connection(self) -> Result:
        hr = self.verify()
        if hr != 0:
            return hr
        hr = self.qbXMLRPPtr.CloseConnection()
        self.qbXMLRPPtr = None
        return hr

    def begin_session(self, qb_filename: str, req_filemode: QBFileModeE ) -> tuple:
        hr = self.verify()
        if hr != 0:
            return hr, ''
        ticket = self.qbXMLRPPtr.BeginSession(qb_filename, req_filemode)
        return hr, ticket

    def end_session(self, ticket: str) -> str:
        hr = self.verify()
        if hr != 0:
            return hr
        return self.qbXMLRPPtr.EndSession(ticket)

    def get_current_company_filename(self, ticket: str) -> str:
        if self.qbXMLRPPtr is None:
            return ''
        return self.qbXMLRPPtr.GetCurrentCompanyFileName(ticket)

    @property
    def major_version(self) -> str:
        if self.qbXMLRPPtr is None:
            return ''
        return self.qbXMLRPPtr.MajorVersion

    @property
    def minor_version(self) -> str:
        if self.qbXMLRPPtr is None:
            return ''
        return self.qbXMLRPPtr.MinorVersion

    @property
    def release_level(self) -> str:
        if self.qbXMLRPPtr is None:
            return ''
        return self.qbXMLRPPtr.ReleaseLevel

    @property
    def release_number(self) -> str:
        if self.qbXMLRPPtr is None:
            return ''
        return self.qbXMLRPPtr.ReleaseNumber

    def qbxmlversions_for_session(self, ticket: str):
        if self.qbXMLRPPtr is None:
            return ''
        return self.qbXMLRPPtr.QBXMLVersionsForSession(ticket)

    @property
    def connection_type(self) -> str:
        if self.qbXMLRPPtr is None:
            return ''
        return self.qbXMLRPPtr.ConnectionType

    def process_subscription(self, input_request: str) -> str:
        if input_request is None:
            return ''
        if self.qbXMLRPPtr is None:
            return ''
        output_response = self.qbXMLRPPtr.ProcessSubscription(input_request)
        return output_response
    