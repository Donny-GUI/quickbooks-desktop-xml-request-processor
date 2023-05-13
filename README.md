# quickbooks-desktop-xml-request-processor
# Why?
For pure python Quickbooks desktop integration.
Seems like a language that specializes in data would be a great for an accountant application that specializes in data.


![class_rp](https://github.com/Donny-GUI/quickbooks-desktop-xml-request-processor/assets/108424001/462c7888-545c-4775-921c-380b1a097d17)


quickbooks desktop xml request processor implemented in python3.10.


# BSTR Class

```Python3
   
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
```

# Safe Array Class

```Python3

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
    

```
