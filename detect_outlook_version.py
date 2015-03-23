# Python version of https://msdn.microsoft.com/en-us/library/office/dd941331.aspx
# "How to: Check the Version of Outlook"

from __future__ import print_function
from ctypes import windll, create_string_buffer, c_int, byref

#UINT MsiProvideQualifiedComponent(
#  _In_     LPCTSTR szComponent,
#  _In_     LPCTSTR szQualifier,
#  _In_     DWORD dwInstallMode,
#  _Out_    LPTSTR lpPathBuf,
#  _Inout_  DWORD *pcchPathBuf
#);

def installedOutlookVersion():
    dll = windll.LoadLibrary('msi.dll')
    
    outlookversions = {
        '{E83B4360-C208-4325-9504-0D23003A74A5}': '2013',
        '{1E77DE88-BCAB-4C37-B9E5-073AF52DFD7A}': '2010',
        '{24AAE126-0911-478F-A019-07B875EB9996}': '2007',
        '{BC174BAD-2F53-4855-A1D5-0D575C19B1EA}': '2003'
    }

    outlook32 = create_string_buffer('outlook.exe')
    outlook64 = create_string_buffer('outlook.x64.exe')
    INSTALLMODE_DEFAULT = c_int(0)
    pathlen = c_int(0)

    for component, version in outlookversions.items():
        componentbuf = create_string_buffer(component)
        if 0 == dll.MsiProvideQualifiedComponentA(byref(componentbuf),
                                                  byref(outlook32),
                                                  INSTALLMODE_DEFAULT,
                                                  None,
                                                  byref(pathlen)):
            return version
        if 0 == dll.MsiProvideQualifiedComponentA(byref(componentbuf),
                                                  byref(outlook64),
                                                  INSTALLMODE_DEFAULT,
                                                  None,
                                                  byref(pathlen)):
            return version

    return None

print(installedOutlookVersion())
