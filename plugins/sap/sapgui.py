import os
import time
import subprocess
import traceback
import win32com.client

class SapLogonNotFoundError(Exception):
    pass

class SapGui():
    
    def __init__(self, environment: str, path: str = r'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe'):
        self._environment = environment
        self._path = path

    def _check_subprocess(self) -> bool:
        imagename = os.path.basename(self._path)
        tlcall = 'TASKLIST', '/FI', f'imagename eq {imagename}'
        tlproc = subprocess.Popen(tlcall, shell=True, stdout=subprocess.PIPE)
        message = str(tlproc.communicate()[0])
        return imagename in message

    def open(self) -> win32com.client.CDispatch:
        if not os.path.exists(self._path):
            raise SapLogonNotFoundError(f'SAP LOGON não foi encontrado. ({self._path})')
        
        if not self._path.lower().endswith('saplogon.exe'):
            raise SapLogonNotFoundError(f'Caminho indicado não é do executável do SAP GUI. ({self._path})')
        
        self.close()
        subprocess.Popen(self._path)
        time.sleep(5)
        
        sapgui_obj = win32com.client.GetObject('SAPGUI')
        application = sapgui_obj.GetScriptingEngine
        connection = application.OpenConnection(self._environment, True)
        session = connection.Children(0)


        return session

    def close(self):
        if self._check_subprocess():
            imagename = os.path.basename(self._path)
            subprocess.call(["taskkill", "/IM", imagename, "/T", "/F"])

    def __enter__(self) -> win32com.client.CDispatch:
        session = self.open()
        return session

    def __exit__(self, exc_type, exc_value, tb) -> bool:
        self.close()
        if exc_type is not None:
            traceback.print_exception(exc_type, exc_value, tb)
            return False # uncomment to pass exception through
        return True
    
class SapConnection():
    """Esse é para fazer parte do template... usa quem quer e herda quem quer"""

    def __init__(self, session: win32com.client.CDispatch):
        self.session = session

    def login(self, username: str, password: str):
        self.session.findById('wnd[0]/usr/txtRSYST-BNAME').text = username
        self.session.findById('wnd[0]/usr/pwdRSYST-BCODE').text = password
        self.session.findById('wnd[0]').sendVKey(0)

    def change_transaction(self, transaction: str):
        self.session.findById('wnd[0]/tbar[0]/okcd').text = f'/n{transaction}'
        self.session.findById('wnd[0]/tbar[0]/btn[0]').press()
        return self

    def capture_result(self) -> str:
        return self.session.findById('wnd[0]/sbar').text()
