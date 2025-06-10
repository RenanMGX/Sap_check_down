import win32gui #type: ignore
import win32con #type: ignore
import win32api #type: ignore
import win32com.client
from time import sleep 
import traceback
import psutil

class SapCheck:
    @property
    def lista(self):
        self.__checar_janelas()
        return self.__lista
    
    def __init__(self):
        self.__lista = {}
        
    def __checar_janelas(self):
        lista_temp = {}
        def enumHandler(hwnd, lParam):
            nonlocal lista_temp
            title = win32gui.GetWindowText(hwnd)
            lista_temp[hwnd] = title

        win32gui.EnumWindows(enumHandler, None)        
        self.__lista = lista_temp.copy()
        
    def find_per_title(self, title:str):
        result = [[hwnd, t] for hwnd, t in self.lista.items() if title == t]
        if result:
            return result[0]
        return []
    
    def para_frente(self, hwnd:int):
        # Traz a janela para frente
        win32gui.SetForegroundWindow(hwnd)

    
    def aperta_enter(self, hwnd:int):
        win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
    
    @staticmethod
    def fechar_app_sap():
        try:
            SapGuiAuto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")# type: ignore
            application: win32com.client.CDispatch = SapGuiAuto.GetScriptingEngine# type: ignore
        except:
            print("O SAP GUI não está aberto ou o SAP GUI Scripting não está habilitado.")
            return
        
        esta_aberto = False
        for i in range(0, 6):
            try:
                application.Children(i)
                esta_aberto = True
            except Exception as e:
                if 'The enumerator of the collection cannot find an element with the specified index.' in traceback.format_exc():
                    print(f"No SAP GUI session found. {i}")
                    continue
                else:
                    raise e
                
        if not esta_aberto:
            print("O SAP GUI não está em sessão então fehando.")
            for process in psutil.process_iter(['name']):
                if "saplogon" in process.info['name'].lower():
                    process.kill()
            return
        else:
            print("O sap está em sessão")
                
        

if __name__ == "__main__":
    sap = SapCheck()
    janela = sap.find_per_title("SAP GUI for Windows 770")
    print(janela)
    if janela:
        sap.para_frente(janela[0])
        sleep(1)
        sap.para_frente(janela[0])
        sap.para_frente(janela[0])
        sleep(1)
        sap.para_frente(janela[0])

    
    