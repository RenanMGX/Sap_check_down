import win32gui #type: ignore
import win32con #type: ignore
import win32api #type: ignore
from time import sleep 

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

    
    