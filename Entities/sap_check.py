import win32gui #type: ignore
import win32con #type: ignore
import win32api #type: ignore
import win32com.client
from time import sleep 
import traceback
import psutil
import win32process
import os
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

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
    def fechar_app_sap(time_to_close:int=2):
        if os.environ['date'] == "":
            os.environ['date'] = datetime.now().isoformat()
            
        try:
            SapGuiAuto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")# type: ignore
            application: win32com.client.CDispatch = SapGuiAuto.GetScriptingEngine# type: ignore
        except:
            print("O SAP GUI não está aberto ou o SAP GUI Scripting não está habilitado.")
            time_date = datetime.fromisoformat(os.environ['date'])
            if datetime.now() >= (time_date + relativedelta(minutes=time_to_close)):
                if contar_janelas_sap() > 1:
                    SapCheck.encerrando_tarefa("saplogon")
                else:
                    print("O SAP está com mais de 1 janela aberta então não sera finalizado")
                os.environ['date'] = datetime.now().isoformat()
            else:
                print(f"precisa estar aberto há '{time_to_close}' para ser encerrado!")
                print(f"falta {(time_date + relativedelta(minutes=time_to_close)) - datetime.now()} minutos")
            return
        
        esta_aberto = False
        for i in range(0, 6):
            try:
                application.Children(i)
                esta_aberto = True
            except Exception as e:
                if 'The enumerator of the collection cannot find an element with the specified index.' in traceback.format_exc():
                    #print(f"No SAP GUI session found. {i}")
                    continue
                else:
                    raise e
                
        if not esta_aberto:
            print("O SAP GUI não está em sessão então fechando.")
            time_date = datetime.fromisoformat(os.environ['date'])
            if datetime.now() >= (time_date + relativedelta(minutes=time_to_close)):
                SapCheck.encerrando_tarefa("saplogon")
                os.environ['date'] = datetime.now().isoformat()
            else:
                print(f"precisa estar aberto há '{time_to_close}' para ser encerrado!")
                print(f"falta {(time_date + relativedelta(minutes=time_to_close)) - datetime.now()} minutos")
                
            return
        else:
            os.environ['date'] = datetime.now().isoformat()
            print("O sap está em sessão")
                
        
        
    @staticmethod
    def encerrando_tarefa(tarefa:str):
            for process in psutil.process_iter(['name']):
                if tarefa in process.info['name'].lower():
                    process.kill()
                    print(f"processo '{process.info['name']}' foi finalizada!")
                    
        
def contar_janelas_sap():
    def enum_windows():
        def callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                try:
                    proc = psutil.Process(pid)
                    windows.append((proc.name(), win32gui.GetWindowText(hwnd)))
                except psutil.NoSuchProcess:
                    pass
            return True

        windows = []
        win32gui.EnumWindows(callback, windows)
        return windows

    janelas_sap = 0
    for name, title in enum_windows():
        if "sap" in name.lower():
            janelas_sap += 1

    return janelas_sap



if __name__ == "__main__":
    os.environ['date'] = datetime.now().isoformat()
    #print(contar_janelas_sap())
    sap = SapCheck()
    sap.fechar_app_sap()
    
      
    # janela = sap.find_per_title("SAP GUI for Windows 770")
    # print(janela)
    # if janela:
    #     sap.para_frente(janela[0])
    #     sleep(1)
    #     sap.para_frente(janela[0])
    #     sap.para_frente(janela[0])
    #     sleep(1)
    #     sap.para_frente(janela[0])

    
    