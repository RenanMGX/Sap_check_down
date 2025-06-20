from Entities.sap_check import SapCheck, sleep, datetime
from Entities.dependencies.functions import P
import traceback
import os
os.environ['date'] = datetime.now().isoformat()

if __name__ == "__main__":
    while True:
        print(os.environ['date'])
        try:
            count = 1
            max_range = 10
            for i in range(max_range):
                try:
                    sap = SapCheck()
                    janela = sap.find_per_title("SAP GUI for Windows 770")
                    if janela:
                        sap.para_frente(janela[0])
                        sap.aperta_enter(janela[0])
                        print(P(f"[{count}/{max_range}] A janela {janela} foi fechada com sucesso!", color="green"))
                    else:
                        print(P(f"[{count}/{max_range}]", color="red"))
                except Exception as err:
                    print(P(f"[{count}/{max_range}] - {err}", color="red"))
                    print(traceback.format_exc())
                
                count += 1
                del sap
                sleep(1)
            
            try:
                SapCheck.fechar_app_sap(3) 
            except:
                with open("log.txt", "a") as log_file:
                    log_file.write(traceback.format_exc())
        except:
            pass
        
        sleep(60 * 1)
