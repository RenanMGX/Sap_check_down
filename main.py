from Entities.sap_check import SapCheck, sleep, datetime
from patrimar_dependencies.functions import P
import traceback
import os
os.environ['date'] = datetime.now().isoformat()
from botcity.maestro import * #type:ignore


class ExecuteAPP:
    @staticmethod
    def start(maestro:BotMaestroSDK|None=None, *, log_name:str="check_sap_down_log"):
        if maestro:
            try:
                maestro.new_log_entry(activity_label=log_name, values={"texto": f"Iniciando em {datetime.now().isoformat}"})
            except:
                columns = [model.Column(
                    name="Texto",
                    label="texto"
                )]
                maestro.new_log(activity_label=log_name, columns=columns)
                maestro.new_log_entry(activity_label=log_name, values={"texto": f"Iniciando em {datetime.now().isoformat}"})
        
        
        print(os.environ['date'])
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
                    if maestro:
                        maestro.new_log_entry(activity_label=log_name, values={"texto": f"[{count}/{max_range}] A janela {janela} foi fechada com sucesso!"})
                else:
                    print(P(f"[{count}/{max_range}]", color="red"))
                    if maestro:
                        maestro.new_log_entry(activity_label=log_name, values={"texto": f"[{count}/{max_range}]"})
            except Exception as err:
                if maestro:
                    maestro.error(task_id=int(maestro.get_execution().task_id), exception=err) 

                print(P(f"[{count}/{max_range}] - {err}", color="red"))
                print(traceback.format_exc())
                
            count += 1
            try:
                del sap #type: ignore
            except:
                pass
            sleep(1)
            
            # try:
            #     SapCheck.fechar_app_sap(3) 
            # except:
            #     with open("log.txt", "a") as log_file:
            #         log_file.write(traceback.format_exc())        

if __name__ == "__main__":
    ExecuteAPP.start()
    