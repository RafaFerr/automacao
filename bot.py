"""
WARNING:

Please make sure you install the bot with `pip install -e .` in order to get all the dependencies
on your Python environment.

Also, if you are using PyCharm or another IDE, make sure that you use the SAME Python interpreter
as your IDE.

If you get an error like:
```
ModuleNotFoundError: No module named 'botcity'
```

This means that you are likely using a different Python interpreter than the one used to install the bot.
To fix this, you can either:
- Use the same interpreter as your IDE and install your bot with `pip install -e .`
- Use the same interpreter as the one used to install the bot (`pip install -e .`)

Please refer to the documentation for more information at https://documentation.botcity.dev/
"""

from botcity.core import DesktopBot


# Uncomment the line below for integrations with BotMaestro
# Using the Maestro SDK
# from botcity.maestro import *


class Bot(DesktopBot):
    def action(self, execution=None):
        # Fetch the Activity ID from the task:
        # task = self.maestro.get_task(execution.task_id)
        # activity_id = task.activity_id

        import pandas as pd
        basedados = pd.read_excel(r'excel.xlsx', 'urbano', keep_default_na=False)
        basedados4 = pd.read_excel(r'excel.xlsx', 'autonoma', keep_default_na=False)
        basedados5 = pd.read_excel(r'excel.xlsx', 'rural', keep_default_na=False)

        for i in range(0):

            if not self.find( "matricula", matching=0.97, waiting_time=10000):
                self.not_found("matricula")
            self.double_click_relative(170, 15)
            self.paste(str(basedados["MATRICULA"][i]))
            self.enter()
            self.enter()

            if not self.find( "alterar", matching=0.97, waiting_time=10000):
                self.not_found("alterar")
            self.click()
            self.copy_to_clipboard(str(basedados["IMOVEL"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados["LOTE"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados["QUADRA"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados["SETOR"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados["IND.FISCAL"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados["LOCALIZACAO"][i]))
            self.paste()
            self.enter()
            self.delete()
            self.wait(500)
            self.copy_to_clipboard(str(basedados["AREA"][i]))
            self.paste()
            self.enter()
            self.page_up()

            tipo_unidademedidaurbano = str(basedados["UNID"][i]).upper()
            if tipo_unidademedidaurbano == "M":
                recorte1 = "metro"
            elif tipo_unidademedidaurbano == "H":
                recorte1 = "hectare"
            elif tipo_unidademedidaurbano == "A":
                recorte1 = "Alqueire"
            else:
                recorte1 = "metro_2"
            if not self.find(recorte1, matching=0.97, waiting_time=10000):
                self.not_found(recorte1)
            self.click()
            self.enter()
            self.copy_to_clipboard(str(basedados["AREA CONSTR"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados["PL FISCAL"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados["BENFEITORIA"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados["VAGAS"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados["CEP"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados["CIDADE"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados["ESTADO"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados["VIA"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados["ENDERECO"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados["NUMERO"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados["COMPLEMENTO"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados["BAIRRO"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados["OBS"][i]))
            self.paste()
            if not self.find( "salvar", matching=0.97, waiting_time=10000):
                self.not_found("salvar")
            self.click()
            self.space()

        for j in range(0):

            if not self.find( "matricula", matching=0.97, waiting_time=10000):
                self.not_found("matricula")
            self.double_click_relative(170, 15)
            self.paste(str(basedados4["MATRICULA"][j]))
            self.enter()
            self.enter()
            if not self.find( "alterar", matching=0.97, waiting_time=10000):
                self.not_found("alterar")
            self.click()
            self.copy_to_clipboard(str(basedados4["IMOVEL"][j]))
            self.paste()
            self.enter()
            self.page_up()

            tipo_unidadeautonoma = str(basedados4["TIPO"][j]).upper()
            if tipo_unidadeautonoma == "R":
                recorte2 = "residencial"
            else:
                recorte2 = "comercial"
            if not self.find(recorte2, matching=0.97, waiting_time=10000):
                self.not_found(recorte2)
            self.click()
            self.enter()
            self.copy_to_clipboard(str(basedados4["TIPO_IMOVEL"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["NUMERO"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["ANDAR"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["BLOCO"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["PAVIMENTO"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["SETOR"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["IND.FISCAL"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["QUADRA"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["LOTE"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["LOCALIZACAO"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["EMPREENDIMENTO"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["AREA"][j]))
            self.paste()
            self.enter()
            self.page_up()

            tipo_unidademedidaautonoma = str(basedados4["UNID"][j]).upper()
            if tipo_unidademedidaautonoma == "M":
                recorte3 = "metro"
            else:
                recorte3 = "hectare"
            self.kb_type(recorte3[0])

            # if not self.find(recorte3, matching=0.97, waiting_time=10000):
            #     self.not_found(recorte3)
            self.click()
            self.enter()
            self.copy_to_clipboard(str(basedados4["AREA CONTRUIDA"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["AREA PRIVATIVA"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["AREA USO COMUM"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["FRACAO"][j]))
            self.paste()
            self.enter()
            self.page_up()

            tipo_unidadefracao = str(basedados4["UNID FRACAO"][j]).upper()
            if tipo_unidadefracao == "%":
                self.type_down()
            else:
                recorte4 = "metro"
                self.kb_type(recorte4[0])
            self.enter()

            #if not self.find(recorte4, matching=0.999999, waiting_time=10000):
            #    self.not_found(recorte4)
            #self.click()
            #self.enter()
            self.copy_to_clipboard(str(basedados4["OUTRAS AREAS"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["VAGAS"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["PL FISCAL"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["CEP"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["CIDADE"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["ESTADO"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["VIA"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["ENDERECO"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["NUMERO"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["COMPLEMENTO"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["BAIRRO"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["OBS"][j]))
            self.paste()

            if not self.find( "area_priv_total", matching=0.97, waiting_time=10000):
                self.not_found("area_priv_total")
            self.double_click_relative(214, 11)
            self.double_click()
            self.copy_to_clipboard(str(basedados4["AREA PRIV TOTAL"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados4["AREA REAL TOTAL"][j]))
            self.paste()
            self.enter()

            if not self.find( "salvar", matching=0.97, waiting_time=10000):
                self.not_found("salvar")
            self.click()
            self.space()

        for t in range(4):

            if not self.find( "matricula", matching=0.97, waiting_time=10000):
                self.not_found("matricula")
            self.double_click_relative(170, 15)
            self.paste(str(basedados5["MATRICULA"][t]))
            self.enter()
            self.enter()
            if not self.find( "alterar", matching=0.97, waiting_time=10000):
                self.not_found("alterar")
            self.click()
            self.copy_to_clipboard(str(basedados5["IMOVEL"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados5["LOTE"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados5["QUADRA"][t]))
            self.paste()
            self.enter()

            if self.find("matcad"):
                if not self.find("cancel", matching=0.97, waiting_time=1000):
                    self.not_found("cancel")
                self.click()
                self.copy_to_clipboard(str(basedados5["INCRA"][t]))
                self.paste()
            else:
                self.copy_to_clipboard(str(basedados5["INCRA"][t]))
                self.paste()
            self.enter()

            self.copy_to_clipboard(str(basedados5["ITR"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados5["LOCALIZACAO"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados5["AREA"][t]))
            self.paste()
            self.enter()
            self.page_up()

            tipo_unidademedidarural = str(basedados5["UNIDADE"][t]).upper()
            if tipo_unidademedidarural == "M":
                recorte5 = "metro"
            elif tipo_unidademedidarural == "H":
                recorte5 = "hectare"
            elif tipo_unidademedidarural == "A":
                recorte5 = "Alqueire"
            else:
                recorte5 = "metro_2"
            if not self.find(recorte5, matching=0.97, waiting_time=10000):
                self.not_found(recorte5)
            self.click()
            self.enter()
            self.copy_to_clipboard(str(basedados5["AREA CONSTR"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados5["BENFEITORIA"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados5["CEP"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados5["CIDADE"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados5["ESTADO"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados5["VIA"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados5["ENDERECO"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados5["NUMERO"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados5["COMPLEMENTO"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados5["BAIRRO"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados5["OBS"][t]))
            self.paste()

            if not self.find( "salvar", matching=0.97, waiting_time=10000):
                self.not_found("salvar")
            self.click()
            self.space()











            
            
            


            

            
            
            
            

            

            

           
            


            
            






                
                
                
               
                



                # Uncomment to mark this task as finished on BotMaestro
    # self.maestro.finish_task(
    #     task_id=execution.task_id,
    #     status=AutomationTaskFinishStatus.SUCCESS,
    #     message="Task Finished OK."
    # )


def not_found(self, label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()


