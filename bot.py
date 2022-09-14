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
from datetime import datetime


# Uncomment the line below for integrations with BotMaestro
# Using the Maestro SDK
# from botcity.maestro import *


class Bot(DesktopBot):
    def action(self, execution=None):
        # Fetch the Activity ID from the task:
        # task = self.maestro.get_task(execution.task_id)
        # activity_id = task.activity_id

        import pandas as pd
        basedados_urbano = pd.read_excel(r'C:\Users\rafael\Downloads\INDICADOR_REAL.xlsx', 'urbano', keep_default_na=False)
        basedados_autonoma = pd.read_excel(r'C:\Users\rafael\Downloads\INDICADOR_REAL.xlsx', 'autonoma', keep_default_na=False)
        basedados_rural = pd.read_excel(r'C:\Users\rafael\Downloads\INDICADOR_REAL.xlsx', 'rural', keep_default_na=False)
        arquivo = open(r'C:\Users\rafael\Downloads\log.txt', 'w')

        for i in range(1):
            dataatual = datetime.now()

            if not self.find("matricula", matching=0.97, waiting_time=10000):
                self.not_found("matricula")
            self.double_click_relative(170, 15)
            self.copy_to_clipboard(str(basedados_urbano["MATRICULA"][i]))
            self.paste()
            self.enter()
            self.enter()
            arquivo.write(str(basedados_urbano["MATRICULA"][i]))
            arquivo.write(" - ")
            arquivo.write(dataatual.strftime('%d/%m/%Y %H:%M'))
            arquivo.write("\n")
            arquivo.write("------------------- ")
            arquivo.write("\n")

            if not self.find("alterar", matching=0.97, waiting_time=10000):
                self.not_found("alterar")
            self.click()
            self.copy_to_clipboard(str(basedados_urbano["IMOVEL"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_urbano["LOTE"][i]))
            self.paste()
            self.enter()

            if self.find("matcad"):
                if not self.find("cancel", matching=0.97, waiting_time=10000):
                    self.not_found("cancel")
                self.click()
                self.enter()
                self.copy_to_clipboard(str(basedados_urbano["QUADRA"][i]))
                self.paste()
            else:
                self.copy_to_clipboard(str(basedados_urbano["QUADRA"][i]))
                self.paste()
            self.enter()

            if self.find("matcad"):
                if not self.find("cancel", matching=0.97, waiting_time=10000):
                    self.not_found("cancel")
                self.click()
                self.enter()
                self.copy_to_clipboard(str(basedados_urbano["SETOR"][i]))
                self.paste()
            else:
                self.copy_to_clipboard(str(basedados_urbano["SETOR"][i]))
                self.paste()
            self.enter()

            self.copy_to_clipboard(str(basedados_urbano["IND.FISCAL"][i]))
            self.paste()
            self.enter()

            if self.find("atencao"):
                if not self.find("ok_atencao", matching=0.97, waiting_time=10000):
                    self.not_found("ok_atencao")
                self.click()
                if not self.find("click_localizacao", matching=0.97, waiting_time=10000):
                    self.not_found("click_localizacao")
                self.click_relative(116, 9)

            else:
                self.wait(10)

            self.copy_to_clipboard(str(basedados_urbano["LOCALIZACAO"][i]))
            self.paste()
            self.enter()
            self.delete()
            self.wait(500)
            self.copy_to_clipboard(str(basedados_urbano["AREA"][i]))
            self.paste()
            self.enter()
            self.page_up()

            tipo_unidademedidaurbano = str(basedados_urbano["UNID"][i]).upper()
            if tipo_unidademedidaurbano == "M":
                recorte1 = "metro"
            elif tipo_unidademedidaurbano == "H":
                recorte1 = "hectare"
            elif tipo_unidademedidaurbano == "ALQ.":
                recorte1 = "Alqueire"
            else:
                recorte1 = "unidade_vazio"
            if not self.find(recorte1, matching=0.97, waiting_time=10000):
                self.not_found(recorte1)
            self.click()
            self.enter()
            self.copy_to_clipboard(str(basedados_urbano["AREA CONSTR"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_urbano["PL FISCAL"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_urbano["BENFEITORIA"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_urbano["VAGAS"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_urbano["CEP"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_urbano["CIDADE"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_urbano["ESTADO"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_urbano["VIA"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_urbano["ENDERECO"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_urbano["NUMERO"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_urbano["COMPLEMENTO"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_urbano["BAIRRO"][i]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_urbano["OBS"][i]))
            self.paste()
            if not self.find("salvar", matching=0.97, waiting_time=10000):
                self.not_found("salvar")
            self.click()
            self.space()

        for j in range(1):
            dataatual = datetime.now()

            if not self.find("matricula", matching=0.97, waiting_time=10000):
                self.not_found("matricula")
            self.double_click_relative(170, 15)
            self.copy_to_clipboard(str(basedados_autonoma["MATRICULA"][j]))
            self.paste()
            self.enter()
            self.enter()
            arquivo.write(str(basedados_urbano["MATRICULA"][j]))
            arquivo.write(" - ")
            arquivo.write(dataatual.strftime('%d/%m/%Y %H:%M'))
            arquivo.write("\n")
            arquivo.write("------------------- ")
            arquivo.write("\n")
            
            if not self.find("alterar", matching=0.97, waiting_time=10000):
                self.not_found("alterar")
            self.click()
            self.copy_to_clipboard(str(basedados_autonoma["IMOVEL"][j]))
            self.paste()
            self.enter()
            self.page_up()

            tipo_unidadeautonoma = str(basedados_autonoma["TIPO"][j]).upper()
            if tipo_unidadeautonoma == "RESIDENCIAL":
                recorte2 = "residencial"
            else:
                recorte2 = "comercial"
            if not self.find(recorte2, matching=0.97, waiting_time=10000):
                self.not_found(recorte2)
            self.click()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["TIPO_IMOVEL"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["NUMERO"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["ANDAR"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["BLOCO"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["PAVIMENTO"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["SETOR"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["IND.FISCAL"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["QUADRA"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["LOTE"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["LOCALIZACAO"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["EMPREENDIMENTO"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["AREA"][j]))
            self.paste()
            self.enter()
            self.page_up()

            tipo_unidademedidaautonoma = str(basedados_autonoma["UNID"][j]).upper()
            if tipo_unidademedidaautonoma == "Metro":
                recorte3 = "metro"
            else:
                recorte3 = "hectare"
            self.kb_type(recorte3[0])

            # if not self.find(recorte3, matching=0.97, waiting_time=10000):
            #     self.not_found(recorte3)
            self.click()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["AREA CONSTRUIDA"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["AREA PRIVATIVA"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["AREA USO COMUM"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["FRACAO"][j]))
            self.paste()
            self.enter()
            self.page_up()

            tipo_unidadefracao = str(basedados_autonoma["UNID FRACAO"][j]).upper()
            if tipo_unidadefracao == "%":
                self.type_down()
            else:
                recorte4 = "metro"
                self.kb_type(recorte4[0])
            self.enter()

            # if not self.find(recorte4, matching=0.999999, waiting_time=10000):
            #    self.not_found(recorte4)
            # self.click()
            # self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["OUTRAS AREAS"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["VAGAS"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["PL FISCAL"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["CEP"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["CIDADE"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["ESTADO"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["VIA"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["ENDERECO"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["NUMERO"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["COMPLEMENTO"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["BAIRRO"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["OBS"][j]))
            self.paste()

            if not self.find("area_priv_total", matching=0.97, waiting_time=10000):
                self.not_found("area_priv_total")
            self.double_click_relative(214, 11)
            self.double_click()
            self.copy_to_clipboard(str(basedados_autonoma["AREA PRIV TOTAL"][j]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_autonoma["AREA REAL TOTAL"][j]))
            self.paste()
            self.enter()

            if not self.find("salvar", matching=0.97, waiting_time=10000):
                self.not_found("salvar")
            self.click()
            self.space()

        for t in range(1):
            dataatual = datetime.now()

            if not self.find("matricula", matching=0.97, waiting_time=10000):
                self.not_found("matricula")
            self.double_click_relative(170, 15)
            self.copy_to_clipboard(str(basedados_rural["MATRICULA"][t]))
            self.paste()
            self.enter()
            self.enter()
            arquivo.write(str(basedados_urbano["MATRICULA"][i]))
            arquivo.write(" - ")
            arquivo.write(dataatual.strftime('%d/%m/%Y %H:%M'))
            arquivo.write("\n")
            arquivo.write("------------------- ")
            arquivo.write("\n")
            
            if not self.find("alterar", matching=0.97, waiting_time=10000):
                self.not_found("alterar")
            self.click()
            self.copy_to_clipboard(str(basedados_rural["IMOVEL"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_rural["LOTE"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_rural["QUADRA"][t]))
            self.paste()
            self.enter()

            if self.find("matcad"):
                if not self.find("cancel", matching=0.97, waiting_time=1000):
                    self.not_found("cancel")
                self.click()
                self.enter()
                self.copy_to_clipboard(str(basedados_rural["INCRA"][t]))
                self.paste()
            else:
                self.copy_to_clipboard(str(basedados_rural["INCRA"][t]))
                self.paste()
            self.enter()

            self.copy_to_clipboard(str(basedados_rural["ITR"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_rural["LOCALIZACAO"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_rural["AREA"][t]))
            self.paste()
            self.enter()
            self.page_up()

            tipo_unidademedidarural = str(basedados_rural["UNID"][t]).upper()
            if tipo_unidademedidarural == "M":
                recorte5 = "metro"
            elif tipo_unidademedidarural == "H":
                recorte5 = "hectare"
            elif tipo_unidademedidarural == "ALQ":
                recorte5 = "Alqueire"
            else:
                recorte5 = "metro"

            if not self.find(recorte5, matching=0.97, waiting_time=10000):
                self.not_found(recorte5)
            self.click()
            self.enter()
            self.copy_to_clipboard(str(basedados_rural["AREA CONSTR"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_rural["BENFEITORIA"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_rural["CEP"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_rural["CIDADE"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_rural["ESTADO"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_rural["VIA"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_rural["ENDERECO"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_rural["NUMERO"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_rural["COMPLEMENTO"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_rural["BAIRRO"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_rural["OBS"][t]))
            self.paste()

            if not self.find("certificado", matching=0.97, waiting_time=10000):
                self.not_found("certificado")
            self.double_click_relative(172, 11)
            self.copy_to_clipboard(str(basedados_rural["CERTIFICADO"][t]))
            self.paste()
            self.enter()
            self.copy_to_clipboard(str(basedados_rural["CAR"][t]))
            self.paste()

            if not self.find("salvar", matching=0.97, waiting_time=10000):
                self.not_found("salvar")
            self.click()
            self.space()
            arquivo.close()

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

