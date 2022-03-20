import win32com.client as win32

arquivo = open("C:\\TMP\\verificador_de_suprimentos\\Suprimentos_SDH.txt")
sdh = arquivo.read()

arquivo = open("C:\\TMP\\verificador_de_suprimentos\\Suprimentos_VILACA.txt")
pasjc = arquivo.read()

arquivo = open("C:\\TMP\\verificador_de_suprimentos\\Suprimentos_LITORAL.txt")
litoral = arquivo.read()

arquivo = open("C:\\TMP\\verificador_de_suprimentos\\Suprimentos_SEDE.txt")
sede = arquivo.read()


outlook = win32.Dispatch('outlook.application')

email = outlook.CreateItem(0)

email.To = "suporte-lista@unimedsjc.com.br"
email.Subject = 'SUPRIMENTOS DAS UNIDADES'
email.body = f""" Segue suprimentos para troca:

SDH:

{sdh}

PA SJC:

{pasjc}

Litoral:

{litoral}

SEDE:

{sede}
"""

email.Send()