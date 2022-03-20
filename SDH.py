from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# comando para abrir o navegador
navegador = webdriver.Chrome()

# escolha de unidade, tipo de visualização, click em atualizar dados
navegador.get("http://printer:81/PrintSuperVision/Printers.aspx?GroupP_ID=78")
navegador.get("http://printer:81/PrintSuperVision/Printers.aspx?PropertiesView=756")
navegador.find_element_by_xpath('//*[@id="pgform"]/p/input').click()

# abre/cria arquivo de texto em branco para saída de dados
arquivo = open('C:\\TMP\\verificador_de_suprimentos\\Suprimentos_SDH.txt', 'w')
arquivo.write("")
arquivo.close()

#lista de numero de id's das impressoras
lista = [21, 24, 29, 45, 48, 49, 54, 55, 57, 59, 65, 68, 69, 72, 73, 76, 153, 211, 212, 213, 214, 215, 216, 217, 219,
         220, 223, 224, 227, 268, 277, 278, 279, 298, 303, 306, 307, 311]

# a pagina é estruturada de maneira que o padrão de id é a letra "p" seguido do nº da impressora seguido do id da coluna
# laço de repetição faz a procura dos campos a serem extraidos os dados, com a concatenação das informações acima
for x in lista:
    try:
        nomeimp = WebDriverWait(navegador, 2).until(
            EC.presence_of_element_located((By.ID, 'p' + str(x) + 'k123'))
        )
    except:
        pass

    else:
        modelo = WebDriverWait(navegador, 2).until(
            EC.presence_of_element_located((By.ID, 'p' + str(x) + 'k126'))
        )

        toner = WebDriverWait(navegador, 2).until(
            EC.presence_of_element_located((By.ID, 'p' + str(x) + 'k469'))
        )

        cilindro = WebDriverWait(navegador, 2).until(
            EC.presence_of_element_located((By.ID, 'p' + str(x) + 'k472'))
        )

        fusor = WebDriverWait(navegador, 2).until(
            EC.presence_of_element_located((By.ID, 'p' + str(x) + 'k465'))
        )

        ciano = WebDriverWait(navegador, 2).until(
            EC.presence_of_element_located((By.ID, 'p' + str(x) + 'k466'))
        )

        magenta = WebDriverWait(navegador, 2).until(
            EC.presence_of_element_located((By.ID, 'p' + str(x) + 'k467'))
        )

        yellow = WebDriverWait(navegador, 2).until(
            EC.presence_of_element_located((By.ID, 'p' + str(x) + 'k468'))
        )

        esteira = WebDriverWait(navegador, 2).until(
            EC.presence_of_element_located((By.ID, 'p' + str(x) + 'k464'))
        )
        
        cilciano = WebDriverWait(navegador, 2).until(
            EC.presence_of_element_located((By.ID, 'p' + str(x) + 'k463'))
        )
        
        cilyellow = WebDriverWait(navegador, 2).until(
            EC.presence_of_element_located((By.ID, 'p' + str(x) + 'k471'))
        )
        
        cilmagenta = WebDriverWait(navegador, 2).until(
            EC.presence_of_element_located((By.ID, 'p' + str(x) + 'k470'))
        )
        #conversão do nome e modelo para string
        vnomeimp = str(nomeimp.text)
        vmodelo = str(modelo.text)

        # verificação se campo esta em branco, para que não acuse valor 0 e alerte necessidade de troca é add valor "99"
        try:
            vtoner = int(toner.text)

        except ValueError:
            vtoner = 99

        finally:

            try:
                vcilindro = int(cilindro.text)

            except ValueError:
                vcilindro = 99

            finally:

                try:
                    vfusor = int(fusor.text)

                except ValueError:
                    vfusor = 99

                finally:

                    try:
                        vciano = int(ciano.text)

                    except ValueError:
                        vciano = 99

                    finally:

                        try:
                            vesteira = int(esteira.text)

                        except ValueError:
                            vesteira = 99

                        finally:
                            try:
                                vmagenta = int(magenta.text)

                            except ValueError:
                                vmagenta = 99

                            finally:
                                try:
                                    vyellow = int(yellow.text)

                                except ValueError:
                                    vyellow = 99

                                finally:
                                    try:
                                        cciano = int(cilciano.text)

                                    except ValueError:
                                        cciano = 99

                                    finally:
                                        try:
                                            cyellow = int(cilyellow.text)

                                        except ValueError:
                                            cyellow = 99

                                        finally:
                                            try:
                                                cmagenta = int(cilmagenta.text)

                                            except ValueError:
                                                cmagenta = 99

                                            finally:
                                                #criado listas onde indices coincidem para inserção das informações
                                                suprimentos = [vtoner, vciano, vmagenta, vyellow, vcilindro, cciano, cyellow, cmagenta, vfusor, vesteira]
                                                nsup = [" Toner Preto ", " Toner Ciano ", " Toner Magenta ", " Toner Yellow ", " Cilindro Preto ", " Cilindro Ciano ", " Cilindro Yellow ", " Cilindro Magenta ", " Fusor ",
                                                        " Esteira "]
                                                contnome = 0
                                                cont = 0
                                                #y representa o indice da lista de suprimentos e tbm da legenda a ser inserida no TXT
                                                #condicional onde o valor de cada suprimento adiciona uma linha de text com as informações caso seja <= a 20
                                                for y in suprimentos:

                                                    if y <= 20:
                                                        #caso seja o primeiro suprimento abaixo de 20 entra no if que inclui nome e modelo da impressora
                                                        if contnome == 0:
                                                            arquivo = open('C:\\TMP\\verificador_de_suprimentos\\Suprimentos_SDH.txt', 'a')
                                                            arquivo.write(vnomeimp + " | " + vmodelo + " | " + nsup[cont] + str(y) + " | ")
                                                            arquivo.close()
                                                            contnome = contnome + 1
                                                            print("verificando")
                                                        #caso a linha ja exista com nome e modelo o contador é maior que 0 e adiciona apenas o suprimento atual no fim da linha
                                                        else:
                                                            arquivo = open('C:\\TMP\\verificador_de_suprimentos\\Suprimentos_SDH.txt', 'a')
                                                            arquivo.write(nsup[cont] + str(y) + " | ")
                                                            arquivo.close()

                                                    else:
                                                        print("verificando")
                                                    cont = cont + 1

                                                #caso algum suprimento tenha sido descrito e as verificações acabaram, contador diz que é necessario pular a linha para o proximo
                                                #do caso contrario contador não foi alterado e não é necessario mudar a linha
                                                if contnome > 0:
                                                    arquivo = open('C:\\TMP\\verificador_de_suprimentos\\Suprimentos_SDH.txt', 'a')
                                                    arquivo.write("\n")
                                                    arquivo.close()

    finally:
        print(x)

print("fim")
navegador.quit()
