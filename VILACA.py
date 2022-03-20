from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# comando para abrir o navegador
navegador = webdriver.Chrome()

# escolha de unidade, tipo de visualização, click em atualizar dados
navegador.get("http://printer:81/PrintSuperVision/Printers.aspx?GroupP_ID=157")
navegador.get("http://printer:81/PrintSuperVision/Printers.aspx?PropertiesView=756")
navegador.find_element_by_xpath('//*[@id="pgform"]/p/input').click()

arquivo = open('C:\\TMP\\verificador_de_suprimentos\\Suprimentos_VILACA.txt', 'w')
arquivo.write("")
arquivo.close()

lista = [25, 31, 38, 41, 43, 60, 61, 62, 63, 187, 190, 191, 192, 193, 194, 195, 196, 197, 198, 228, 271, 276, 291, 296, 315]

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
        
        vnomeimp = str(nomeimp.text)
        vmodelo = str(modelo.text)

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
                                                suprimentos = [vtoner, vciano, vmagenta, vyellow, vcilindro, cciano, cyellow, cmagenta, vfusor, vesteira]
                                                nsup = [" Toner Preto ", " Toner Ciano ", " Toner Magenta ", " Toner Yellow ", " Cilindro Preto ", " Cilindro Ciano ", " Cilindro Yellow ", " Cilindro Magenta ", " Fusor ",
                                                        " Esteira "]
                                                contnome = 0
                                                cont = 0

                                                for y in suprimentos:

                                                    if y <= 20:
                                                        if contnome == 0:
                                                            arquivo = open('C:\\TMP\\verificador_de_suprimentos\\Suprimentos_VILACA.txt', 'a')
                                                            arquivo.write(vnomeimp + " | " + vmodelo + " | " + nsup[cont] + str(y) + " | ")
                                                            arquivo.close()
                                                            contnome = contnome + 1
                                                            print("verificando")
                                                        else:
                                                            arquivo = open('C:\\TMP\\verificador_de_suprimentos\\Suprimentos_VILACA.txt', 'a')
                                                            arquivo.write(nsup[cont] + str(y) + " | ")
                                                            arquivo.close()

                                                    else:
                                                        print("verificando")
                                                    cont = cont + 1
                                                if contnome > 0:
                                                    arquivo = open('C:\\TMP\\verificador_de_suprimentos\\Suprimentos_VILACA.txt', 'a')
                                                    arquivo.write("\n")
                                                    arquivo.close()

    finally:
        print(x)

print("fim")
navegador.quit()
