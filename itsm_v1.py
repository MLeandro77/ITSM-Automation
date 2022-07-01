import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager

evento = []
contador = int('0')
for contador in range(3):
    evento.append(input('Digite o evento: '))
    contador + 1

def linha(txt):
    print('-'*41)
    print(txt)
    print('-'*41)

# Instalação da última versão do Webdriver (Microsoft Edge)
driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()))

# Maximizar janela
driver.maximize_window()

# Abertura da página inicial do ITSM
driver.get("url.omitida")
time.sleep(2)

for evt in evento:
    if evt != (''):

        # Acessando o iFrame default (Necessário para fechar mais de 1 EVT)
        driver.switch_to.default_content()

        # Click na lupa de busca da página inicial
        lupa = driver.find_element(By.XPATH, '/html/body/div/div/div/header/div[1]/div/div[2]/div/div[4]/form').click()

        # Inserção e busca pelo EVENTO
        busca = driver.find_element(By.XPATH, '//*[@id="sysparm_search"]')
        busca.send_keys(evt)
        busca.send_keys(Keys.ENTER)
        time.sleep(2)
        busca.click()
        busca.send_keys(Keys.ESCAPE)

        # Acessando o iFrame onde estão os elementos da página
        driver.switch_to.frame("gsft_main")

        # ****** Priorization - Baixando a Prioridade do ticket ******
        # Impact
        dropdown = driver.find_element(By.NAME, "u_rim_event.impact")
        ddown = Select(dropdown)
        ddown.select_by_visible_text("Significant")
        # Urgency
        dropdown = driver.find_element(By.NAME, "u_rim_event.urgency")
        ddown = Select(dropdown)
        ddown.select_by_visible_text("Normal")

        # ****** APLICAÇÃO DO TEMPLATE Incident Event ******
        dotsClick = driver.find_element(By.XPATH, '//*[@id="toggleMoreOptions"]').click()
        tempBar = driver.find_element(By.XPATH, '//*[@id="template-toggle-button"]').click()
        dotsBarClick = driver.find_element(By.XPATH, '//*[@id="template-bar-aria-container"]/div/button[1]').click()
        filtTemp = driver.find_element(By.XPATH, '//*[@id="overflowTemplateSearch"]')
        filtTemp.send_keys("Incident Event")
        tempClick = driver.find_element(By.XPATH, '//*[@id="templateOverflowContainer"]/div/ul/li[6]/a[1]').click()
        time.sleep(5)
        saveBtn = driver.find_element(By.XPATH, '//*[@id="sysverb_update_and_stay"]').click()
        time.sleep(5)

        # ****** APLICAÇÃO DO TEMPLATE (Oscilação de Link) ******
        dotsBarClick = driver.find_element(By.XPATH, '//*[@id="template-bar-aria-container"]/div/button[1]').click()
        filtTemp = driver.find_element(By.XPATH, '//*[@id="overflowTemplateSearch"]')
        filtTemp.send_keys("Oscilação de Link")
        tempClick = driver.find_element(By.XPATH, '//*[@id="templateOverflowContainer"]/div/ul/li[11]/a[1]').click()
        time.sleep(5)
        barClose = driver.find_element(By.XPATH, '//*[@id="template-bar-aria-container"]/div/button[3]').click()

        # ****** SELEÇÃO DO STATUS CLOSE OR CANCEL TASK ******
        dropdown = driver.find_element(By.NAME, "u_rim_event.u_next_step_displayed")
        ddown = Select(dropdown)
        ddown.select_by_visible_text("Close or cancel task")

        # ****** Obter valores dos campos de FILA e NOME ******
        fila = driver.find_element(By.XPATH, '//*[@id="sys_display.u_rim_event.u_owner_group"]').get_attribute('value')
        nome = driver.find_element(By.XPATH, '//*[@id="sys_display.u_rim_event.assigned_to"]').get_attribute('value')

        # ****** Campos Ownership & Assignment ******

        # Responsible Group
        resGroupCheck = bool(driver.find_element(By.XPATH, '//*[@id="sys_display.u_rim_event.u_responsible_owner_group"]').get_attribute('value'))
        
        if resGroupCheck == False or resGroupCheck != (fila):
            resGroup = driver.find_element(By.XPATH, '//*[@id="sys_display.u_rim_event.u_responsible_owner_group"]')
            resGroup.clear()
            resGroup.send_keys(fila)
            time.sleep(2)
            
        # Responsible Owner
        resOwnerCheck = bool(driver.find_element(By.XPATH, '//*[@id="sys_display.u_rim_event.u_responsible_owner"]').get_attribute('value'))
        if resOwnerCheck == False or resOwnerCheck != (nome):
            resOwner = driver.find_element(By.XPATH, '//*[@id="sys_display.u_rim_event.u_responsible_owner"]')
            resOwner.clear()
            resOwner.send_keys(nome)
            time.sleep(2)

        # ****** CAMPOS DA ABA CLOSURE DETAILS ******

        # Aba Closure Details
        abaClosDet = driver.find_element(By.XPATH, '/html/body/div[2]/form/div[1]/span[6]/span[1]').click()

        # Closure Details
        closDetails = driver.find_element(By.XPATH, '//*[@id="u_rim_event.close_notes"]')
        closDetails.clear()
        closDetails.send_keys("Oscilação no link da operadora, normalizado sem intervenção.")

        # Resolved by
        resolvedByCheck = bool(driver.find_element(By.XPATH, '//*[@id="sys_display.u_rim_event.u_resolved_by"]').get_attribute('value'))    
        if resolvedByCheck == False:
            resolvedBy = driver.find_element(By.XPATH, '//*[@id="sys_display.u_rim_event.u_resolved_by"]')
            resolvedBy.send_keys(nome)
            time.sleep(2)

        # Clique do botão Save ao finalizar o processo
        saveBtn = driver.find_element(By.XPATH, '//*[@id="sysverb_update_and_stay"]').click()
        time.sleep(10)

        # ****** SELEÇÃO DO STATUS CLOSE FINAL ******
        dropdown = driver.find_element(By.NAME, "u_rim_event.u_next_step_displayed")
        ddown = Select(dropdown)
        ddown.select_by_visible_text("Set to closed")
        time.sleep(1)

        # ****** OK DO ALERT FINAL ******
        alert = Alert(driver)
        alert.accept()
        #alert.accept()
        saveBtn = driver.find_element(By.XPATH, '//*[@id="sysverb_update_and_stay"]').click()
        os.system("cls")
        linha('Validando o encerramento. Aguarde...')
        time.sleep(5)

        # Validando o encerramento ou não do evento informado
        encerra = driver.find_element(By.XPATH, '//*[@id="sys_readonly.u_rim_event.state"]').get_attribute('value')
        if encerra == '7':
            # os.system("cls")
            linha('O ' + evt + ' foi encerrado com sucesso!')
            time.sleep(5)
        else:
            # os.system("cls")
            (linha('Não foi possível encerrar o ' + evt + '!'))
            os.system("pause")
                
driver.close()
os.system("pause")
driver.quit()
