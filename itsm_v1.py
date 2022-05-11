import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager

evento = input('Digite o evento: ')

#Linha abaixo para a instalação da última versão do Webdriver (Microsoft Edge)
driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()))


#Maximizar janela
driver.maximize_window()

#Abertura da página inicial do ITSM
driver.get("url.omitida")
time.sleep(2)

#Click na lupa de busca da página inicial
lupa = driver.find_element_by_xpath('/html/body/div/div/div/header/div[1]/div/div[2]/div/div[4]/form')
lupa.click()

#Inserção e busca pelo EVENTO
busca = driver.find_element_by_xpath('//*[@id="sysparm_search"]')
busca.send_keys(evento)
busca.send_keys(Keys.ENTER)
time.sleep(10)

#Acessando o iFrame onde estão os elementos da página
driver.switch_to.frame("gsft_main")

# ****** APLICAÇÃO DO TEMPLATE (Oscilação de Link) ******
dotsClick = driver.find_element_by_xpath('//*[@id="toggleMoreOptions"]').click()
tempBar = driver.find_element_by_xpath('//*[@id="template-toggle-button"]').click()
dotsBarClick = driver.find_element_by_xpath('//*[@id="template-bar-aria-container"]/div/button[1]').click()
filtTemp = driver.find_element_by_xpath('//*[@id="overflowTemplateSearch"]')
filtTemp.send_keys("Oscilação de Link")
tempClick = driver.find_element_by_xpath('//*[@id="templateOverflowContainer"]/div/ul/li[11]/a[1]').click()
time.sleep(5)
barClose = driver.find_element_by_xpath('//*[@id="template-bar-aria-container"]/div/button[3]').click()

# ****** SELEÇÃO DO STATUS CLOSE OR CANCEL TASK ******
dropdown = driver.find_element_by_name("u_rim_event.u_next_step_displayed")
ddown = Select(dropdown)
ddown.select_by_visible_text("Close or cancel task")

# ****** Obter valores dos campos de FILA e NOME ******
fila = driver.find_element_by_xpath('//*[@id="sys_display.u_rim_event.u_owner_group"]').get_attribute('value')
nome = driver.find_element_by_xpath('//*[@id="sys_display.u_rim_event.assigned_to"]').get_attribute('value')

# ****** Campos Ownership & Assignment ******

#Responsible Group
resGroup = driver.find_element_by_xpath('//*[@id="sys_display.u_rim_event.u_responsible_owner_group"]')
resGroup.clear()
resGroup.send_keys(fila)
time.sleep(2)
#resGroup.send_keys(Keys.ARROW_DOWN)
#resGroup.send_keys(Keys.ENTER)

#Responsible Owner
resOwner = driver.find_element_by_xpath('//*[@id="sys_display.u_rim_event.u_responsible_owner"]')
resOwner.clear()
resOwner.send_keys(nome)
time.sleep(2)
#resOwner.send_keys(Keys.ARROW_DOWN)
#resOwner.send_keys(Keys.ENTER)

# ****** CAMPOS DA ABA CLOSURE DETAILS ******

#Aba Closure Details
abaClosDet = driver.find_element_by_xpath('/html/body/div[2]/form/div[1]/span[6]/span[1]')
abaClosDet.click()

""" #Resolution Code
resCode = driver.find_element_by_xpath('//*[@id="sys_display.u_rim_event.u_task_resolution_code"]')
resCode.send_keys("Carrier")
time.sleep(5)
resCode.send_keys(Keys.ARROW_DOWN)
resCode.send_keys(Keys.ENTER)

#Root cause
rootCause = driver.find_element_by_xpath('//*[@id="sys_display.u_rim_event.u_task_rootcause"]')
rootCause.send_keys("Carrier")
time.sleep(5)
rootCause.send_keys(Keys.ARROW_DOWN)
rootCause.send_keys(Keys.ENTER) """

#Closure details
closDetails = driver.find_element_by_xpath('//*[@id="u_rim_event.close_notes"]')
closDetails.clear()
closDetails.send_keys("Oscilação no link da operadora, normalizado sem intervenção.")

""" #Root cause comments
rootComm = driver.find_element_by_xpath('//*[@id="u_rim_event.u_root_cause_comments"]')
rootComm.send_keys("Breve instabilidade no link da operadora.") """

#Resolved by
resolvedBy = driver.find_element_by_xpath('//*[@id="sys_display.u_rim_event.u_resolved_by"]')
resolvedBy.clear()
resolvedBy.send_keys(nome)
time.sleep(3)

#Clique do botão Save ao finalizar o processo
saveBtn = driver.find_element_by_xpath('//*[@id="sysverb_update_and_stay"]').click()
time.sleep(10)

# ****** SELEÇÃO DO STATUS CLOSE FINAL ******
dropdown = driver.find_element_by_name("u_rim_event.u_next_step_displayed")
ddown = Select(dropdown)
ddown.select_by_visible_text("Set to closed")
time.sleep(1)

# ****** OK DO ALERT FINAL ******
alert = Alert(driver)
alert.accept()
saveBtn = driver.find_element_by_xpath('//*[@id="sysverb_update_and_stay"]').click()
time.sleep(10)
driver.quit()