from selenium import webdriver
import time


browser_name = {'browserName': 'opera'}
driver = webdriver.Remote(command_executor='http://127.0.0.1:9515', desired_capabilities=browser_name)
time.sleep(2)
driver.execute_script('''window.open("https://mail.google.com/mail/u/0/?view=cm&fs=1&tf=1&source=mailto&su=Contato+|+Empresa+Júnior+PUC-Rio&to=pefandrade@hotmail.com","_blank");''')
#driver.get('https://mail.google.com/mail/u/0/?view=cm&fs=1&tf=1&source=mailto&su=Contato+|+Empresa+Júnior+PUC-Rio&to=pefandrade@hotmail.com')
body = driver.find_element_by_class_name('gmail_default')
body.send_keys('TESTANDO')
