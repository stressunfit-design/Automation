import xlUtil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

driver = webdriver.Chrome("C:\\Users\\stressunfit\\PycharmProjects\\VTUSGPAAutomation\\ChromeDriver\\chromedriver.exe")
driver.maximize_window()
driver.implicitly_wait(10) # seconds
driver.get("https://vtu-sgpa.weebly.com/vtu-sgpa-calculator.html")

driver.find_element(By.CLASS_NAME, "leadform-popup-close").click()
driver.implicitly_wait(15)
driver.find_element(By.ID, "leadform-popup-close-de8d83b3-18cc-42ae-9fa5-b399f3887d13").click()
driver.find_element(By.XPATH, "//*[@id='wsite-content']/div/div/div/div/div/div[5]/strong/a").click()



path = "C:\\Users\\stressunfit\\PycharmProjects\\VTUSGPAAutomation\\Book2.xlsx"

#Loop through all the marks
rows = xlUtil.getRowCount(path,'Sheet1')
for r in range(2,rows+1):
    Sub1 = xlUtil.readData(path,'Sheet1',r,2)
    Sub2 = xlUtil.readData(path,'Sheet1',r,3)
    Sub3 = xlUtil.readData(path, 'Sheet1', r, 4)
    Sub4 = xlUtil.readData(path, 'Sheet1', r, 5)
    Sub5 = xlUtil.readData(path, 'Sheet1', r, 6)
    Sub6 = xlUtil.readData(path, 'Sheet1', r, 7)
    Sub7 = xlUtil.readData(path, 'Sheet1', r, 8)
    Sub8 = xlUtil.readData(path, 'Sheet1', r, 9)




    # Enters marks to respective Subjects
    driver.find_element(By.ID,"n1").send_keys(Sub1)
    driver.find_element(By.ID,"n2").send_keys(Sub2)
    driver.find_element(By.ID,"n3").send_keys(Sub3)
    driver.find_element(By.ID,"n4").send_keys(Sub4)
    driver.find_element(By.ID,"n5").send_keys(Sub5)
    driver.find_element(By.ID,"n6").send_keys(Sub6)
    driver.find_element(By.ID,"n7").send_keys(Sub7)
    driver.find_element(By.ID,"n8").send_keys(Sub8)

    # Clicks on Calculate
    driver.find_element(By.CLASS_NAME,"button1").click()
    Sgpa = driver.find_element(By.XPATH,"//*[@id='result']").get_attribute("value")
    print(Sgpa)
    xlUtil.writeData(path,"Sheet1",r,10,Sgpa)
    driver.find_element(By.CLASS_NAME,"button2").click()

driver.quit()