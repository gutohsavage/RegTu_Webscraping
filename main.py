from selenium import webdriver

import pandas as pd


from selenium.webdriver.common.by import By


PATH = "C:\pathja\chromedriver"
driver = webdriver.Chrome(PATH)

url = 'https://web.reg.tu.ac.th/registrar/class_info.asp?lang=th'



table = []

for kana in range (9,12):
    for i in range(2):
        driver.get(url)
        driver.find_element(By.XPATH,'/html/body/table/tbody/tr[1]/td[2]/table/tbody/tr[4]/td[2]/font[2]/select').click()
        driver.find_element(By.XPATH,'/html/body/table/tbody/tr[1]/td[2]/table/tbody/tr[4]/td[2]/font[2]/select/option[' + str(kana) + ']').click()
        driver.find_element(By.XPATH,'/html/body/table/tbody/tr[1]/td[2]/table/tbody/tr[6]/td[2]/table/tbody/tr[1]/td[2]/font[1]/select').click()
        driver.find_element(By.XPATH,'/html/body/table/tbody/tr[1]/td[2]/table/tbody/tr[6]/td[2]/table/tbody/tr[1]/td[2]/font[1]/select/option['+str(i+1)+']').click()

    #คลิก
        driver.find_element(By.XPATH,'/html/body/table/tbody/tr[1]/td[2]/table/tbody/tr[7]/td[2]/table/tbody/tr/td/font[3]/input').click()


        while True:
            lenoftr = driver.find_elements(By.XPATH, "/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr")
            for j in range(4,len(lenoftr)-1):
                saka = driver.find_element(By.XPATH,'/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr[' + str(j) + ']/td[2]/font')
                couresno = driver.find_element(By.XPATH,'/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr[' + str(j) + ']/td[3]/font')
                coursecode = driver.find_element(By.XPATH,'/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr[' + str(j) + ']/td[5]/font/a/b')
                coursename = driver.find_element(By.XPATH,'/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr[' + str(j) + ']/td[6]')
                prof = ""
                coursenamess = ""
                credit =  driver.find_element(By.XPATH,'/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr[' + str(j) + ']/td[7]')
                section = driver.find_element(By.XPATH,'/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr[' + str(j) + ']/td[8]/font/b')
                timetolearn = driver.find_element(By.XPATH,'/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr[' + str(j) + ']/td[9]')
                timetoexm = driver.find_element(By.XPATH,'/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr[' + str(j) + ']/td[10]')
                rub = driver.find_element(By.XPATH,'/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr[' + str(j) + ']/td[11]')
                left = driver.find_element(By.XPATH,'/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr[' + str(j) + ']/td[12]')
                status = driver.find_element(By.XPATH,'/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr[' + str(j) + ']/td[13]')
                sakas = saka.text
                coursecodes = coursecode.text
                couresnos = couresno.text
                coursenames = coursename.text
                for line in range(len(coursenames.splitlines())):
                    if line == 0:
                        coursenamess = coursenames.splitlines()[0]
                    else :
                        if coursenames.splitlines()[line][0] != "*":
                            prof = prof + coursenames.splitlines()[line] +" , "
                credits = credit.text
                sections = section.text
                timetolearns = timetolearn.text
                timetoexms = timetoexm.text
                rubs = rub.text
                lefts = left.text
                statuss = status.text
                faculy = driver.find_element(By.XPATH,'/html/body/table/tbody/tr[1]/td[2]/div[1]/font/b')
                faculys = faculy.text
                terms = i+1
                table.append((sakas, couresnos,coursecodes,coursenamess,prof,credits,sections,timetolearns,timetoexms,rubs,lefts,statuss,faculys,terms))

            try:
                href = driver.find_element(By.PARTIAL_LINK_TEXT, "[หน้าต่อไป]")
                href.click()
            except:
                break




df2 = pd.DataFrame(table,columns=['ศูนย์','หลักสูตร','รหัสวิชา','ชื่อวิชา','อาจารย์ผู้สอน','หน่วยกิต','Section','เวลา','เวลาสอบ','จำนวนรับ','เหลือ','สถานะ','คณะ','ภาคเรียนที่'])
df2.to_excel('data.xlsx',sheet_name='Fact Table')
driver.quit()







