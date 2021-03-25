from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from xlwt import Workbook

PATH="C:\Program Files (x86)\chromedriver.exe"

driver = webdriver.Chrome(PATH)

#go to result website
driver.get("https://patnauniversity.ac.in/puresult2021/240321-MAT-bsc-sem1.php")

# Workbook is created 
wb = Workbook()

# add_sheet is used to create sheet. 
sheet1 = wb.add_sheet('1styear2019batch')

#Adding the headings
head=['Roll No.','REG NO',
      'NAME',"FATHER'S NAME",'COLLEGE NAME',
      'Paper-I','Paper-II',
      'Hons. Pract.',
      'Sub.- I Theory','Sub.- I Pract.',
      'Sub.- II Theory','Sub.- II Pract.',
      'Comp. (100 mks)','Comp. - 1 (50 Mks)','Comp. - 2 (50 Mks)',
      'Part -I Total','Result','Remarks']
		
for m in range(len(head)):
    sheet1.write(1, 1+m, head[m])

count=0
for roll in range(40117,40165):#Roll numbers: 40117-40164
    #entering the roll numbers
    search = driver.find_element_by_name("troll")
    search.send_keys(roll)
    search.send_keys(Keys.RETURN)

    text=driver.page_source
    output=[roll]
    
    start=6070
    for i in range(17):
        s=text.find('''<td>''',start)
        #print(s)
        e=text.find('''</td>''',s)
        #print(e)
        out=''
        for j in range(e-s-4):
            out+=text[s+4+j]
        output.append(out)
        start=e

    count+=1
    # writing to results.xls
    for l in range(18):
        sheet1.write(1+count, 1+l, output[l])
        
        
    print(output)

#Saving the workbook
wb.save('result.xls')

print('workbook saved')
