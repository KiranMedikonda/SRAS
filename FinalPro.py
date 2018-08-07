#Packages------------
from tkinter import *
from PIL import Image, ImageDraw, ImageTk, ImageFont
from selenium import webdriver
import time
import xlrd
import pandas as pd
import csv
import os
#FrontView------------
root = Tk()
root.geometry("1840x840+10+10")
image = Image.open(r'C:\Users\kiran\Desktop\AllDepts\process.jpg')
image = image.resize((840, 840), Image.ANTIALIAS) 
width, height = image.size
root.resizable(width=True, height=True)
root.geometry("%sx%s"%(width, height))
draw = ImageDraw.Draw(image)
text_x = 175
text_y = 175
photoimage = ImageTk.PhotoImage(image)
Label(root, image=photoimage).place(x=0,y=0)
root.title("Automation System")
labelk=Label(root,text='RESULTS SCRAPPER',font=("Times New Roman",20,"bold"),fg="blue").place(x=250,y=100)
label1=Label(root,text='Enter Result Url',font=("Times New Roman",18,"bold"),fg="black").place(x=75,y=200)
label2=Label(root,text='Enter Input File Path',font=("Times New Roman",18,"bold"),fg="black").place(x=10,y=300)
url=StringVar()
InputFileName=StringVar()
entry_box=Entry(root,textvariable=url,width=55).place(x=280,y=210)
entry_box1=Entry(root,textvariable=InputFileName,width=55).place(x=280,y=300)
#ChromeDriverPath--------------
path_to_chromedriver = r'C:\Users\kiran\Desktop\chromedriver_win32\chromedriver'
browser = webdriver.Chrome(executable_path = path_to_chromedriver)
#url = r'http://jntukresults.edu.in/view-results-56735897.html'
def Browser():
    browser.get(url.get())
    time.sleep(2)
    while(True):
        #ReadXlsxFile---------
        workbook = xlrd.open_workbook(InputFileName.get())
        worksheet = workbook.sheet_by_name("Sheet1")
        num_rows = worksheet.nrows
        count=num_rows-1
        num_cols = worksheet.ncols
        data1=[]
        data3=[]
        for curr_col in range(0, num_cols):
            count=0
            for curr_row in range(1, num_rows):
                reg_num = worksheet.cell_value(curr_row, curr_col)
                print(reg_num)
                #PassRegistrationToTextFieldOfUniversityWebsite---------
                browser.find_element_by_css_selector('input[id="ht"]').send_keys(reg_num)
                #ClickTheButton--------------
                browser.find_element_by_css_selector("input[type='button']").click()
                time.sleep(1)
                #subjects
                if count==0:
                    subject_names=[]
                    subject_names.append("Reg_Num")
                    #GetTheSubjectNamesByUsingXpath-------------------
                    subjects=browser.find_elements_by_xpath('//*[@id="rs"]/table/tbody/tr/td[2]')
                    for subject in subjects:
                        subject_names.append(subject.text)
                    data3.append(subject_names)
                    subject_names.append("SGPA")
                    subject_names.append("Pass/Fail")
                    subject_names.append("Backlogs")
                    data1.append(subject_names)
                #Grades
                Grades=[]
                Grades.append(reg_num)
                #GetTheGradesByUsingXpath--------------------------
                Grade=browser.find_elements_by_xpath('//*[@id="rs"]/table/tbody/tr/td[3]')
                for Gra in Grade:
                    Grades.append(Gra.text)
                if 'F' in Grades:
                    f="Fail"
                    Grades.append(f)
                elif 'ABSENT' in Grades:
                     a="Fail"
                     Grades.append(a)
                else:
                    p="Pass"
                    Grades.append(p)
                w=Grades.count("F")
                Grades.append(w)
                #ConvertGradesToPointsForCalculatingAvg------------
                gpoints=[]
                for g in Grades:
                    if g=='A':
                        gpoints.append(8)
                    elif g=='B':
                        gpoints.append(7)
                    elif g=='C':
                        gpoints.append(6)
                    elif g=='D':
                        gpoints.append(4)
                    elif g=='O':
                        gpoints.append(10)
                    elif g=='S':
                        gpoints.append(9)
                    elif g=='F':
                        gpoints.append(0)   
                #Credits
                Credit_points=[]
                Credit_points.append(reg_num)
                #GetTheCreditPointsByUsingXpath---------------
                Credits=browser.find_elements_by_xpath('//*[@id="rs"]/table/tbody/tr/td[4]')
                for Credit in Credits:
                    Credit_points.append(Credit.text)
                 #Calculating AvgAnd SGPA-------------
                m=Credit_points[1:]
                
                n=[]
                for i in m:
                    n.append(int(i))
                print(n)
                
                multiply=[a*b for a,b in zip(n,gpoints)]
                print(multiply)
                sgpa=(sum(multiply)/sum(n))
                print(sgpa)
                if 'F' in Grades:
                    Grades.insert(-2,0)
                else:
                    Grades.insert(-2,sgpa)
                cgpa=(sum(n)*sgpa)/sum(n)
                print(cgpa)
                browser.find_element_by_css_selector('input[id="ht"]').clear()
                #-----------------------------------#
                data1.append(Grades)
                myFile1 = open('Grades.csv', 'w')
                with myFile1:
                    writer = csv.writer(myFile1)
                    writer.writerows(data1)
                data3.append(Credit_points)
                myFile3 = open('Credit_point.csv', 'w')
                with myFile3:
                    writer = csv.writer(myFile3)
                    writer.writerows(data3)
                count=count+1
        break
    browser.close()
    #Read TheCsv For Analyse The Data-------------------------------------------
    df=pd.read_csv(r'Grades.csv')
    df.to_excel('accounts.xlsx', index=False)
    ExcelFileName= r'accounts.xlsx'
    workbook = xlrd.open_workbook(ExcelFileName)
    worksheet = workbook.sheet_by_name("Sheet1")
    num_rows = worksheet.nrows
    num_cols = worksheet.ncols
    subjects=[]
    a=[]
    b=[]
    c=[]
    d=[]
    o=[]
    s=[]
    f=[]
    fper=[]
    passper=[]
    #---------------------
    h=[]
    h.append(subjects)
    h.append(a)
    h.append(b)
    h.append(c)
    h.append(d)
    h.append(o)
    h.append(s)
    h.append(f)
    h.append(fper)
    h.append(passper)
    for col in range(1,num_cols-3):
            z=[]
            for row in range(0,num_rows):
                r=worksheet.cell_value(row,col)
                z.append(r)
            val=z[0]
            AS=z.count('A')
            a.append(AS)
            BS=z.count('B')
            b.append(BS)
            CS=z.count('C')
            c.append(CS)
            DS=z.count('D')
            d.append(DS)
            OS=z.count('O')
            o.append(OS)
            SS=z.count('S')
            s.append(SS)
            FS=z.count('F')
            f.append(FS)
            FailPercentage=(FS/count)*100
            fper.append(FailPercentage)
            PassPercentage=(100-FailPercentage)
            passper.append(PassPercentage)
            subjects.append(val)
    dfs = pd.DataFrame(h)
    k=dfs.T
    k.columns=['Subjects','A','B','C','D','O','S','F','FailPercentage','PassPercentage']
    k.to_csv('Analysis.csv', index=False, header=True)
    df1=pd.read_csv(r'Grades.csv')
    df1.to_csv('Grade.csv',index=False)
    df3=pd.read_csv(r'Credit_point.csv')
    df3.to_csv('Credit_points.csv',index=False)
    os.remove("accounts.xlsx")
    os.remove("Grades.csv")
    os.remove("Credit_point.csv")
work=Button(root,text="Click Me !",width=45,height=1,bg="lightgreen",font=("Times New Roman",9,"bold"),fg="red",command=Browser).place(x=280,y=400)
root.mainloop()
