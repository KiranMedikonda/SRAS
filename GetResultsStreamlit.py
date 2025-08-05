import streamlit as st
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import xlrd
import csv
import time
import os
from io import BytesIO

st.set_page_config(page_title="Results Scraper", layout="wide")

st.title("📊 University Results Scraper")
st.markdown("Upload the student registration file and provide the results URL.")

# Input Fields
result_url = st.text_input("Enter Result URL")
uploaded_file = st.file_uploader("Upload Excel File (XLSX)", type=["xls", "xlsx"])

# Chrome WebDriver Path
path_to_chromedriver = r'C:\Users\kiran\Desktop\chromedriver_win32\chromedriver'

def convert_df_to_csv(df):
    return df.to_csv(index=False).encode('utf-8')

# Run Scraper Button
if st.button("Start Scraping"):
    if result_url and uploaded_file:
        try:
            with st.spinner("Initializing browser and reading Excel file..."):
                # Read Excel
                workbook = xlrd.open_workbook(file_contents=uploaded_file.read())
                worksheet = workbook.sheet_by_name("Sheet1")
                num_rows = worksheet.nrows
                num_cols = worksheet.ncols
                
                # Setup Selenium
                browser = webdriver.Chrome(executable_path=path_to_chromedriver)
                browser.get(result_url)
                time.sleep(2)

                data1 = []
                data3 = []

                for curr_col in range(0, num_cols):
                    for curr_row in range(1, num_rows):
                        reg_num = worksheet.cell_value(curr_row, curr_col)
                        st.write(f"Processing Reg Num: {reg_num}")
                        try:
                            browser.find_element(By.CSS_SELECTOR, 'input[id="ht"]').send_keys(reg_num)
                            browser.find_element(By.CSS_SELECTOR, "input[type='button']").click()
                            time.sleep(1)

                            if curr_row == 1 and curr_col == 0:
                                subject_names = ["Reg_Num"]
                                subjects = browser.find_elements(By.XPATH, '//*[@id="rs"]/table/tbody/tr/td[2]')
                                for subject in subjects:
                                    subject_names.append(subject.text)
                                subject_names += ["SGPA", "Pass/Fail", "Backlogs"]
                                data1.append(subject_names)
                                data3.append(["Reg_Num"] + [credit.text for credit in browser.find_elements(By.XPATH, '//*[@id="rs"]/table/tbody/tr/td[4]')])

                            grades = [reg_num]
                            grade_elements = browser.find_elements(By.XPATH, '//*[@id="rs"]/table/tbody/tr/td[3]')
                            for grade in grade_elements:
                                grades.append(grade.text)

                            if 'F' in grades or 'ABSENT' in grades:
                                grades.append("Fail")
                            else:
                                grades.append("Pass")
                            backlogs = grades.count("F")
                            grades.append(backlogs)

                            gpoints = []
                            grade_map = {'A':8, 'B':7, 'C':6, 'D':4, 'O':10, 'S':9, 'F':0}
                            for g in grades[1:-2]:  # Skip reg_num, Pass/Fail, Backlogs
                                gpoints.append(grade_map.get(g, 0))

                            credit_elements = browser.find_elements(By.XPATH, '//*[@id="rs"]/table/tbody/tr/td[4]')
                            credits = [int(c.text) for c in credit_elements]
                            mult = [a * b for a, b in zip(credits, gpoints)]
                            sgpa = round(sum(mult) / sum(credits), 2) if sum(credits) else 0
                            grades.insert(-2, sgpa)

                            data1.append(grades)
                            browser.find_element(By.CSS_SELECTOR, 'input[id="ht"]').clear()
                        except Exception as e:
                            st.warning(f"Error processing {reg_num}: {e}")
                            browser.find_element(By.CSS_SELECTOR, 'input[id="ht"]').clear()
                            continue

                browser.quit()

            # DataFrames for export
            df_grades = pd.DataFrame(data1[1:], columns=data1[0])
            csv_grades = convert_df_to_csv(df_grades)

            st.success("Scraping complete!")
            st.download_button("Download Grades CSV", csv_grades, "Grades.csv", "text/csv")

            # Analysis
            st.subheader("Grade Analysis")
            df = df_grades
            analysis_data = {"Subjects": [], "A":[], "B":[], "C":[], "D":[], "O":[], "S":[], "F":[], "FailPercentage":[], "PassPercentage":[]}
            for col in df.columns[1:-3]:
                analysis_data["Subjects"].append(col)
                grades = df[col]
                counts = grades.value_counts()
                total = len(grades)
                for g in ["A", "B", "C", "D", "O", "S", "F"]:
                    analysis_data[g].append(counts.get(g, 0))
                f_per = (counts.get("F", 0) / total) * 100
                analysis_data["FailPercentage"].append(round(f_per, 2))
                analysis_data["PassPercentage"].append(round(100 - f_per, 2))

            df_analysis = pd.DataFrame(analysis_data)
            csv_analysis = convert_df_to_csv(df_analysis)

            st.dataframe(df_analysis)
            st.download_button("Download Analysis CSV", csv_analysis, "Analysis.csv", "text/csv")

        except Exception as e:
            st.error(f"An error occurred: {e}")
    else:
        st.warning("Please provide both the result URL and Excel file.")
