import pandas as pd
import datetime
from datetime import date
import inflect
import os
from docx import Document
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from openpyxl import load_workbook
import openpyxl
from tkinter import * 
from tkinter import ttk
from tkinter import filedialog
from os import path
from shutil import copyfile
import csv
from openpyxl import Workbook
import shutil
from tkinter.messagebox import showinfo


import_year=""
import_month=""
import_file_path=""
import_data=pd.DataFrame()
months_dict=dict()
months=["January","February","March","April","May","June","July","August","September","October","November","December"]


TITLE="Twills Clothing Pvt. Ltd."
RESOLUTION="700x500"



def delete_import_entries():
    global import_file_path
    payslip_year_start.delete(0,'end')
    payslip_year_end.delete(0,'end')
    monthchoosen.delete(0,'end')
    
    payslip_data_file_label = Label(root,text = ' '*len(import_file_path)*3)
    payslip_data_file_label.config(font=("Arial", 12))
    payslip_data_file_label.place(x=240,y=135)

def getfilepath():
    global import_file_path
    import_file_path = filedialog.askopenfilename(initialdir = "/", title = "Select a File", filetypes = (("Excel", "*.xls*"),("CSV File", "*.csv*"),("All files", "*.*"))) 
    
    payslip_data_file_label = Label(root,text = import_file_path)
    payslip_data_file_label.config(font=("Arial", 12))
    payslip_data_file_label.place(x=240,y=135)
    
    if path.isfile(import_file_path):
        message_desc=Label(root,text="File exists in the given path")
        message_desc.config(font=("Arial", 12),foreground="green")
        message_desc.place(x=30,y=440)
    else:
        message_desc=Label(root,text="File does not exist in the given path")
        message_desc.config(font=("Arial", 12),foreground="red")
        message_desc.place(x=30,y=440)


def import_data():
    global import_year
    global import_month
    global import_file_path

    import_year=payslip_year_start.get()+"-"+payslip_year_end.get()
    import_month=monthchoosen.get().strip()
    
    #Source file 
    if path.isfile(import_file_path):
        wb=load_workbook(import_file_path,data_only=True)
        employee_df=pd.DataFrame(wb["payslip_data_template"].values)
        employee_df.columns=employee_df.iloc[0]
        employee_df=employee_df[1:]
        employee_df["Employee Code"].fillna("N/A",inplace=True)
        employee_df["EPF No"].fillna("N/A",inplace=True)
        employee_df["Employee Name"].fillna("N/A",inplace=True)
        employee_df["Location"].fillna("N/A",inplace=True)
        employee_df["Designation"].fillna("N/A",inplace=True)
        employee_df["Department"].fillna("N/A",inplace=True)
        employee_df["ESI-Y/N"].fillna("N/A",inplace=True)
        employee_df["EPF Code"].fillna("N/A",inplace=True)
        employee_df["PF-Y/N"].fillna("N/A",inplace=True)
        employee_df["PAN No"].fillna("N/A",inplace=True)
        employee_df["ESIC No"].fillna("N/A",inplace=True)
        employee_df["Gross Salary"].fillna("0.00",inplace=True)
        employee_df["Days Worked"].fillna("0.00",inplace=True)
        employee_df["Net Salary"].fillna("0.00",inplace=True)
        employee_df["Basic Salary"].fillna("0.00",inplace=True)
        employee_df["HRA"].fillna("0.00",inplace=True)
        employee_df["Conveyance"].fillna("0.00",inplace=True)
        employee_df["OT Salary"].fillna("0.00",inplace=True)
        employee_df["Incentive"].fillna("0.00",inplace=True)
        employee_df["Total Pay"].fillna("0.00",inplace=True)
        employee_df["ESI"].fillna("0.00",inplace=True)
        employee_df["EPF"].fillna("0.00",inplace=True)
        employee_df["Professional Tax"].fillna("0.00",inplace=True)
        employee_df["TDS"].fillna("0.00",inplace=True)
        employee_df["Total Deductions"].fillna("0.00",inplace=True)
        employee_df["Net Pay"].fillna("0.00",inplace=True)
        employee_df["V Type"].fillna("0.00",inplace=True)
        employee_df["Other Deductions"].fillna("0.00",inplace=True)

        #check if the destination folder is available and create it 
        cwd=os.getcwd()

        folder_path=os.path.join(cwd,"data",import_year)
        if not os.path.isdir(folder_path):
            os.mkdir(folder_path)

        folder_path=os.path.join(folder_path,import_month)
        if not os.path.isdir(folder_path):
            os.mkdir(folder_path)
        
        file_path=os.path.join(folder_path,"salary-slip-data-master-"+import_month+"-"+import_year+".xlsx")
        
        employee_df.to_excel(file_path, sheet_name=import_month+" "+import_year, index = False)
        
        if path.isfile(file_path):
            message_desc=Label(root,text="Data Imported Successfully!!")
            message_desc.config(font=("Arial", 12),foreground="green")   
            message_desc.place(x=30,y=440)
        else:
            message_desc=Label(root,text="Data did not imported successfully, please check the input file")
            message_desc.config(font=("Arial", 12),foreground="red")   
            message_desc.place(x=30,y=440)

def download_template():
    cwd=os.getcwd()
    source_path=os.path.join(cwd,"data","payslip_data_template.xlsx")
    destination_path=os.path.join(os.path.expanduser("~/Desktop"),"payslip_data_template.xlsx")
    #copyfile("D:\\University of Colorado\\Subjects\Projects\\twills_payslip_generator\\data\\payslip_data_template.csv","C:\\Users\\Hitesh\\Desktop\\payslip_data_template.csv")
    copyfile(source_path,destination_path)
    message_desc=Label(root,text="Template Downloaded to Desktop")
    message_desc.config(font=("Arial", 12),foreground="blue")   
    message_desc.place(x=30,y=440)


def get_words_for_number(salary):
    p = inflect.engine()
    text=p.number_to_words(int(salary))
    text_final="Rupees "
    for i in text.split():
        if i !="and":
            text_final=text_final+i.capitalize()+" "
        else:
            text_final=text_final+i+" "

    text_final=text_final+"and Paise Zero Only"

    return text_final


def generate_payslip_for_employee():
    
    document = docx.Document('data/payslip_template_placeholder.docx')
    style = document.styles['Normal']
    font = style.font
    font.name = 'Verdana'
    font.size = Pt(8)
    
    payslip_year=payslip_year_start.get()+"-"+payslip_year_end.get()
    payslip_month=monthchoosen.get().strip()
    employee_id=employee_id_entry.get()
    
    cwd=os.getcwd()
    folder_path=os.path.join(cwd,"data",payslip_year)
    folder_path=os.path.join(folder_path,payslip_month)
    file_path=os.path.join(folder_path,"salary-slip-data-master-"+payslip_month+"-"+payslip_year+".xlsx")
    
    wb=load_workbook(file_path,data_only=True)
    employee_df=pd.DataFrame(wb[payslip_month+" "+payslip_year].values)
    employee_df.columns=employee_df.iloc[0]
    employee_df=employee_df[1:]
    employee_df=employee_df.loc[employee_df['Employee Code'] == employee_id]

    month_header="Salary Slip for the month of "+str(payslip_month)+" "+str(payslip_year)     
    emp_name=employee_df['Employee Name'].values[0]
    emp_code=employee_df['Employee Code'].values[0]
    emp_location=employee_df['Location'].values[0]
    emp_department=employee_df['Department'].values[0]
    emp_designation=employee_df['Designation'].values[0]
    emp_epf_code=employee_df["EPF Code"].values[0]
    emp_pan_number=employee_df["PAN No"].values[0]
    emp_epf_number=employee_df["EPF No"].values[0]
    emp_working_days=26
    emp_esic_number=employee_df["ESIC No"].values[0]
    emp_paid_days=int(employee_df["Days Worked"].values[0])
    emp_leaves_availed=int(emp_working_days-emp_paid_days)
    #Earnings
    sal_basic_pay=format(float(employee_df["Basic Salary"].values[0]),'.2f')
    sal_conveyance=format(float(employee_df["Conveyance"].values[0]),'.2f')
    sal_incentive=format(float(employee_df["Incentive"].values[0]),'.2f')
    sal_hra=format(float(employee_df["HRA"].values[0]),'.2f')
    sal_ot_salary=format(float(employee_df["OT Salary"].values[0]),'.2f')  
    sal_gross_earnings=format(float(employee_df["Total Pay"].values[0]),'.2f')
    #Deductions
    sal_esic=format(float(employee_df["ESI"].values[0]),'.2f')  
    sal_epf=format(float(employee_df["EPF"].values[0]),'.2f')
    sal_pt=format(float(employee_df["Professional Tax"].values[0]),'.2f')   
    sal_tds=format(float(employee_df["TDS"].values[0]),'.2f')
    sal_other_deductions=format(float(employee_df["Other Deductions"].values[0]),'.2f')
    sal_gross_deductions=format(float(employee_df["Total Deductions"].values[0]),'.2f') 
    sal_net_payment=format(float(employee_df["Net Pay"].values[0]),'.2f')  
    sal_net_temp=float(round(employee_df["Net Pay"].values[0]))  
    sal_net_payment_word=get_words_for_number(sal_net_temp) 
    #File Path
    emp_path=str(emp_code)+"_"+str(payslip_month)+"_"+str(payslip_year)+".docx"
    cwd=os.getcwd()
    emp_output=os.path.join(cwd,"temp",emp_path)
    
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                if "<month_header>" in cell.text:
                    cell.text=month_header
                    cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.CENTER
                    cell.paragraphs[0].runs[0].font.bold=True
                elif "emp" in cell.text:
                    if "<emp_name>" in cell.text:
                        cell.text=str(emp_name)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.LEFT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<emp_code>" in cell.text:
                        cell.text=str(emp_code)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<emp_location>" in cell.text:
                        cell.text=str(emp_location)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.LEFT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<emp_designation>" in cell.text:
                        cell.text=str(emp_designation)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.LEFT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<emp_pan_number>" in cell.text:
                        cell.text=str(emp_pan_number)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.LEFT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<emp_working_days>" in cell.text:
                        cell.text=str(emp_working_days)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.LEFT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<emp_leaves_availed>" in cell.text:
                        cell.text=str(emp_leaves_availed)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.LEFT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<emp_department>" in cell.text:
                        cell.text=str(emp_department)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<emp_epf_code>" in cell.text:
                        cell.text=str(emp_epf_code)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<emp_epf_number>" in cell.text:
                        cell.text=str(emp_epf_number)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<emp_esic_number>" in cell.text:
                        cell.text=str(emp_esic_number)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<emp_paid_days>" in cell.text:
                        cell.text=str(emp_paid_days)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                        cell.paragraphs[0].runs[0].font.bold=True
                elif "sal" in cell.text:
                    if "<sal_basic_pay>" in cell.text:
                        cell.text=sal_basic_pay
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<sal_conveyance>" in cell.text:
                        cell.text=str(sal_conveyance)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<sal_hra>" in cell.text:
                        cell.text=str(sal_hra)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<sal_ot_salary>" in cell.text:
                        cell.text=str(sal_ot_salary)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<sal_incentive>" in cell.text:
                        cell.text=str(sal_incentive)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<sal_gross_earnings>" in cell.text:
                        cell.text=str(sal_gross_earnings)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<sal_net_payment_word>" in cell.text:
                        cell.text=str(sal_net_payment_word)
                        cell.paragraphs[0].runs[0].font.bold=True
                        #cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                    elif "<sal_esic>" in cell.text:
                        cell.text=str(sal_esic)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<sal_epf>" in cell.text:
                        cell.text=str(sal_epf)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<sal_pt>" in cell.text:
                        cell.text=str(sal_pt)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<sal_tds>" in cell.text:
                        cell.text=str(sal_tds)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<sal_other_deductions>" in cell.text:
                        cell.text=str(sal_other_deductions)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<sal_gross_deductions>" in cell.text:
                        cell.text=str(sal_gross_deductions)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "<sal_net_payment>" in cell.text:
                        cell.text=str(sal_net_payment)
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                        cell.paragraphs[0].runs[0].font.bold=True

    document.save(emp_output)
    os.startfile(emp_output)
    
    message_desc=Label(root,text="Payslip for the employee "+emp_name+" created")
    message_desc.config(font=("Arial", 12),foreground="green")   
    message_desc.place(x=30,y=440)

def remove_temp_files():
    cwd=os.getcwd()
    temp_folder=os.path.join(cwd,"temp")
    for f in os.listdir(temp_folder):
        os.remove(os.path.join(temp_folder, f))

def generate_form16_for_employee():

    global months_dict
    global months
    
    for m in months:
        months_dict[m]=None

    form16_year=payslip_year_start.get()+"-"+payslip_year_end.get()
    
    cwd=os.getcwd()
    folder_path=os.path.join(cwd,"data",form16_year)

    for m in months:
        file_name="salary-slip-data-master-"+m+"-"+form16_year+".xlsx"
        temp_path=os.path.join(folder_path,m,file_name)
        if not os.path.isfile(temp_path):
            print("File does not exists")
    
    required_columns=['Employee Code','Employee Name','Location','PAN No','Total Pay','Professional Tax','TDS']
    for m in months:
        file_name="salary-slip-data-master-"+m+"-"+form16_year+".xlsx"
        temp_path=os.path.join(folder_path,m,file_name)
        wb=openpyxl.load_workbook(temp_path,data_only=True)
        sheet_name=m+" "+form16_year
        employee_df=pd.DataFrame(wb[sheet_name].values)
        employee_df.columns=employee_df.iloc[0]
        employee_df=employee_df[1:]
        employee_df["TDS"]=employee_df["TDS"].astype(float)
        employee_df["Total Pay"]=employee_df["Total Pay"].astype(float)
        employee_df["Professional Tax"]=employee_df["Professional Tax"]
        employee_df=employee_df[required_columns]
        months_dict[m]=employee_df.loc[employee_df['TDS'] == 0.00]
    
    employee_code=employee_id_entry.get()


    apr_sal=may_sal=jun_sal=0
    jul_sal=aug_sal=sep_sal=0
    oct_sal=nov_sal=dec_sal=0
    jan_sal=feb_sal=mar_sal=0

    apr_prof_tax=may_prof_tax=jun_prof_tax=0
    jul_prof_tax=aug_prof_tax=sep_prof_tax=0
    oct_prof_tax=nov_prof_tax=dec_prof_tax=0
    jan_prof_tax=feb_prof_tax=mar_prof_tax=0

    emp_name=""
    emp_pan=""
    emp_location=""

    if employee_code in list(months_dict["April"]["Employee Code"]):
        apr_df=months_dict["April"].loc[months_dict["April"]['Employee Code'] == employee_code]    
        apr_sal=apr_df["Total Pay"].iloc[0]
        apr_prof_tax=apr_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=apr_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=apr_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=apr_df["Location"].iloc[0]
    else:
        apr_sal=0
        apr_prof_tax=0

    if employee_code in list(months_dict["May"]["Employee Code"]):
        may_df=months_dict["May"].loc[months_dict["May"]['Employee Code'] == employee_code]    
        may_sal=may_df["Total Pay"].iloc[0]
        may_prof_tax=may_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=may_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=may_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=may_df["Location"].iloc[0]
    else:
        may_sal=0
        may_prof_tax=0

    if employee_code in list(months_dict["June"]["Employee Code"]):
        jun_df=months_dict["June"].loc[months_dict["June"]['Employee Code'] == employee_code]    
        jun_sal=jun_df["Total Pay"].iloc[0]
        jun_prof_tax=jun_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=jun_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=jun_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=jun_df["Location"].iloc[0]
    else:
        jun_sal=0
        jun_prof_tax=0

    if employee_code in list(months_dict["July"]["Employee Code"]):
        jul_df=months_dict["July"].loc[months_dict["July"]['Employee Code'] == employee_code]    
        jul_sal=jul_df["Total Pay"].iloc[0]
        jul_prof_tax=jul_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=jul_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=jul_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=jul_df["Location"].iloc[0]
    else:
        jul_sal=0
        jul_prof_tax=0

    if employee_code in list(months_dict["August"]["Employee Code"]):
        aug_df=months_dict["August"].loc[months_dict["August"]['Employee Code'] == employee_code]    
        aug_sal=aug_df["Total Pay"].iloc[0]
        aug_prof_tax=aug_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=aug_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=aug_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=aug_df["Location"].iloc[0]
    else:
        aug_sal=0
        aug_prof_tax=0
    
    if employee_code in list(months_dict["September"]["Employee Code"]):
        sep_df=months_dict["September"].loc[months_dict["September"]['Employee Code'] == employee_code]    
        sep_sal=sep_df["Total Pay"].iloc[0]
        sep_prof_tax=sep_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=sep_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=sep_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=sep_df["Location"].iloc[0]
    else:
        sep_sal=0
        sep_prof_tax=0

    if employee_code in list(months_dict["October"]["Employee Code"]):
        oct_df=months_dict["October"].loc[months_dict["October"]['Employee Code'] == employee_code]    
        oct_sal=oct_df["Total Pay"].iloc[0]
        oct_prof_tax=oct_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=oct_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=oct_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=oct_df["Location"].iloc[0]
    else:
        oct_sal=0
        oct_prof_tax=0

    if employee_code in list(months_dict["November"]["Employee Code"]):
        nov_df=months_dict["November"].loc[months_dict["November"]['Employee Code'] == employee_code]    
        nov_sal=nov_df["Total Pay"].iloc[0]
        nov_prof_tax=nov_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=nov_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=nov_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=nov_df["Location"].iloc[0]
    else:
        nov_sal=0
        nov_prof_tax=0

    if employee_code in list(months_dict["December"]["Employee Code"]):
        dec_df=months_dict["December"].loc[months_dict["December"]['Employee Code'] == employee_code]    
        dec_sal=dec_df["Total Pay"].iloc[0]
        dec_prof_tax=dec_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=dec_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=dec_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=dec_df["Location"].iloc[0]
    else:
        dec_sal=0
        dec_prof_tax=0

    if employee_code in list(months_dict["January"]["Employee Code"]):
        jan_df=months_dict["January"].loc[months_dict["January"]['Employee Code'] == employee_code]    
        jan_sal=jan_df["Total Pay"].iloc[0]
        jan_prof_tax=jan_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=jan_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=jan_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=jan_df["Location"].iloc[0]
    else:
        jan_sal=0
        jan_prof_tax=0
    
    if employee_code in list(months_dict["February"]["Employee Code"]):
        feb_df=months_dict["February"].loc[months_dict["February"]['Employee Code'] == employee_code]    
        feb_sal=feb_df["Total Pay"].iloc[0]
        feb_prof_tax=feb_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=feb_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=feb_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=feb_df["Location"].iloc[0]
    else:
        feb_sal=0
        feb_prof_tax=0

    if employee_code in list(months_dict["March"]["Employee Code"]):
        mar_df=months_dict["March"].loc[months_dict["March"]['Employee Code'] == employee_code]    
        mar_sal=mar_df["Total Pay"].iloc[0]
        mar_prof_tax=mar_df["Professional Tax"].iloc[0]        
        emp_name=mar_df["Employee Name"].iloc[0]
        emp_pan=mar_df["PAN No"].iloc[0]
        emp_location=mar_df["Location"].iloc[0]
    else:
        mar_sal=0
        mar_prof_tax=0

    Q1_salary=apr_sal+may_sal+jun_sal
    Q2_salary=jul_sal+aug_sal+sep_sal
    Q3_salary=oct_sal+nov_sal+dec_sal
    Q4_salary=jan_sal+feb_sal+mar_sal

    prof_tax=jan_prof_tax+feb_prof_tax+mar_prof_tax+apr_prof_tax+may_prof_tax+jun_prof_tax+jul_prof_tax+aug_prof_tax+sep_prof_tax+oct_prof_tax+nov_prof_tax+dec_prof_tax
    cwd=os.getcwd()
    template_path=os.path.join(cwd,"data","form16_template_placeholder.xlsx")
   
    wb=openpyxl.load_workbook(template_path,data_only=True)
    sheet=wb.active

    sheet['G9'].value=emp_name.upper() 
    sheet['G11'].value=emp_location.upper()
    sheet['G13'].value=emp_pan.upper()
    sheet['F19'].value=Q1_salary
    sheet['F20'].value=Q2_salary
    sheet['F21'].value=Q3_salary
    sheet['F22'].value=Q4_salary
    sheet['G99'].value=prof_tax
    
    cwd=os.getcwd()
    output_folder=os.path.join(cwd,"temp")

    form16_year=payslip_year_start.get()+"-"+payslip_year_end.get()  
    file_name=employee_code+"-Form-16-"+form16_year+".xlsx"
    
    output_file=os.path.join(output_folder,file_name)
    wb.save(output_file)
    os.startfile(output_file)
    
    message_desc=Label(root,text="Form-16 for the employee "+emp_name+" created")
    message_desc.config(font=("Arial", 12),foreground="green")   
    message_desc.place(x=30,y=440)

    

def get_form16(employee_code):
    global months_dict
    global months
    
    apr_sal=may_sal=jun_sal=0
    jul_sal=aug_sal=sep_sal=0
    oct_sal=nov_sal=dec_sal=0
    jan_sal=feb_sal=mar_sal=0

    apr_prof_tax=may_prof_tax=jun_prof_tax=0
    jul_prof_tax=aug_prof_tax=sep_prof_tax=0
    oct_prof_tax=nov_prof_tax=dec_prof_tax=0
    jan_prof_tax=feb_prof_tax=mar_prof_tax=0

    emp_name=""
    emp_pan=""
    emp_location=""

    if employee_code in list(months_dict["April"]["Employee Code"]):
        apr_df=months_dict["April"].loc[months_dict["April"]['Employee Code'] == employee_code]    
        apr_sal=apr_df["Total Pay"].iloc[0]
        apr_prof_tax=apr_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=apr_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=apr_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=apr_df["Location"].iloc[0]
    else:
        apr_sal=0
        apr_prof_tax=0

    if employee_code in list(months_dict["May"]["Employee Code"]):
        may_df=months_dict["May"].loc[months_dict["May"]['Employee Code'] == employee_code]    
        may_sal=may_df["Total Pay"].iloc[0]
        may_prof_tax=may_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=may_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=may_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=may_df["Location"].iloc[0]
    else:
        may_sal=0
        may_prof_tax=0

    if employee_code in list(months_dict["June"]["Employee Code"]):
        jun_df=months_dict["June"].loc[months_dict["June"]['Employee Code'] == employee_code]    
        jun_sal=jun_df["Total Pay"].iloc[0]
        jun_prof_tax=jun_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=jun_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=jun_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=jun_df["Location"].iloc[0]
    else:
        jun_sal=0
        jun_prof_tax=0

    if employee_code in list(months_dict["July"]["Employee Code"]):
        jul_df=months_dict["July"].loc[months_dict["July"]['Employee Code'] == employee_code]    
        jul_sal=jul_df["Total Pay"].iloc[0]
        jul_prof_tax=jul_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=jul_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=jul_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=jul_df["Location"].iloc[0]
    else:
        jul_sal=0
        jul_prof_tax=0

    if employee_code in list(months_dict["August"]["Employee Code"]):
        aug_df=months_dict["August"].loc[months_dict["August"]['Employee Code'] == employee_code]    
        aug_sal=aug_df["Total Pay"].iloc[0]
        aug_prof_tax=aug_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=aug_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=aug_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=aug_df["Location"].iloc[0]
    else:
        aug_sal=0
        aug_prof_tax=0
    
    if employee_code in list(months_dict["September"]["Employee Code"]):
        sep_df=months_dict["September"].loc[months_dict["September"]['Employee Code'] == employee_code]    
        sep_sal=sep_df["Total Pay"].iloc[0]
        sep_prof_tax=sep_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=sep_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=sep_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=sep_df["Location"].iloc[0]
    else:
        sep_sal=0
        sep_prof_tax=0

    if employee_code in list(months_dict["October"]["Employee Code"]):
        oct_df=months_dict["October"].loc[months_dict["October"]['Employee Code'] == employee_code]    
        oct_sal=oct_df["Total Pay"].iloc[0]
        oct_prof_tax=oct_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=oct_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=oct_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=oct_df["Location"].iloc[0]
    else:
        oct_sal=0
        oct_prof_tax=0

    if employee_code in list(months_dict["November"]["Employee Code"]):
        nov_df=months_dict["November"].loc[months_dict["November"]['Employee Code'] == employee_code]    
        nov_sal=nov_df["Total Pay"].iloc[0]
        nov_prof_tax=nov_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=nov_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=nov_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=nov_df["Location"].iloc[0]
    else:
        nov_sal=0
        nov_prof_tax=0

    if employee_code in list(months_dict["December"]["Employee Code"]):
        dec_df=months_dict["December"].loc[months_dict["December"]['Employee Code'] == employee_code]    
        dec_sal=dec_df["Total Pay"].iloc[0]
        dec_prof_tax=dec_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=dec_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=dec_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=dec_df["Location"].iloc[0]
    else:
        dec_sal=0
        dec_prof_tax=0

    if employee_code in list(months_dict["January"]["Employee Code"]):
        jan_df=months_dict["January"].loc[months_dict["January"]['Employee Code'] == employee_code]    
        jan_sal=jan_df["Total Pay"].iloc[0]
        jan_prof_tax=jan_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=jan_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=jan_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=jan_df["Location"].iloc[0]
    else:
        jan_sal=0
        jan_prof_tax=0
    
    if employee_code in list(months_dict["February"]["Employee Code"]):
        feb_df=months_dict["February"].loc[months_dict["February"]['Employee Code'] == employee_code]    
        feb_sal=feb_df["Total Pay"].iloc[0]
        feb_prof_tax=feb_df["Professional Tax"].iloc[0]
        if emp_name=="":
            emp_name=feb_df["Employee Name"].iloc[0]
        if emp_pan=="":
            emp_pan=feb_df["PAN No"].iloc[0]
        if emp_location=="":
            emp_location=feb_df["Location"].iloc[0]
    else:
        feb_sal=0
        feb_prof_tax=0

    if employee_code in list(months_dict["March"]["Employee Code"]):
        mar_df=months_dict["March"].loc[months_dict["March"]['Employee Code'] == employee_code]    
        mar_sal=mar_df["Total Pay"].iloc[0]
        mar_prof_tax=mar_df["Professional Tax"].iloc[0]        
        emp_name=mar_df["Employee Name"].iloc[0]
        emp_pan=mar_df["PAN No"].iloc[0]
        emp_location=mar_df["Location"].iloc[0]
    else:
        mar_sal=0
        mar_prof_tax=0

    Q1_salary=apr_sal+may_sal+jun_sal
    Q2_salary=jul_sal+aug_sal+sep_sal
    Q3_salary=oct_sal+nov_sal+dec_sal
    Q4_salary=jan_sal+feb_sal+mar_sal

    prof_tax=jan_prof_tax+feb_prof_tax+mar_prof_tax+apr_prof_tax+may_prof_tax+jun_prof_tax+jul_prof_tax+aug_prof_tax+sep_prof_tax+oct_prof_tax+nov_prof_tax+dec_prof_tax
    cwd=os.getcwd()
    template_path=os.path.join(cwd,"data","form16_template_placeholder.xlsx")
   
    wb=openpyxl.load_workbook(template_path,data_only=True)
    sheet=wb.active

    sheet['G9'].value=emp_name.upper() 
    sheet['G11'].value=emp_location.upper()
    sheet['G13'].value=emp_pan.upper()
    sheet['F19'].value=Q1_salary
    sheet['F20'].value=Q2_salary
    sheet['F21'].value=Q3_salary
    sheet['F22'].value=Q4_salary
    sheet['G99'].value=prof_tax
    
    form16_year=payslip_year_start.get()+"-"+payslip_year_end.get()
    folder_name="Form-16"
    output_folder=os.path.join(cwd,"data",form16_year,folder_name)
    file_name=employee_code+"-Form-16-"+form16_year+".xlsx"
    
    output_file=os.path.join(output_folder,file_name)
    wb.save(output_file)

def generate_all_payslips():
    
    #get the input payslip data
    #create each employee payslip in the respective folder 
    cwd=os.getcwd()
    payslip_year=payslip_year_start.get()+"-"+payslip_year_end.get()
    payslip_month=monthchoosen.get().strip()

    document = docx.Document('data/payslip_template_placeholder.docx')
    style = document.styles['Normal']
    font = style.font
    font.name = 'Verdana'
    font.size = Pt(8)
    
    
    folder_path=os.path.join(cwd,"data",payslip_year)
    folder_path=os.path.join(folder_path,payslip_month)

    file_path=os.path.join(folder_path,"salary-slip-data-master-"+payslip_month+"-"+payslip_year+".xlsx")
    
    wb=load_workbook(file_path,data_only=True)

    employee_df=pd.DataFrame(wb[payslip_month+" "+payslip_year].values)
    employee_df.columns=employee_df.iloc[0]
    employee_df=employee_df[1:]
    
    #Data Cleaning, replacing empty cells with N/A for the personal details columns and with 0.00 for the salary details
    employee_df["Employee Code"].fillna("N/A",inplace=True)
    employee_df["EPF No"].fillna("N/A",inplace=True)
    employee_df["Employee Name"].fillna("N/A",inplace=True)
    employee_df["Location"].fillna("N/A",inplace=True)
    employee_df["Designation"].fillna("N/A",inplace=True)
    employee_df["Department"].fillna("N/A",inplace=True)
    employee_df["ESI-Y/N"].fillna("N/A",inplace=True)
    employee_df["EPF Code"].fillna("N/A",inplace=True)
    employee_df["PF-Y/N"].fillna("N/A",inplace=True)
    employee_df["PAN No"].fillna("N/A",inplace=True)
    employee_df["ESIC No"].fillna("N/A",inplace=True)
    employee_df["Gross Salary"].fillna("0.00",inplace=True)
    employee_df["Days Worked"].fillna("0.00",inplace=True)
    employee_df["Net Salary"].fillna("0.00",inplace=True)
    employee_df["Basic Salary"].fillna("0.00",inplace=True)
    employee_df["HRA"].fillna("0.00",inplace=True)
    employee_df["Conveyance"].fillna("0.00",inplace=True)
    employee_df["OT Salary"].fillna("0.00",inplace=True)
    employee_df["Incentive"].fillna("0.00",inplace=True)
    employee_df["Total Pay"].fillna("0.00",inplace=True)
    employee_df["ESI"].fillna("0.00",inplace=True)
    employee_df["EPF"].fillna("0.00",inplace=True)
    employee_df["Professional Tax"].fillna("0.00",inplace=True)
    employee_df["TDS"].fillna("0.00",inplace=True)
    employee_df["Total Deductions"].fillna("0.00",inplace=True)
    employee_df["Net Pay"].fillna("0.00",inplace=True)
    employee_df["V Type"].fillna("0.00",inplace=True)
    employee_df["Other Deductions"].fillna("0.00",inplace=True)
    
    for i in range(0,len(employee_df)):
        document = docx.Document('data/payslip_template_placeholder.docx')
        style = document.styles['Normal']
        font = style.font
        font.name = 'Verdana'
        font.size = Pt(8)
        
        #Header
        month_header="Salary Slip for the month of "+str(payslip_month)+" "+str(payslip_year)   
        emp_name=employee_df['Employee Name'].iloc[i]
        emp_code=employee_df['Employee Code'].iloc[i]
        emp_location=employee_df['Location'].iloc[i]
        emp_department=employee_df['Department'].iloc[i]
        emp_designation=employee_df['Designation'].iloc[i]
        emp_epf_code=employee_df["EPF Code"].iloc[i]
        emp_pan_number=employee_df["PAN No"].iloc[i]
        emp_epf_number=employee_df["EPF No"].iloc[i]
        emp_working_days=26
        emp_esic_number=employee_df["ESIC No"].iloc[i]
        emp_paid_days=int(employee_df["Days Worked"].iloc[i])
        emp_leaves_availed=int(emp_working_days-emp_paid_days)
        #Earnings
        sal_basic_pay=format(float(employee_df["Basic Salary"].iloc[i]),'.2f')
        sal_conveyance=format(float(employee_df["Conveyance"].iloc[i]),'.2f')
        sal_incentive=format(float(employee_df["Incentive"].iloc[i]),'.2f')
        sal_hra=format(float(employee_df["HRA"].iloc[i]),'.2f')
        sal_ot_salary=format(float(employee_df["OT Salary"].iloc[i]),'.2f')
        sal_gross_earnings=format(float(employee_df["Total Pay"].iloc[i]),'.2f')
        #Deductions
        sal_esic=format(float(employee_df["ESI"].iloc[i]),'.2f')
        sal_epf=format(float(employee_df["EPF"].iloc[i]),'.2f')
        sal_pt=format(float(employee_df["Professional Tax"].iloc[i]),'.2f')
        sal_tds=format(float(employee_df["TDS"].iloc[i]),'.2f')
        sal_other_deductions=format(float(employee_df["Other Deductions"].iloc[i]),'.2f')
        sal_gross_deductions=format(float(employee_df["Total Deductions"].iloc[i]),'.2f')
        sal_net_payment=format(float(employee_df["Net Pay"].iloc[i]),'.2f')
        sal_net_temp=float(round(employee_df["Net Pay"].iloc[i]))
        sal_net_payment_word=get_words_for_number(sal_net_temp)
        #File Path
        emp_path=str(emp_code)+"_"+str(payslip_month)+"_"+payslip_year_start.get()+"_"+payslip_year_end.get()+".docx"
        emp_output=os.path.join(folder_path,emp_path)
        
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    if "<month_header>" in cell.text:
                        cell.text=month_header
                        cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.CENTER
                        cell.paragraphs[0].runs[0].font.bold=True
                    elif "emp" in cell.text:
                        if "<emp_name>" in cell.text:
                            cell.text=str(emp_name)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.LEFT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<emp_code>" in cell.text:
                            cell.text=str(emp_code)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<emp_location>" in cell.text:
                            cell.text=str(emp_location)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.LEFT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<emp_designation>" in cell.text:
                            cell.text=str(emp_designation)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.LEFT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<emp_pan_number>" in cell.text:
                            cell.text=str(emp_pan_number)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.LEFT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<emp_working_days>" in cell.text:
                            cell.text=str(emp_working_days)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.LEFT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<emp_leaves_availed>" in cell.text:
                            cell.text=str(emp_leaves_availed)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.LEFT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<emp_department>" in cell.text:
                            cell.text=str(emp_department)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<emp_epf_code>" in cell.text:
                            cell.text=str(emp_epf_code)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<emp_epf_number>" in cell.text:
                            cell.text=str(emp_epf_number)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<emp_esic_number>" in cell.text:
                            cell.text=str(emp_esic_number)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<emp_paid_days>" in cell.text:
                            cell.text=str(emp_paid_days)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                            cell.paragraphs[0].runs[0].font.bold=True
                    elif "sal" in cell.text:
                        if "<sal_basic_pay>" in cell.text:
                            cell.text=sal_basic_pay
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<sal_conveyance>" in cell.text:
                            cell.text=str(sal_conveyance)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<sal_hra>" in cell.text:
                            cell.text=str(sal_hra)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<sal_ot_salary>" in cell.text:
                            cell.text=str(sal_ot_salary)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<sal_incentive>" in cell.text:
                            cell.text=str(sal_incentive)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<sal_gross_earnings>" in cell.text:
                            cell.text=str(sal_gross_earnings)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<sal_net_payment_word>" in cell.text:
                            cell.text=str(sal_net_payment_word)
                            cell.paragraphs[0].runs[0].font.bold=True
                            #cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                        elif "<sal_esic>" in cell.text:
                            cell.text=str(sal_esic)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<sal_epf>" in cell.text:
                            cell.text=str(sal_epf)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<sal_pt>" in cell.text:
                            cell.text=str(sal_pt)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<sal_tds>" in cell.text:
                            cell.text=str(sal_tds)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<sal_other_deductions>" in cell.text:
                            cell.text=str(sal_other_deductions)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<sal_gross_deductions>" in cell.text:
                            cell.text=str(sal_gross_deductions)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                            cell.paragraphs[0].runs[0].font.bold=True
                        elif "<sal_net_payment>" in cell.text:
                            cell.text=str(sal_net_payment)
                            cell.paragraphs[0].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                            cell.paragraphs[0].runs[0].font.bold=True

        document.save(emp_output)

    shutil.make_archive(folder_path, 'zip', folder_path)
    message_desc=Label(root,text="Payslips for the employees created at the path "+folder_path)
    message_desc.config(font=("Arial", 12),foreground="green")   
    message_desc.place(x=30,y=440)





def generate_all_form16():
    global months_dict
    global months

    for m in months:
        months_dict[m]=None

    form16_year=payslip_year_start.get()+"-"+payslip_year_end.get()
    
    cwd=os.getcwd()
    folder_path=os.path.join(cwd,"data",form16_year)

    for m in months:
        file_name="salary-slip-data-master-"+m+"-"+form16_year+".xlsx"
        temp_path=os.path.join(folder_path,m,file_name)
        if not os.path.isfile(temp_path):
            print("File does not exists")
    
    required_columns=['Employee Code','Employee Name','Location','PAN No','Total Pay','Professional Tax','TDS']
    for m in months:
        file_name="salary-slip-data-master-"+m+"-"+form16_year+".xlsx"
        temp_path=os.path.join(folder_path,m,file_name)
        wb=openpyxl.load_workbook(temp_path,data_only=True)
        sheet_name=m+" "+form16_year
        employee_df=pd.DataFrame(wb[sheet_name].values)
        employee_df.columns=employee_df.iloc[0]
        employee_df=employee_df[1:]
        employee_df["TDS"]=employee_df["TDS"].astype(float)
        employee_df["Total Pay"]=employee_df["Total Pay"].astype(float)
        employee_df["Professional Tax"]=employee_df["Professional Tax"]
        employee_df=employee_df[required_columns]
        months_dict[m]=employee_df.loc[employee_df['TDS'] == 0.00]
    
    employee_id_master=list()

    for k,v in months_dict.items():
        employee_id_master=employee_id_master+months_dict[k]["Employee Code"].tolist()

    employee_id_master=set(employee_id_master)

    form16_folder=os.path.join(folder_path,"Form-16")
    
    if os.path.isdir(form16_folder):
        pass
    else:
        os.mkdir(form16_folder)

    for code in employee_id_master:
        get_form16(code)
    
    
    shutil.make_archive(form16_folder, 'zip', form16_folder)
    message_desc=Label(root,text="Form-16 for the employees created at the path "+folder_path)
    message_desc.config(font=("Arial", 12),foreground="green")   
    message_desc.place(x=30,y=440)

def instructions_popup():

    window = Tk()
    window.title("Instructions")
    window.geometry("600x550")

    header_label_one = Label(window, text="Welcome to the Payslip and Form-16 Automation Tool",anchor="w")
    header_label_one.config(font=("Arial", 14))
    header_label_one.place(x=10,y=10)

    header_label_one = Label(window, text="Instructions",anchor="w")
    header_label_one.config(font=("Arial", 12))
    header_label_one.place(x=10,y=50)
    
    header_label_one = Label(window, text=
    "1. To import the data we need to enter Financial year (in 20xx-20xx format), Month and select the Payroll file."+
    " Make sure to have the appropriate template to import, you have to download the template (which will be copied to desktop) "
    +"and update the columns and import the document to avoid errors",anchor="w")
    header_label_one.config(font=("Arial", 10),wraplength=550,justify="left")
    header_label_one.place(x=10,y=80)

    header_label_one = Label(window, text=
    "2. To generate payslips for all the employees, enter the Financial Year (in 20xx-20xx format), select the month"+
    " and click the button \"Generate All Payslips\" button, make sure you have /data/(year)/(month)/salary-slip-data-master-(month)-(year).xlsx File",anchor="w")
    header_label_one.config(font=("Arial", 10),wraplength=550,justify="left")
    header_label_one.place(x=10,y=160)

    header_label_one = Label(window, text=
    "3. To generate payslip for an employee, enter the Financial Year (in 20xx-20xx format), select the month, enter the Employee ID (in TWILLSXXX format)"+
    " and click the button \"Generate Payslip\" button, make sure you have /data/(year)/(month)/salary-slip-data-master-(month)-(year).xlsx File",anchor="w")
    header_label_one.config(font=("Arial", 10),wraplength=550,justify="left")
    header_label_one.place(x=10,y=230)

    header_label_one = Label(window, text=
    "4. To generate Form-16 for all the employees, enter the Financial Year (in 20xx-20xx format) and click the button \"Generate All Form-16s\""+
    " make sure to have the file /data/(year)/(month)/salary-slip-data-master-(month)-(year).xlsx for each month of the Financial year",anchor="w")
    header_label_one.config(font=("Arial", 10),wraplength=550,justify="left")
    header_label_one.place(x=10,y=320)

    header_label_one = Label(window, text=
    "5. To generate Form-16 for an employee, enter the Financial Year (in 20xx-20xx format) and enter the Employee ID (int TWILLSXXX) and click the button \"Generate Form-16\""+
    " make sure to have the file /data/(year)/(month)/salary-slip-data-master-(month)-(year).xlsx for each month of the Financial year",anchor="w")
    header_label_one.config(font=("Arial", 10),wraplength=550,justify="left")
    header_label_one.place(x=10,y=400)


    button_close = Button(window, text="Close", command=window.destroy)
    button_close.config(font=("Arial", 12))
    button_close.place(x=250,y=500)

    window.mainloop()
#--Main starts from here--#

root = Tk()
root.title(TITLE)
root.geometry(RESOLUTION)

cwd=os.getcwd()
temp_path=os.path.join(cwd,"temp")
if os.path.isdir(temp_path):
    pass
else:
    os.mkdir(temp_path)

data_path=os.path.join(cwd,"data")
if os.path.isdir(data_path):
    pass
else:
    os.mkdir(temp_path)


remove_temp_files()

header_label_one=Label(root,text="Payslip and Form-16",anchor="w")
header_label_one.config(font=("Arial", 16))
header_label_one.place(x=10,y=10)

instruction_button = Button(root, text="Instructions",command=instructions_popup)
instruction_button.config(font=("Arial", 12))
instruction_button.place(x=580,y=10)

payslip_year=Label(root,text="Financial Year : ",font=("bold",10))
payslip_year.config(font=("Arial", 12))
payslip_year.place(x=10,y=50)

payslip_year_start=Entry(root,width=4)
payslip_year_start.config(font=("Arial", 12))
payslip_year_start.place(x=140,y=50)

to_label=Label(root,text="-",font=("bold",10))
to_label.config(font=("Arial", 12))
to_label.place(x=190,y=50)

payslip_year_end=Entry(root,width=4)
payslip_year_end.config(font=("Arial", 12))
payslip_year_end.place(x=210,y=50)

payslip_month=Label(root,text="Month : ",font=("bold",10))
payslip_month.config(font=("Arial", 12))
payslip_month.place(x=10,y=90)


n = StringVar() 
monthchoosen = ttk.Combobox(root, width = 27, textvariable = n) 
monthchoosen.config(font=("Arial", 12))
 # Adding combobox drop down list 
monthchoosen['values'] = (' January',  
                            ' February', 
                            ' March', 
                            ' April', 
                            ' May', 
                            ' June', 
                            ' July', 
                            ' August', 
                            ' September', 
                            ' October', 
                            ' November', 
                            ' December') 
    
monthchoosen.place(x=140,y=90) 
monthchoosen.current() 


payslip_data_label=Label(root,text="Payroll Data : ",font=("bold",10))
payslip_data_label.config(font=("Arial", 12))
payslip_data_label.place(x=10,y=135)

payslip_data_file = Button(root,text = "Choose File",command = getfilepath)
payslip_data_file.config(font=("Arial", 12))
payslip_data_file.place(x=140,y=130)

payslip_data_file_label = Label(root,text = "")
payslip_data_file_label.config(font=("Arial", 12))
payslip_data_file_label.place(x=240,y=135)


button=Button(root,text="Import Data",command=import_data)
button.config(font=("Arial", 12))
button.place(x=10,y=190)

button=Button(root,text="Download Template",command=download_template)
button.config(font=("Arial", 12))
button.place(x=130,y=190)

button=Button(root,text="Clear",command=delete_import_entries)
button.config(font=("Arial", 12))
button.place(x=320,y=190)

employee_id_label=Label(root,text="Employee ID : ",font=("bold",10))
employee_id_label.config(font=("Arial", 12))
employee_id_label.place(x=10,y=250)

employee_id_entry=Entry(root)
employee_id_entry.config(font=("Arial", 12))
employee_id_entry.place(x=140,y=250)

button=Button(root,text="Generate Playslip",command=generate_payslip_for_employee)
button.config(font=("Arial", 12))
button.place(x=140,y=290)

button=Button(root,text="Generate Form-16",command=generate_form16_for_employee)
button.config(font=("Arial", 12))
button.place(x=320,y=290)

payslip_all_desc_label=Label(root,text="Click Here to ",font=("bold",10))
payslip_all_desc_label.config(font=("Arial", 12))
payslip_all_desc_label.place(x=10,y=345)

button=Button(root,text="Generate All Playslips",command=generate_all_payslips)
button.config(font=("Arial", 12))
button.place(x=140,y=340)

button=Button(root,text="Generate All Form-16's",command=generate_all_form16)
button.config(font=("Arial", 12))
button.place(x=355,y=340)

message_label=Label(root,text="Message :",font=("bold",10))
message_label.config(font=("Arial", 12))
message_label.place(x=10,y=400)

message_desc=Label(root,text="Welcome !!")
message_desc.config(font=("Arial", 12),foreground="blue")
message_desc.place(x=30,y=440)


root.mainloop()
    
