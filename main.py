import openpyxl as op
import os
import pandas as pd
import shutil
from tkinter import *
from tkinter import filedialog

books_file_path=""
gst_file_path=""


def clear_file_paths():
    global books_file_path
    global gst_file_path

    books_path_label = Label(root,text = ' '*len(books_file_path)*3)
    books_path_label.config(font=("Arial", 12))
    books_path_label.place(x=240,y=50)

    gst_path_label = Label(root,text = ' '*len(gst_file_path)*3)
    gst_path_label.config(font=("Arial", 12))
    gst_path_label.place(x=240,y=100)



def get_books_file_path():
    global books_file_path
    books_file_path = filedialog.askopenfilename(initialdir = "/", title = "Select a File", filetypes = (("XLSX File", "*.xlsx*"),("CSV File", "*.csv*"),("Excel", "*.xls*"),("All files", "*.*"))) 

    books_path_label = Label(root,text = books_file_path)
    books_path_label.config(font=("Arial", 12))
    books_path_label.place(x=240,y=50)

    if os.path.isfile(books_file_path):
        message_desc=Label(root,text="Books File exists in the given path")
        message_desc.config(font=("Arial", 12),foreground="green")
        message_desc.place(x=20,y=250)
        message_desc.after(1000,message_desc.destroy)
    else:
        message_desc=Label(root,text="File does not exist in the given path")
        message_desc.config(font=("Arial", 12),foreground="red")
        message_desc.place(x=20,y=250)
        message_desc.after(1000,message_desc.destroy)

def get_gst_file_path():
    global gst_file_path
    
    gst_file_path = filedialog.askopenfilename(initialdir = "/", title = "Select a File", filetypes = (("XLSX File", "*.xlsx*"),("CSV File", "*.csv*"),("Excel", "*.xls*"),("All files", "*.*"))) 
    gst_path_label = Label(root,text = gst_file_path)
    gst_path_label.config(font=("Arial", 12))
    gst_path_label.place(x=240,y=100)

    if os.path.isfile(books_file_path):
        message_desc=Label(root,text="GST File exists in the given path")
        message_desc.config(font=("Arial", 12),foreground="green")
        message_desc.place(x=20,y=250)
        message_desc.after(1000,message_desc.destroy)
    else:
        message_desc=Label(root,text="File does not exist in the given path")
        message_desc.config(font=("Arial", 12),foreground="red")
        message_desc.place(x=20,y=250)
        message_desc.after(1000,message_desc.destroy)


def generate_gst_reco():

    global books_file_path
    global gst_file_path
    
    accounts_df=pd.DataFrame()
    gstr_df=pd.DataFrame()

    accounts_file_path=books_file_path
    gstr_file_path=gst_file_path

    if os.path.exists(accounts_file_path):
        accounts_wb=op.load_workbook(accounts_file_path)
        accounts_df=pd.DataFrame(accounts_wb.active.values)
        accounts_df.columns=accounts_df.iloc[4]
        accounts_df=accounts_df[5:-3]
        mask = accounts_df.applymap(lambda x: x is None)
        cols = accounts_df.columns[(mask).any()]
        for col in accounts_df[cols]:
            accounts_df.loc[mask[col], col] = ' '
        accounts_df["PartyGSTIN"]= accounts_df["PartyGSTIN"].str.strip()
    else:
        print("Accounts file missing")


    if os.path.exists(gstr_file_path):
        gstr_wb=op.load_workbook(gstr_file_path)
        gstr_df=pd.DataFrame(gstr_wb.active.values)
        gstr_df.columns=["GSTIN of supplier","Trade/Legal name","Invoice number","Invoice Date","Invoice Value",
        "Taxable Value","Integrated Tax","Central Tax","State/UT Tax","Rate(%)","GSTR-1/5 Period","GSTR-1/5 Filing Date",
        "ITC Availability","Place of supply","Supply Attract Reverse Charge","Reason","Applicable percentage of Tax Rate","Source",
        "IRN","IRN Date"]
        gstr_df=gstr_df[1:]
        mask = gstr_df.applymap(lambda x: x is None)
        cols = gstr_df.columns[(mask).any()]
        for col in gstr_df[cols]:
            gstr_df.loc[mask[col], col] = ' '
    else:
        print("GSTR file missing")

    company_gst_master=dict()

    for index, row in accounts_df.iterrows():
        company_gst_master[row['PartyGSTIN'].strip()]=row['Party Account']

    cwd=os.getcwd()
    reco_file_two_path=os.path.join(cwd,"temp","gst_reco.csv")

    pd.DataFrame({}).to_csv(reco_file_two_path)
    
    for key,value in company_gst_master.items():
        company_df=pd.DataFrame([key,value]).transpose()
        account_temp_df=accounts_df.loc[accounts_df['PartyGSTIN'] == key]
        acc_df=pd.DataFrame(columns=["Source","Date","Narration","Taxable Value"])
        for index,row in account_temp_df.iterrows():       
            temp_df = {"Source": "Books", 'Date': row["Date"],"Narration":row["Supplier Bill No"].strip(),"Taxable Value":round(float(row["Taxable Value"]),2)}
            acc_df=acc_df.append(temp_df,ignore_index=True)

        gstr_temp_df=gstr_df.loc[gstr_df['GSTIN of supplier']==key]
        gst_df=pd.DataFrame(columns=["GST Source","Invoice Date","Invoice number","Taxable Value"])
        for index,row in gstr_temp_df.iterrows():
            row["Invoice number"]=row["Invoice number"].strip()
            temp_df={"GST Source":"GSTR2B Input","Invoice Date":row["Invoice Date"],"Invoice number":row["Invoice number"],"Taxable Value":round(float(row["Taxable Value"]),2)}
            gst_df=gst_df.append(temp_df,ignore_index=True)
        
        acc_gst_merge=pd.merge(acc_df,gst_df,left_on="Narration",right_on="Invoice number",how="outer")
        company_df.to_csv(reco_file_two_path, mode='a',header=False,index=False)
        acc_gst_merge.to_csv(reco_file_two_path, mode='a',index=False)
    
    os.startfile(reco_file_two_path)


if __name__=="__main__":

    root = Tk()
    root.title("Twills Clothing Pvt. Ltd.")
    root.geometry("550x300")

    header_label_one=Label(root,text="GST Reconciliation Tool",anchor="w")
    header_label_one.config(font=("Arial", 16))
    header_label_one.place(x=10,y=10)

    instruction_button = Button(root, text="Instructions")
    instruction_button.config(font=("Arial", 12))
    instruction_button.place(x=400,y=10)

    books_label=Label(root,text="Books : ",font=("bold",10))
    books_label.config(font=("Arial", 12))
    books_label.place(x=10,y=50)

    books_data_file = Button(root,text = "Choose File",command=get_books_file_path)
    books_data_file.config(font=("Arial", 12))
    books_data_file.place(x=140,y=50)

    books_path_label = Label(root,text = "")
    books_path_label.config(font=("Arial", 12))
    books_path_label.place(x=240,y=50)

    gst_label=Label(root,text="GST : ",font=("bold",10))
    gst_label.config(font=("Arial", 12))
    gst_label.place(x=10,y=100)

    gst_data_file = Button(root,text = "Choose File",command=get_gst_file_path)
    gst_data_file.config(font=("Arial", 12))
    gst_data_file.place(x=140,y=100)

    gst_path_label = Label(root,text = "")
    gst_path_label.config(font=("Arial", 12))
    gst_path_label.place(x=240,y=100)

    button=Button(root,text="Reconcile Data",command=generate_gst_reco)
    button.config(font=("Arial", 12))
    button.place(x=10,y=150)

    button=Button(root,text="Clear",command=clear_file_paths)
    button.config(font=("Arial", 12))
    button.place(x=180,y=150)

    message_label = Label(root,text = "Message :")
    message_label.config(font=("Arial", 12))
    message_label.place(x=10,y=200)

    message_label=Label(root,text="Welcome !!")
    message_label.config(font=("Arial", 12),fg="blue")
    message_label.place(x=20,y=250)
    message_label.after(1000,message_label.destroy)
    
    root.mainloop()

