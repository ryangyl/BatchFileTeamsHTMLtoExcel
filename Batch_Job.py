# Basic Libraries
import numpy as np
import pandas as pd

# Saving to network drive
import datetime
import os 
import win32com.client

# Comments: '# checks' are used to check the output of the codes, and it will be helpful when debugging
# Otherwise, they provide no use in the actual output of the code and you can make it a comment by placing a '#' in the very beginning of the code line

today = datetime.date.today()
# print("Today's Date: ", today) #  checks

# Initiating outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items

for message in messages:
    # not all items in your Inbox are MailItem objects - you can also have MeetingItem and ReportItem objects. 
    # we only want MailItem objects, so we need to check that Class property = 43 (OlObjectClass.olMail)
    # filtering only MailItem objects 
    if message.Class == 43:
        # filtering out messages based on the message subject and whether it was sent today 
        # we want the ones sent today because email is automatically sent daily, so we want the newest one
        # you can change the message subject accordingly
        if ("Job" in message.Subject) and (message.SentOn.date() == today):
            # print("Message Subject: ", message.Subject) # checks
            # getting the attachments
            attachments = message.Attachments
            for attachment in message.Attachments:
                # saving the attachment
                # print("Attachment Name: ", attachment.FileName)
                attachment.SaveAsFile(os.path.join(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job', attachment.FileName))

for filename in os.listdir(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job'):
    if ".htm" in filename: # checking if file is a html file
        # print("File is html file") # checks
        if filename == "Job MB52_STK LVL, Step 1.htm": # separating this file because its extremely long
            print("This is Job MB52_STK LVL, Step 1.htm") # checks
            pass
        elif "Job YIFR_ENG SPARE INCOMING_V1, Step 1" in filename:# separating this file because a lot of tables
            print("This is Job YIFR")
            pass
        elif "Job MC48, Step 1" in filename: # separating this file because its extremely long
            print("This is Job MC48, Step 1.htm") # checks
            pass
        elif "Job MC50 (2), Step 1" in filename: # separating this file because its extremely long
            print("This is Job MC50 (2), Step 1.htm") # checks
            pass
        elif "Job ROP_MB51_365DAY, Step 1" in filename: # separating this file because its extremely long
            print("This is Job ROP_MB51_365DAY, Step 1.htm") # checks
            pass
        elif "Job ROP_ZMMM_RPT_EXTRACT, Step 1" in filename: # separating this file because its extremely long
            print("This is Job ROP_ZMMM_RPT_EXTRACT, Step 1.htm") # checks
            pass
        else: # checks
            # checking the file path
            file_path = os.path.join(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job', filename)
            print("File Path: ", file_path) # checks
            # checking the data type of the file_path, i.e. ensuring its string
            print(type(file_path)) #checks
            # checking the file name
            print("File Name: ", filename) # checks
            # getting the name of the future excel file
            excel_name = filename.replace(".htm", ".xlsx")
            print("Excel Name: ", excel_name) # checks
            # getting the file path of the future excel file
            excel_file_path = os.path.join(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job', excel_name)
            print("Excel File Path: ", excel_file_path) # checks
            # reading the html file
            file_html = pd.read_html(file_path)
            # finding out the number of tables
            no_of_tables = len(file_html)
            print("No of tables: ", no_of_tables) # checks
            for i in range(no_of_tables):
                # converting the tables into dataframes
                file_df = pd.DataFrame(file_html[i])
                # replacing headers from numbers to the actual headers
                file_df = file_df.rename(columns=file_df.iloc[0])
                # dropping extra row with actual headers name
                file_df = file_df.drop(file_df.index[0])
                # display(file_df) # checks
                # getting the sheet name as want to save different tables in different excel sheets in the same workbook
                excel_sheet_name = "Table " + str(i+1)
                print("Sheet Name: ", excel_sheet_name) # checks
                # creating the excel workbook
                if i == 0: # for the first table we are creating new excel workbook
                    file_df.to_excel(excel_file_path, sheet_name=excel_sheet_name, index=False)
                else: # for subsequent tables we are appending table to existing excel sheet
                    with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
                        file_df.to_excel(writer, sheet_name=excel_sheet_name, index=False)
    else: # file is not a html file
        print("File is not a html file") # checks
        pass

for filename in os.listdir(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job'):
    if ".htm" in filename:
        print("File is a html file") # checks
        if "Job YIFR_ENG SPARE INCOMING_V1, Step 1" in filename: # separating this file because a lot of tables
            # checking the file path
            file_path = os.path.join(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job', filename)
            print("File Path: ", file_path) # checks
            # checking the data type of the file_path, i.e. ensuring its a string
            print(type(file_path)) # checks
            # checking the file name
            print("File Name: ", filename) # checks
            # getting the name of the future excel file
            excel_name = "Job YIFR_ENG SPARE INCOMING_V1, Step 1.xlsx"
            print("Excel Name: ", excel_name) # checks
            # getting the file path of the future excel file
            excel_file_path = os.path.join(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job', excel_name)
            print("Excel File Path: ", excel_file_path) # checks
            # reading the html file
            file_html = pd.read_html(file_path)
            # finding out the number of tables
            no_of_tables = len(file_html)
            print("No. of tables: ", no_of_tables) # checks
            counter = 0
            for i in range(no_of_tables):
                file_df = pd.DataFrame(file_html[i])
                no_of_columns = file_df.shape[1] # getting the number of columns in the table
                print("No of Columns: ", no_of_columns) # checks
                if no_of_columns <= 5:
                    print("First few tables that don't contain incoming po data") # checks
                    counter += 1
                else: # the rest of the tables
                    # replacing headers from numbers to the actual headers
                    file_df = file_df.rename(columns=file_df.iloc[0])
                    # dropping extra row with actual headers name
                    file_df = file_df.drop(file_df.index[0])
                    # combining the separate dataframes into one giant dataframe as they share the same headers
                    if i == counter: # first official table containing incoming spares data
                        combined_df = pd.DataFrame(file_df)
                    else: # for subsequent tables
                        combined_df = pd.concat([combined_df,file_df], axis=0)
            df_columns_list = combined_df.columns.tolist()
            print("Column Names: ", df_columns_list) # checks
            # creating excel workbook
            if "Open\xa0Quantity" in df_columns_list:
                print("Batch Report contains Open Quantity column") # checks
                combined_df = combined_df[["Material", "Short\xa0Text", "Deliv.Date", "Open\xa0Quantity"]]
                combined_df = combined_df.rename(columns = {"Short\xa0Text":"Short Text", "Open\xa0Quantity":"Still to be delivered (qty)", "Deliv.Date":"Delivery date"})
                combined_df.to_excel(excel_file_path, index=False)
                break
            elif "Order\xa0Quantity" in df_columns_list:
                print("Batch Report contains Order Quantity column") # checks
                combined_df = combined_df[["Material", "Short\xa0Text", "Deliv.date", "Order\xa0Quantity"]]
                combined_df = combined_df.rename(columns = {"Short\xa0Text":"Short Text", "Order\xa0Quantity":"Still to be delivered (qty)", "Deliv.Date":"Delivery date"})
                combined_df.to_excel(excel_file_path, index=False)
                break
            else: # does not contain open or order quantity
                break
        else: # other html files
            pass
    else: # other files which are not html
        pass

consumables_master_df = pd.read_excel(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Consumables BoM_2023\Consolidated BoM for All VS_Ver1.xlsx', sheet_name="Consumables Master List")
consumables_master_list = consumables_master_df["Stock Code"].tolist()
# display(consumables_master_df) # checks
# print("Consumables Master List: ") # checks
# print(consumables_master_list) # checks
# print("Length of consumables master list: ", len(consumables_master_list)) # checks

# filtering out the incoming pos of the relevant consumables
for filename in os.listdir(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job'):
    if ".xlsx" in filename:
        # print("File is an excel file") # checks
        if filename == "Job YIFR_ENG SPARE INCOMING_V1, Step 1.xlsx": # separating this file because a lot of tables
            file_path = os.path.join(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job', filename)
            yifr_df = pd.read_excel(file_path)
            # display(yifr_df) # checks
            no_of_incoming_pos = len(yifr_df)
            # print("No of incoming pos: ", no_of_incoming_pos) # checks
            for i in range(no_of_incoming_pos):
                stockcode = yifr_df.loc[i, "Material"]
                # print("Stock Code: ", stockcode) # checks
                if stockcode in consumables_master_list: # stock code is listed in the consumables master list
                    # print("Stock code needed for campaigns") # checks
                    # print() # checks
                    counter += 1 
                else: # stock code is not listed in the consumables master list
                    # print("Stock code not needed for campaigns") # checks
                    # print() # checks
                    yifr_df = yifr_df.drop([i])
            # print("Counter: ", counter) # checks
            # print("Length of new df: ", len(yifr_df))
            # changing the data types into the correct ones
            yifr_df["Material"] = yifr_df["Material"].astype(str)
            yifr_df["Still to be delivered (qty)"] = yifr_df["Still to be delivered (qty)"].astype(int)
            yifr_df["Delivery date"] = pd.to_datetime(yifr_df["Delivery date"],format="%d.%m.%Y")
            # saving into excel file
            # display(yifr_df) # checks
            # ifr_df.info() # checks
            yifr_df.to_excel(file_path, index=False)

for filename in os.listdir(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job'):
    if ".htm" in filename:
        # print("File is a html file") # checks
        if filename == "Job MB52_STK LVL, Step 1.htm": # separating this file because it is really long
            # checking the file path
            file_path = os.path.join(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job', filename)
            # print("File Path: ", file_path) # checks
            # checking the data type of the file_path, i.e. ensuring its string
            # print(type(file_path)) #checks
            # checking the file name
            # print("File Name: ", filename) # checks
            # getting the name of the future excel file
            excel_name = filename.replace(".htm", ".xlsx")
            # print("Excel Name: ", excel_name) # checks
            # getting the file path of the future excel file
            excel_file_path = os.path.join(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job', excel_name)
            # print("Excel File Path: ", excel_file_path) # checks
            # reading the html file
            file_html = pd.read_html(file_path)
            # finding out the number of tables
            no_of_tables = len(file_html)
            # print("No of tables: ", no_of_tables) # checks
            for i in range(no_of_tables):
                # converting the tables into dataframes
                file_df = pd.DataFrame(file_html[i])
                # replacing headers from numbers to the actual headers
                file_df = file_df.rename(columns=file_df.iloc[0])
                # dropping extra row with actual headers name
                file_df = file_df.drop(file_df.index[0])
                # display(file_df) # checks
                # combining the separate dataframes into one giant dataframe as they share the same headers
                if i == 0: # first table
                    combined_df = pd.DataFrame(file_df)
                else: # for subsequent tables
                    combined_df = pd.concat([combined_df,file_df], axis=0)
            # display(combined_df) # checks
            # creating excel workbook
            combined_df.to_excel(excel_file_path, index=False)
        else: # other html files
            pass
    else: # other files which are not html
        pass

for filename in os.listdir(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job'):
    if ".htm" in filename:
        # print("File is a html file") # checks
        if filename == "Job MC48, Step 1.htm": # separating this file because it is really long
            # checking the file path
            file_path = os.path.join(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job', filename)
            # print("File Path: ", file_path) # checks
            # checking the data type of the file_path, i.e. ensuring its string
            # print(type(file_path)) #checks
            # checking the file name
            # print("File Name: ", filename) # checks
            # getting the name of the future excel file
            excel_name = filename.replace(".htm", ".xlsx")
            # print("Excel Name: ", excel_name) # checks
            # getting the file path of the future excel file
            excel_file_path = os.path.join(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job', excel_name)
            # print("Excel File Path: ", excel_file_path) # checks
            # reading the html file
            file_html = pd.read_html(file_path)
            # finding out the number of tables
            no_of_tables = len(file_html)
            # print("No of tables: ", no_of_tables) # checks
            for i in range(no_of_tables):
                # converting the tables into dataframes
                file_df = pd.DataFrame(file_html[i])
                # replacing headers from numbers to the actual headers
                # the actual headers are the first three rows
                file_df.columns = pd.MultiIndex.from_arrays([file_df.iloc[0], file_df.iloc[1], file_df.iloc[2]])
                # dropping the extra rows with the column headers
                file_df = file_df.drop([0,1,2], axis=0)
                # display(file_df) # checks
                # combining the separate dataframes into one giant dataframe as they share the same headers
                if i == 0: # first table
                    combined_df = pd.DataFrame(file_df)
                else: # for subsequent tables
                    combined_df = pd.concat([combined_df,file_df], axis=0)
            # display(combined_df) # checks
            # creating excel workbook
            combined_df.to_excel(excel_file_path)
        else: # other html files
            pass
    else: # other files which are not html
        pass

for filename in os.listdir(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job'):
    if ".htm" in filename:
        # print("File is a html file") # checks
        if filename == "Job MC50 (2), Step 1.htm": # separating this file because it is really long
            # checking the file path
            file_path = os.path.join(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job', filename)
            # print("File Path: ", file_path) # checks
            # checking the data type of the file_path, i.e. ensuring its string
            # print(type(file_path)) #checks
            # checking the file name
            # print("File Name: ", filename) # checks
            # getting the name of the future excel file
            excel_name = filename.replace(".htm", ".xlsx")
            # print("Excel Name: ", excel_name) # checks
            # getting the file path of the future excel file
            excel_file_path = os.path.join(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job', excel_name)
            # print("Excel File Path: ", excel_file_path) # checks
            # reading the html file
            file_html = pd.read_html(file_path)
            # finding out the number of tables
            no_of_tables = len(file_html)
            # print("No of tables: ", no_of_tables) # checks
            for i in range(no_of_tables):
                # converting the tables into dataframes
                file_df = pd.DataFrame(file_html[i])
                # replacing headers from numbers to the actual headers
                # the actual headers are the first three rows
                file_df.columns = pd.MultiIndex.from_arrays([file_df.iloc[0], file_df.iloc[1], file_df.iloc[2]])
                # dropping the extra rows with the column headers
                file_df = file_df.drop([0,1,2], axis=0)
                # display(file_df) # checks
                # combining the separate dataframes into one giant dataframe as they share the same headers
                if i == 0: # first table
                    combined_df = pd.DataFrame(file_df)
                else: # for subsequent tables
                    combined_df = pd.concat([combined_df,file_df], axis=0)
            # display(combined_df) # checks
            # creating excel workbook
            combined_df.to_excel(excel_file_path)
        else: # other html files
            pass
    else: # other files which are not html
        pass

for filename in os.listdir(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job'):
    if ".htm" in filename:
        # print("File is a html file") # checks
        if filename == "Job ROP_MB51_365DAY, Step 1.htm": # separating this file because it is really long
            # checking the file path
            file_path = os.path.join(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job', filename)
            # print("File Path: ", file_path) # checks
            # checking the data type of the file_path, i.e. ensuring its string
            # print(type(file_path)) #checks
            # checking the file name
            # print("File Name: ", filename) # checks
            # getting the name of the future excel file
            excel_name = filename.replace(".htm", ".xlsx")
            # print("Excel Name: ", excel_name) # checks
            # getting the file path of the future excel file
            excel_file_path = os.path.join(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job', excel_name)
            # print("Excel File Path: ", excel_file_path) # checks
            # reading the html file
            file_html = pd.read_html(file_path)
            # finding out the number of tables
            no_of_tables = len(file_html)
            # print("No. of tables: ", no_of_tables) # checks
            for i in range(no_of_tables):
                file_df = pd.DataFrame(file_html[i])
                if i==0:
                    # print("First few tables that don't contain the actual data ")
                    file_df = file_df.rename(columns=file_df.iloc[0])
                    # dropping extra row with actual headers name
                    file_df = file_df.drop(file_df.index[0])
                    # display(file_df) # checks
                    # getting the sheet name as want to save different tables in different excel sheets in the same workbook
                    excel_sheet_name = "Table 1"
                    # print("Sheet Name: ", excel_sheet_name) # checks
                    # creating the excel workbook
                    # for the first table we are creating new excel workbook
                    file_df.to_excel(excel_file_path, sheet_name=excel_sheet_name, index=False)
                else: # The rest of the tables
                    # replacing headers from numbers to the actual headers
                    file_df = file_df.rename(columns=file_df.iloc[0])
                    # dropping extra row with actual headers name
                    file_df = file_df.drop(file_df.index[0])
                    # display(file_df) # checks
                    # combining the separate dataframes into one giant dataframe as they share the same headers
                    if i == 1: # first official table containing incoming spares data
                        combined_df = pd.DataFrame(file_df)
                    else: # for subsequent tables
                        combined_df = pd.concat([combined_df,file_df], axis=0)
            # display(combined_df) # checks
            # appending to existing excel workbook
            excel_sheet_name2 = "Table 2"
            with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
                combined_df.to_excel(writer, sheet_name=excel_sheet_name2, index=False)
        else: # other html files
            pass
    else: # other files which are not html
        pass

for filename in os.listdir(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job'):
    if ".htm" in filename:
        # print("File is a html file") # checks
        if filename == "Job ROP_ZMMM_RPT_EXTRACT, Step 1.htm": # separating this file because it is really long
            # checking the file path
            file_path = os.path.join(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job', filename)
            # print("File Path: ", file_path) # checks
            # checking the data type of the file_path, i.e. ensuring its string
            # print(type(file_path)) #checks
            # checking the file name
            # print("File Name: ", filename) # checks
            # getting the name of the future excel file
            excel_name = filename.replace(".htm", ".xlsx")
            # print("Excel Name: ", excel_name) # checks
            # getting the file path of the future excel file
            excel_file_path = os.path.join(r'\\siwdsntv002\SG_PSC_SG1_PL_04_Public_Shared\Consumables Forecast 2023\Batch Job', excel_name)
            # print("Excel File Path: ", excel_file_path) # checks
            # reading the html file
            file_html = pd.read_html(file_path)
            # finding out the number of tables
            no_of_tables = len(file_html)
            # print("No. of tables: ", no_of_tables) # checks
            counter = 0
            for i in range(no_of_tables):
                file_df = pd.DataFrame(file_html[i])
                no_of_columns = file_df.shape[1] # getting the no of columns
                # print("No. of columns: ", no_of_columns) # checks
                if no_of_columns <= 5:
                    # print("First few tables that don't contain the actual data ")
                    file_df = file_df.rename(columns=file_df.iloc[0])
                    # dropping extra row with actual headers name
                    file_df = file_df.drop(file_df.index[0])
                    # display(file_df) # checks
                    # getting the sheet name as want to save different tables in different excel sheets in the same workbook
                    excel_sheet_name = "Table " + str(i+1)
                    counter += 1
                    # print("Sheet Name: ", excel_sheet_name) # checks
                    # creating the excel workbook
                    if i == 0: # for the first table we are creating new excel workbook
                        file_df.to_excel(excel_file_path, sheet_name=excel_sheet_name, index=False)
                    else: # for the subsequent tables, we are apppending it to the same workbook but different sheets
                        with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
                            file_df.to_excel(writer, sheet_name=excel_sheet_name, index=False)
                else: # The rest of the tables
                    # replacing headers from numbers to the actual headers
                    file_df = file_df.rename(columns=file_df.iloc[0])
                    # dropping extra row with actual headers name
                    file_df = file_df.drop(file_df.index[0])
                    # display(file_df) # checks
                    # combining the separate dataframes into one giant dataframe as they share the same headers
                    if i == counter: # first official table containing incoming spares data
                        combined_df = pd.DataFrame(file_df)
                    else: # for subsequent tables
                        combined_df = pd.concat([combined_df,file_df], axis=0)
            # display(combined_df) # checks
            # appending to existing excel workbook
            excel_sheet_name1 = "Table " + str(counter+1)
            with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
                combined_df.to_excel(writer, sheet_name=excel_sheet_name1, index=False)
        else: # other html files
            pass
    else: # other files which are not html
        pass