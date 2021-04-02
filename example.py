import json
from tkinter import filedialog
from werge.pdfLetter import PdfLetter
  
from werge.docxParser import Parser, image_to_b64
import sys
from os import system
from pandas import DataFrame, read_excel, read_csv, read_sql_query
from collections import OrderedDict


import pyodbc
import re
from time import sleep

import numpy as np

DEBUG = False

from tkinter.filedialog import asksaveasfile


cls = lambda *args: system('cls')
df_col_mapper = lambda cols: {col:re.sub(pattern=r" +",repl="_",string=col )for col in cols}
exit = lambda *args:sys.exit(0)

 

 

class Letter:

 

    def __init__ (self,dataframe,format_file):

        self.dataframe = dataframe
        self.format_file = format_file
        self.field_contents = {}
        

 

    @property
    def get_file_name (self):
        # Add logic to set the filename here. In this case we pull the APPLICATION_ID from
        # the field_contents dictionary appened to the days date.
        return "{}_{}.pdf".format(self.get_date,self.field_contents["application_id"]["APPLICATION_ID"])

   

    @property
    def get_date (self):
        return date.today().strftime("%Y-%m-%d")



    def validate_data (self)->DataFrame:

        """
            rules and logic that can be used to automatically verify the data can be defined here.

             args:
                dataframe (DataFrame): The validated dataframe

        """


        try:
            # Add validation logic here
            pass

        except Exception as e:
            raise e

 
        return self.dataframe

 

    def format_data (self)->DataFrame:

        """
                This function is where you can define custom formating of the data.

                args:
                    dataframe (DataFrame): The formated datafram

        """

        try:
            self.dataframe.replace(np.nan,'',regex=True,inplace=True)

        except Exception as e:
            raise e

 

        return self.dataframe

 

    def trim_data (self):

        """

                This function trims white space surrounding column contents and verifies that the trim was successful


        """

        # Iterate through the columns and call the pandas.Series.strip function, the equivelent to the excel Trim()
        # A Pandas DataFrame is made up of Series, these are either Column Series (As in the case below), or can be a "Row" Series
        for cols in self.dataframe.columns:
            try:
                self.dataframe[cols] = self.dataframe[cols].str.strip()
            except AttributeError:
                pass

        return self.dataframe

       

     

    @classmethod
    def load_data_from_spreadsheet (cls,format_file,*args,**kwargs):

        # read_excel allows to create a dataframe from a specific tab or  if the spreadsheet can be multi-tabbed, return an OrderedDict with the keys being tab names and values being DataFrame objects
        # to only read a specific sheet add the following key-value pair to the below function sheet_name="[SHEET NAME]" or an index value, in this case 0 reads just the first tab.


        print("Load the utf-8 encoded csv now.")#()

        sleep(1)

        df = read_csv(filedialog.askopenfile(filetypes = [("CSV File", "*.csv")]))
        # ensure column names are uppercase with whitespace replaced by "_"
        df = df.rename(df_col_mapper(df.columns),axis=1)
        
        letter = cls(dataframe=df,format_file=format_file)
 

        letter.trim_data()
        letter.validate_data()
        letter.format_data()


        return letter

 

 

    @classmethod
    def load_data_from_sql (cls,format_file,*args,**kwargs):


        conn = pyodbc.connect(driver='{ODBC Driver 17 for SQL Server}',
                  server='SERVER NAME',        
                  trusted_connection='yes')

        cursor = conn.cursor()

        sql_query = """ QUERY HERE """
        df = read_sql_query(sql_query, conn)
        df = df.rename(df_col_mapper(df.columns),axis=1)

        letter = cls(dataframe=df,format_file=format_file)
        letter.trim_data()
        letter.validate_data()

        return letter

 

   

    @staticmethod
    def merge_pdf (letters,temp_directory_prefix="letters_",merged_pdf_loc="C:\\temp",merged_pdf_name="mergedPdf.pdf"):

        merger = PdfFileMerger()
        tempdirecpdf = TemporaryDirectory(prefix=temp_directory_prefix)

        print("Building PDF Now")
        for letter in letters:
            merger.append(letter.save_pdf(dir_location=tempdirecpdf.name))
            
        try:
            f_name = asksaveasfile(filetypes = [("PDF File",".pdf")], defaultextension = [("pdf File",".pdf")]).name 
            merger.write(f"{f_name}")
            merger.close()
        except Exception as e:
            raise e

        tempdirecpdf.cleanup()

 

def word_to_json(*args,**kwargs):

    # docx file, using tkinter filedialog the mode "rb" must be specified or a fileError will be thrown    
    input("Press Any Key to load a docx file to parse.")
    file_loc = filedialog.askopenfile(filetypes = [("word File", "*.docx")] ,mode="rb")
    
    parser = Parser.from_file(file_loc,prompt=True)
    file_loc.close()
    input("Parsing complete, press any key to save the Json template")
    parser.build_json(asksaveasfile(filetypes = [("Json File",".json")], defaultextension = [("Json File",".json")]).name )

 

 

def create_pdf_from_json (field_contents=None,pdf_name=None):

    # generate the pdfs to merge into one long PDF
    # this is required by the mailroom
    generated_pdfs = []

 

    try:

        print("Select the json template to load")

        sleep(0.5)
        
        format_file = json.load(filedialog.askopenfile(filetypes = [("json File", "*.json")]))
       
        if field_contents == None:
            letter = Letter.load_data_from_spreadsheet(format_file)
    
        generated_pdfs = PdfLetter.create_pdf(format_file,letter.dataframe)
        PdfLetter.merge_pdf(generated_pdfs,merged_pdf_name=asksaveasfile(filetypes = [("PDF File",".pdf")], defaultextension = [("pdf File",".pdf")]).name )
        


    except AttributeError as ae:
        print("Unable to locate the template specified, please check the name and try again.")
        if DEBUG == True:
            raise ae

 

def menu(field_contents=None):

   
    menu_items = {"Parse Word File":"word_to_json","Create PDF From JSON":"create_pdf_from_json","Exit":"exit"}


    while True:

        cls()

        print("Make A selection:")
        for enum,item in enumerate(menu_items.keys()):
            print(f"{enum}.{item}")

        selection = list(menu_items.keys())[int(input("Enter a number: ")[0])]
        try:
            getattr(sys.modules[__name__],menu_items[selection])(field_contents)


        except IndexError as ie:
            input("Invalid Entry, please just enter a number corresponding to an option. Press Any Key to try again")

        except Exception as e:
            raise e


if __name__ == '__main__':
    menu()