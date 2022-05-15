
# werge

### Overview
werge is a python library that allows for the centralization and automation of Microsoft Word (2007+) Mail Merge proccesses. With werge you can convert existing Microsoft Word documents into a JSON structure that can be combined with custom content to produce letters from a database, fileshare, or other systems without having to worry about addtional licensing or word compatability. 

There are two primary modules within werge:

#### docxParser:
Responsible for parsing the word document and returning a json representation of the word file. Because of the discrepencies in the Word and PDF standards, the parsing is not 100% accurate and never will be. Font and styles in general are often not able to be accurately parsed due differences in the PDF and MS Word structures*.  

##### Currently Supports
* Tables
* Paragraphs
* Listed Paragraphs/Bullets
* Images/Graphics
* MergeFields
* Standard PDF Fonts

\*This project has given me a new appreciation for just how important the PDF standard is/was. 

#### pdfLetter
Responsible for combining the json template and merge_field content to produce pdf file(s), the produced files can be stored locally or stored in a database. 
##### Currently Supports
* Headers/Footers
* Decking (Combining multiple letters into a single PDF)
* Multiple Inline styles
* Tables
* Static/Dynamic Images
* Bulleted Lists
##### Pending
* HTML Output
* custom table styles
* Non-Standard PDF Fonts

#### To Install
Download the zip and extract it then run the setup.py


#### Example
The example folder contains a example letter, json file and test data. The example uses a CSV file but any pandas DataFrame can be used to populate the mailMerge Fields, including DataFrames produced from SQL queries. 
 

#### JSON Structure
Essential to werge's functionality is the custom json structure representing the MS word file. 
Below is the default_structure found in the config folder, used as the basis for the word file converstion. 

It's the basis for defining the layout of the resulting PDF, with full support for ReportLab and PDF text styles.

You can see an example json structure and letter in the example folder.

```
{
    "page_style":
       {
       "template":"SimpleDocTemplate",
       "default_name":"name_here.pdf",
       "style_sheet":"getSampleStyleSheet",
       "type":"Normal",
       "rightMargin":0.50,
       "leftMargin":0.50,
       "topMargin":0.5,
       "bottomMargin":0.5
       },
    "data_map":[],
    "content_keys":[],
   "default_spacer":[1.1,10],
   "table_styles":[],
   "text_styles":[
                   {"name":"Justify",
                   "alignment":"TA_JUSTIFY",
                   "fontName":"Times",
                   "fontSize":9
                   },
                   {
                   "name": "Footer",
                   "alignment": "TA_CENTER",
                   "fontName": "Times",
                   "fontSize":9
               },
               {
                   "name": "Header",
                   "alignment": "TA_RIGHT",
                   "fontName": "Times",
                   "fontSize":9
               }
   ],  
   "pages":
           {"header":[],
           "footer":[],
           "tables":[],
           "body":[
                   {
                   "type": "Spacer",
                   "spacer": [1.1,50]
                   }
           ]
   }


```

### Requirements:
* pandas
* PyPdf2
* python-docx
* reportlab
 
 
 
 ### Example
 The example.py is a small example app that shows functionality of both docxParser and pdfLetter,for simplicity a CSV file is used to populate the data fields rather than a centralized database like you'd expect in a production implementation of werge. 
