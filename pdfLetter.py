from reportlab import platypus
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer,  Image, PageBreak, Table, TableStyle,ListFlowable
from reportlab.lib import enums as rl_enums
from reportlab.lib import styles as rl_styles
from reportlab.lib.units import inch,mm
from reportlab.lib.pagesizes import LETTER, letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors
import re
from pandas import DataFrame
from reportlab.pdfbase.pdfmetrics import registerFontFamily
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFError
from reportlab.rl_config import defaultPageSize
import base64, json
from datetime import date
from os import system
import os
import configparser
import sys





FONT_DIR = "C:\\Windows\\Fonts"

BASE_DIR = os.path.dirname(__file__)

build_path = lambda rel_file: os.path.join(BASE_DIR,rel_file)


config = configparser.ConfigParser()



# Will need to be updated when moved to Demisto

config.read_file(open(build_path('config\\config.ini')))#os.path.join(BASE_DIR,'config\\config.ini')))



DEBUG = config["DEFAULTS"].getboolean("debug")




class PdfLetter:

    """
        PdfLetter is responsible for the merge between the json content and the custom_content passed.

        args:

            style_sheet
            page_style
            default_style
            doc
            title
            custom_content



    """

    def __init__(self,style_sheet,page_style,default_style,doc,title:str,custom_content:dict,*args,**kwargs):
        self.style_sheet = style_sheet
        self.page_style = page_style
        self.default_style = default_style
        self.doc = doc
        self.title = title
        self.story = []
        self.custom_content = custom_content


    @property
    def footer_text (self):
        return [footer for footer in self.footer if footer["type"].lower() != "image"]


    @property
    def footer_images(self):
        return [footer for footer in self.footer if footer["type"].lower() == "image"]


    @property
    def header_text (self):
        return [header for header in self.header if header["type"].lower() != "image"]


    @property
    def header_images (self):
        return [header for header in self.header if header["type"].lower() == "image"]


    @property
    def _file_name (self):
        return self.doc.filename



    def save_pdf (self,dir_location)->str:


        #if new_file_name != None:
        self.doc.filename = dir_location + self._file_name
        self.doc.build(self.story,onFirstPage=self._add_header_footer, onLaterPages=self._add_header_footer)


        return self.doc.filename

   

    def _add_style_from_json (self,new_style:json)->None:

        """
        Adds a style to the pages style sheet from json formating file.

        arguments:
            new_style (json): A json object representing a style to add to the PdfLetter style_sheet

        returns:
            None

        """

        for style in new_style:

            p_style = {}


            for k,v in style.items():



                # Convert any text rl_style enum values to the enum itself
                # this allows the whole dictionary of values to be passed as it relates to ReportLab enums
                # ex. {'alignment':'TA_JUSTIFY'} -> {'alignment':TA_JUSTIFY}
                style_enum = getattr(rl_enums,f"{v}",None)
                if style_enum == None:
                    p_style[k] = v
                else:
                    p_style[k] = style_enum

            try:
                self.style_sheet.add(rl_styles.ParagraphStyle(**p_style ))
            except KeyError as ke:
                pass



    def register_font (self,font_name:str)->None:

        """

            To be completed later

            args:
                font_name (str): The name of the font to register.

            returns:
                None

        """

        try:
            pdfmetrics.registerFont(TTFont(font_name, f'{font_name}.ttf'))

        except TTFError:

            print(f"Unable to locate font {font_name}, skiping for now. Check the format_file to ensure the name is correct and installed.")

        except Exception:

            pass


    def add_paragraph (self,jsf_paragraph:json, append_to_story=None)->Paragraph:

        """

            Adds a paragraph item to the story.

            arguments:
                jsf_paragraph (json): A json object representing a paragraph, including the style, content, and formating information. if "spacer" arguement is not None, the list will be converted into the appropriate tuple and passed to Spacer to append a spacer at the end of the paragraph.
                 append_to_story (boolean): default = False, Determines whether the Paragraph object is added to the story. For headers and footers they have to be drawn on the canvas this allowss the Paragraph object to be returned to the calling function.

            returns:

                Default: None,appends the Paragraph to the story;  append_to_story=True: return para (str) the the formated paragraph text

        """



        # format the text with font
        para_text = None



        # Check if a custom variable or item is set for this paragraph.
        # if one is set, call it.
        if jsf_paragraph.get('paragraph_key',False):

            try:

                # paragraph_key json key allows the use of multiple custom variables
                formated_content = jsf_paragraph["content"].format(**self.custom_content[jsf_paragraph['paragraph_key']])
                para_text = jsf_paragraph["font"].format(*formated_content.split(";;"))

        

            except Exception as e:
                print("unable to add custom content to {}".format(jsf_paragraph['font'])) # should switch to f strings, old habbits die hard
                if DEBUG:
                    raise e

        else:
            para_text = jsf_paragraph['font'].format(*jsf_paragraph['content'].split(";;"))

        
        para = Paragraph(para_text,self.style_sheet[jsf_paragraph['style'] ])



        if  append_to_story == "False":

            return para



        try:

            # If the paragraph is to be used as a heaader or footer have it return to the calling function to properly be drawn on the document

            self.story.append( para )

        except Exception as e:
            print("Error when trying to add the following paragraph :\n{}".format(jsf_paragraph['content']))
            if DEBUG:
                raise e



        if jsf_paragraph.get('spacer',False) :
            # Add spacer after the content
            self.story.append(self.add_spacer(jsf_paragraph['spacer'][0],jsf_paragraph['spacer'][1] ))



    def add_image (self,jsf_image:json, append_to_story=None)->Image:

        """

            Adds an image to the document

            arguements:

                jsf_image (json): a json object that describes the image including any base64 content that represents the image.



            returns:

                None

        """



        imgContent = build_path(jsf_image['name'])


        try:
            # Check to see if a named file already exists in the current directory
            with open(imgContent, 'rb') as f:
                f.read()



        except FileNotFoundError:

            imgdata = base64.b64decode(json.dumps(jsf_image['content']))
            imgContent = imgContent.split("\\")[-1] if os.name == 'nt' else imgContent.split("/")[-1]

            # Save the image data for future use if the file content isn't found
            with open(imgContent, 'wb') as f:
                f.write(imgdata)



        except Exception as e:
            print("Unable to save base64 encoded image data named {}".format(jsf_image['name']))
            if DEBUG:
                raise e


        img = Image(imgContent, jsf_image["width"]*inch,jsf_image["height"]*inch)


        # set the alignment of the image

        img.hAlign = jsf_image['hAlign']

        if  append_to_story == "False":
            return img



        self.story.append(img)

        if jsf_image.get("spacer",False):
            self.add_spacer( jsf_image["spacer"][0],jsf_image["spacer"][1])


        return img



    def add_bullet_list (self,jsf_list:json, append_to_story=None):
        lf = ListFlowable(
        [Paragraph(p, self.style_sheet[jsf_list['style'] ]) for p in jsf_list["content"]],
        bulletType=jsf_list['bulletType'],
        
        )
     
        self.story.append(lf)
        

        return

    def _add_header_footer (self,canvas,*args)->None:

        """

            Add the header and footers to the canvas

            args:
                canvas (canvas): The reportlab.canvas context to draw the header/footers
            returns:

                None

        """

        # Returns a list of just headers that are a paragraph type, this is done as writing multiple paragrpahs
        # to a header results in text that overlaps
        for i,header in enumerate(self.header):

            canvas.saveState()


            content_type = "add_" + header['type'].lower().replace(" ","_")
            p = getattr(self,content_type)(header, append_to_story="False")
            loc_y = (self.doc.height - ((i+1)/2)*(self.doc.topMargin*mm / (len(self.header)) )) + self.doc.topMargin



            if header.get("same_line",False):
                loc_y = (self.doc.height - ((i)/2)*(self.doc.topMargin*mm / (len(self.header)) )) + self.doc.topMargin


            w,h = p.wrap(self.doc.width,self.doc.topMargin) #(self.style_sheet[header["style"]].fontSize*mm))#


            p.drawOn(canvas, self.doc.leftMargin,loc_y)#

            canvas.restoreState()

        # Draw the footer paragraphs onto the document
        for i,footer in enumerate(self.footer[::-1]):
            canvas.saveState()

            content_type = "add_" + footer['type'].lower().replace(" ","_")
            p = getattr(self,content_type)(footer, append_to_story="False")
            loc_y = ((i+1)/2)*(self.doc.bottomMargin*mm / (len(self.footer)))
            if footer.get("same_line",False):
                loc_y = ((i+1))*(self.doc.bottomMargin*mm / (len(self.footer)))

            w,h = p.wrap(self.doc.width,self.doc.bottomMargin)

            p.drawOn(canvas, self.doc.leftMargin,loc_y)#



        canvas.restoreState()



    def add_spacer (self,*args)->None:

        """


            Adds a spacer to the story

            args:
                spacer_x (float): x cordinates of the spacer
                spacer_y (float): y cordinates of the spacer
            returns:
                None

        """

        try:

            if args[0].get("spacer",False) :
                spacer_x = args[0]["spacer"][0]
                spacer_y = args[0]["spacer"][1]

            else:
                spacer_x = args[0]
                spacer_y = args[1]


            self.story.append(Spacer(spacer_x,spacer_y))

        except Exception as e:

            print("Unable to add spacer, tried to add: {} {}".format(spacer_x,spacer_y))

            if DEBUG:

                raise e





    def add_current_date (self,jsf_current_date):

        """
            adds the current date to the story.
            to-do: allow for custom formates

            arguements:
                jsf_current_date (json): json object representing the date in json.


        """

        self.story.append(Paragraph(jsf_current_date['font'].format(date.today()),self.style_sheet[jsf_current_date['style'] ]))
        self.add_spacer(jsf_current_date['spacer'][0],jsf_current_date['spacer'][1])



    def add_table_reference (self,jsf_tables:json)->Table:
        # structure data
        t_data = [    [v["content"] for v in list(jsf_tables["headers"].values())] ]


        for row in jsf_tables["rows"][1:]:
            t_data.append([r["content"].format(**self.custom_content[r['paragraph_key']]) for r in list(row.values())[0]] )


        col_count = len(t_data[0])
        row_count = len(t_data)

        style=[
        ('ALIGN',(0,0),(-1,-1),'CENTER'),
        ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
        ('BOX', (0,0), (-1,-1), 0.25, colors.black)
        ]


        table = Table(t_data)
        table.setStyle(TableStyle(style))

        self.story.append(table)

        return None


    @staticmethod
    def get_date ():
        return date.today().strftime("%Y-%m-%d")


    def _parse_page_content (self,page_content:json)->None:

        """

                Parses the page content, calling the appropriate function to add the content to the story.

                args:
                    page_content (json): A json object representing a paragraph, including the style, content, and formating information. if "spacer" arguement is not None, the list will be converted into the appropriate tuple and passed to Spacer to append a spacer at the end of the paragraph.

                returns:
                    None



        """



        headers = []
        footers = []



        #for page in page_content:
        headers.extend(page_content["header"])
        footers.extend(page_content["footer"])



        for body_content in page_content['body']:
            if body_content["type"] == "table_reference":

                self.add_table_reference(page_content["tables"][body_content['table_reference']])

            else:
                content_type = "add_" + body_content['type'].lower().replace(" ","_")
                getattr(self,content_type)(body_content)



        # add page break on new pages

        self.story.append(PageBreak())
        self.header = headers
        self.footer = footers

    @staticmethod
    def convert_dataframe (dataframe:DataFrame,format_file:json)->dict:

        """


          The final data structure would look like this a list of nested dictionaries, where each item in the list is representing the data used in one letter. the ooutermost key of each dictionary item in the list corresponds to each paragraph_key named in the tamplate
          the key's used inner most  nested dictionary correspond to column names used in the letter.

          example of the structure
          [
            {
                paragraph_key: {ACCOUNT_LAST_4: value},
                "paragraph_key_2":{APPLICATION_ID: value},

                  ...

             },

          ]
          custom_content can be overriden with the list of dictionary items.
          This can be code that queries a database or loads an excel spreadsheet like in this example.
          the outer most key ("paragraph_key") represents how the code find the appropriate dictionary for the paragraph
          the second dictionary uses "content" keys to populate the appropriate value


          """
        template_keys,required_columns = PdfLetter.template_variables(format_file)
        if len(format_file["data_map"]) > 0:
            required_columns = [dm for dm in format_file["data_map"]]

        struct_df = dataframe[required_columns]
        struct_df_dict = struct_df.to_dict("records")       
        custom_content = []

        for i in struct_df_dict:
            # Create a dictionary to hold the paragraph keys (paragraph_keys)
            new_dict = {}
            for k,v in  template_keys.items():
                # set the paragraph_key and create a dictionary for the content keys
                new_dict[k] = {}
                for a in v:
                    # set the correct content key and value
                    new_dict[k][a] = i[a]
                
            custom_content.append(new_dict)


        return custom_content


    @staticmethod
    def template_variables (format_file):

        """

        Find the custom variables defined within the template and build the appropriate data structure to be populated by column data later.
            returns:

                self.custom_Vars (dict): Preliminary structure with which column data will populate

        """

        template_vars = {}
        required_columns = []
        # get all the custom vars that are defined in the sample paragraph
        for pType in ["header","footer","body"]:
            for cv in format_file["pages"][pType]:

                if  cv.get("paragraph_key","") != "":                  
                    content_keys = [s for s in cv["paragraph_key"].split(":")]#re.findall(r"(\{*\S+\}*)",cv["content"])]
                    required_columns.extend(content_keys)
                    template_vars[cv["paragraph_key"]] = content_keys
                else:
                    continue

        # finding content keys in tables, requires addtional iteration due to the table structure, which favors more human readability
        # over efficiency.
        for table in list(format_file["pages"]["tables"]):
            for r in list(table["rows"]):
                for c in  r.values():
                    for v in c:
                        content_keys = [s[1:-1] for s in re.findall(r"(\{\S*\s*\})+",v["content"])]
                        required_columns.extend(content_keys)
                        template_vars[v["paragraph_key"]] = content_keys
                        

        return template_vars,required_columns


    @classmethod
    def from_json_file (cls, format_file:json,custom_content:dict,file_name:str):
        """parses the Json file to create an instance of PdfFile, calling the appropriate methods to construct the story and set styles.

            args:

                format_file (json): A json object that describes the format of the pdf file
                custom_content(ordered dict): An Ordered dictionary that contains keys called by a content item's "paragraph_key" key.
                each value in the custom_content dictionary should be a list or tuple, even if it contains just one item.

            returns:

                None, saves the PDF file based on the json structure and default name provided.

        """

        ff_style = format_file["page_style"]

        doc = getattr(platypus,ff_style['template'])(file_name,
                                                    pagesize=LETTER,
                                                    rightMargin=ff_style["rightMargin"]*inch,
                                                    leftMargin=ff_style["leftMargin"]*inch,
                                                    topMargin=ff_style["topMargin"]*inch,
                                                    bottomMargin=ff_style["bottomMargin"]*inch)
        # intialize the doc, using the default_file name format

        pdfFile = cls(
                        style_sheet=getattr(rl_styles,ff_style['style_sheet'])(),
                        page_style =ff_style,
                        default_style=getattr(rl_styles,ff_style['style_sheet'])()[ff_style["default_style"]],
                          doc=doc,
                          title="Sample PDF",
                          custom_content=custom_content
                      )



        pdfFile._add_style_from_json(format_file['text_styles'])
        pdfFile._parse_page_content(format_file['pages'])

        return pdfFile

     