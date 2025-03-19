import pathlib
import pandas as pd
import os
import base64
import pdfkit
# TODO: fiks pdf: https://ironpdf.com/python/blog/using-ironpdf-for-python/python-create-pdf-with-text-and-images/


class readAndSummarize:

    def __init__(self, file):
        self.path = str(pathlib.Path("calc.py").parent.resolve()) #<-- string in Path() should be the name of this python-file.
        self.read(file)
        self.clean()
        self.calculate()
        self.save_reports()

    def read(self, file):
        self.sheet = pd.read_excel(file)

    def clean(self):
        self.periode = self.sheet.iloc[3][1]
        self.sheet = self.sheet.iloc[5:] #Removing top 5 rows
        self.sheet.columns = self.sheet.iloc[0] #Changing column values into first row.
        self.sheet = self.sheet.iloc[1:] #Removing first row
        self.sheet_sellers = self.sheet[self.sheet["Navn"].str.startswith("SELGER")] #Filtering out all non-personal sales
    
    def calculate(self):
        self.reports = list(map(lambda x: {"Selgernummer": x[-4:], 
                                                 "antall salg": self.sheet_sellers.loc[self.sheet_sellers["Navn"] == x, "Totalt antall"].item(),
                                                 "total salgsum": self.sheet_sellers.loc[self.sheet_sellers["Navn"] == x, "Total inkl. mva. (NOK)"].item(),
                                                 "utbetaling": round(self.sheet_sellers.loc[self.sheet_sellers["Navn"] == x, "Total inkl. mva. (NOK)"].item() * 0.5825, 2)
                                                }, self.sheet_sellers["Navn"].unique()))
    
    def save_reports(self):
        
        os.mkdir(self.path + "\\personal_seller_reports\\" + str(self.periode))
        
        path_wkhtmltopdf = "C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe"
        config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf) #<-- CHANGE TO PERSONAL WKHTMLTOPDF-PATH 
        for report in self.reports:
            file_path = self.path + "\\personal_seller_reports\\" + str(self.periode) + "\\" + report["Selgernummer"]
            #print(file_path)
            report_content = """ 
            <!DOCTYPE html>
            <html lang="en">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>PDF Generation Example</title>
                <style>
                    body {{
                        margin: 100px; /* Remove default margin */
                        padding: 20px; /* Add padding for better visibility */
                        position: relative; /* Set the body as a relative position */
                        }}
                        header {{
                                text-align: center;
                                margin-bottom: 40px;
                                }}
                        section {{
                            margin-top: 20px; /* Add margin to separate sections */
                                }}
                        img {{
                            position: absolute;
                            top: 0px; /* Adjust the top position */
                            right: 20px; /* Adjust the right position */
                            }}
                        p {{
                            letter-spacing: 20px;
                            }}
                </style>
            </head>
            <body>
                <header>
                    <h1>ÖY Salgsrapport</h1>
                </header>
                <section id="contentSection">
                    <h2>Oversikt for perioden {periode}</h2>
                    <p>Selgernummer: {selger_nr}</p>
                    <p>Antall salg: {salg_antall}</p>
                    <p>Total salgsum: {salg_sum}</p>
                    <p>Til utbetaling: {utbetaling}</p>
                        """.format(periode = str(self.periode),selger_nr = str(report['Selgernummer']), salg_antall = str(report['antall salg']), salg_sum = str(report['total salgsum']),utbetaling = str(report["utbetaling"]))
            
            with open("OY_logo_svart.png", "rb") as f:
                pngBinaryData = f.read()
                imgDataUri = "data:image/png;base64," + base64.b64encode(pngBinaryData).decode("utf-8")
                imgHtml = f"""
                            <!-- Embedded Image -->
                            <img src='{imgDataUri}' width=100px height=100px
                                alt="ÖY logo">
                            </section>
                            </body>
                            </html>
                           """
            
            report_content += imgHtml #<-- Adding image to html-report
       
            pdfkit.from_string(report_content, file_path + ".pdf", configuration = config)
            

