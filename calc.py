import pathlib
import pandas as pd
import os

class readAndSummarize:

    def __init__(self):
        self.path = str(pathlib.Path("calc.py").parent.resolve()) #<-- string in Path() should be the name of this python-file.

    def read(self, file):
        self.sheet = pd.read_excel(self.path + f"\\zettle_reports\\{file}", engine="openpyxl")
    
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
                                                 "utbetaling": round(self.sheet_sellers.loc[self.sheet_sellers["Navn"] == x, "Total inkl. mva. (NOK)"].item() * 0.6825, 2)
                                                }, self.sheet_sellers["Navn"].unique()))
    
    def save_reports(self):
        print(self.path)
     
        directory_index = len(os.listdir(self.path + "\\personal_seller_reports"))
        os.mkdir(self.path + "\\personal_seller_reports\\" + str(directory_index))

        for i, report in enumerate(self.reports):
           # print(report["Selgernummer"])
            file_path = self.path + "\\personal_seller_reports\\" + str(directory_index) + "\\" + report["Selgernummer"]
            print(file_path)
            with open(file_path + ".pdf", "w", encoding="utf-8") as file:
                file.write("OVERSIKT FOR PERIODEN " + str(self.periode) + "\n\n"
                            + "Selgernummer: " + str(report['Selgernummer']) + "\n"
                            + "Antall salg: " + str(report['antall salg']) + "\n"
                            + "Total salgsum: " + str(report['total salgsum']) + "\n"
                            + "Til utbetaling: " + str(report["utbetaling"])
                           )



        



test = readAndSummarize()
test.read("Zettle-Sales-By-Product-Report-20240923-20241006.xlsx")
test.clean()
test.calculate()
test.save_reports()
