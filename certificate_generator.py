from openpyxl import load_workbook #used to import the Excel Data
from datetime import datetime #used to work with date times
#used for merge tags. If there is an error, uninstall and install docx-mailmerge
from mailmerge import MailMerge
import pandas as pd

def extractSheetData(path):
    df = pd.read_excel(path)
    print(df.head())


# # Getting Unique reps. Need to make each of their reports
# rep_list = []
# for cell_row in range(2 , max_row+1):
#     rep = sheet.cell(row = cell_row, column = 3).value
#     rep_list.append(rep)
# unique_rep_list = list(set(rep_list)) #getting unique list of reps

# For each rep, create their order reports
def mergeDocument(df=None):
    template_doc = "new_certificate_template.docx"
    # word_doc = MailMerge(template_doc)
    with MailMerge(template_doc) as document:
        print(document.get_merge_fields())


if __name__ == "__main__":
    extractSheetData('project_certificate_sheet.xlsx')
    mergeDocument()