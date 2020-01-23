import sys
from mailmerge import MailMerge
import openpyxl
import os
from datetime import date

""" 
The Algorithm 

1. Export the name + major from the excel workbook
2. Match the major with the template file name
3. Update the template with the name and the contact info 
4. Populate a separate word doc with all the email drafts for that day  
5. Manually send emails 
"""

major_to_template = {"Economics": "EconomicsBA.docx",
                     "Computer Science": "ComputerScienceBS.docx",
                     "Exercise Science": "ExerciseScienceBS.docx",
                     "Russian Language": "RussianLanguageBS.docx",
                     "Sports Management": "SportsManagementBA.docx",
                     "Finance": "FinanceBS.docx"}


def load_wb_data():
    contacts_wb = str(sys.argv[1])

    wb = openpyxl.load_workbook(contacts_wb)
    sheet = wb['Sheet1']
    num_contact = len(sheet['A'])
    todays_date = date.today()

    todays_responses_directory = 'Responses/{}'.format(str(todays_date))
    os.mkdir(todays_responses_directory)
    for row in sheet.iter_rows(min_row=2, min_col=1, max_col=3, max_row=num_contact):
        name = str(row[0].value).strip()
        email = str(row[1].value).strip()
        major = str(row[2].value).strip()
        update_template(name, todays_responses_directory, major)


def update_template(name, directory, major):
    try:
        template = major_to_template[major]
    except KeyError:
        print("No such key in the major_to_template dict")
        # TODO: remove dir if error 
        return

    template_file = 'Templates/{}'.format(template)

    with MailMerge(template_file) as document:
        fields = document.get_merge_fields()
        document.merge(FullName=name, Major=major)
        file_name = "{}_{}.docx".format(name, major)
        document.write('{}/{}'.format(directory, file_name))


if __name__ == "__main__":
    # update_template()
    load_wb_data()