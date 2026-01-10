
from openpyxl import Workbook, load_workbook
from config import *
from pprint import pprint

# abhi ke liye assume karunga ki control file me ek hi sheet hogi not sure ki alag alag branches ke liye alag sheet karni hai ya nahi
def get_control_dict(sheet):
    r = config.ROW_STARTING 
    control_dict = {}
    while sheet.cell(row=r, column=1).value is not None:
        name = sheet.cell(row=r, column=3).value
        btid = sheet.cell(row=r, column=2).value
        toenter = 4 
        subject_sets = set()
        while(sheet.cell(row=r, column=toenter).value is not None):     
            subject_sets.add(
                sheet.cell(row=r,column=toenter).value.strip()
            )
            toenter+=1
        data ={
            'row' : r,
            'name' : name,
            'toenter': toenter,
            'subjects_set': subject_sets
        }
        control_dict[btid] = data
        r+=1
    return control_dict

       


def process_subject_sheet(sheet, control):
    return 

def process_subject_file(file, control):
    return 


def main():
    control_file = "control.xlsx"
    control_wb  = load_workbook(control_file)
    control_sheet = control_wb.active
    control_dict  = get_control_dict(control_sheet)
    print(control_dict)


if __name__ == "__main__":
    main()