
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

       


def process_subject_sheet(sheet, control, newdict):
    
    def find_course_name():
        c=1
        r=config.COURSE_PAGE_HEADER_WITH_COURSE_TITLE_ROW_NO-1
        while sheet.cell(row=r, column=c).value!='Cource Title':
            c+=1
        name= sheet.cell(row=r+1,column=c).value.strip()
       
        name = name.replace('\n', ' ')
        return name
    
    def get_cols():
        r=config.COURSE_PAGE_HEADER_WITH_COURSE_TITLE_ROW_NO+1
        c=1
        reexam_col = -1
        name_col = -1
        btid_col = -1

        while True:
            v = sheet.cell(row=r, column=c)
            
            if(v.value is not None):
                vl = v.value.strip().replace('\n', ' ')
                if(vl=="Re-Exam Grades"):  
                    reexam_col = c
                    break
                elif(vl=="Name"):
                    name_col = c
                elif(vl=="Roll No."):
                    btid_col = c
            if(c>=1000): raise Exception("No column found with the name 'Re-Exam Grades' ")
            c+=1
        return reexam_col, name_col, btid_col


    def find_starting_row():
        r = config.COURSE_PAGE_HEADER_WITH_COURSE_TITLE_ROW_NO + 1
        while sheet.cell(row=r, column=1).value != 1:
            r+=1
        return r

    cource_title = find_course_name()

    print(cource_title)
    starting_row = find_starting_row()
    r = starting_row
    reexam_grade_col, name_col, btid_col = get_cols()
    # print(reexam_grade_col)
    while True:
        
        name_val = sheet.cell(row=r, column=name_col).value
        btid_val = sheet.cell(row=r, column=btid_col).value
        grade_val = sheet.cell(row=r, column=reexam_grade_col).value

      
        if name_val is None and btid_val is None:
            break

    
        name = str(name_val).strip() if name_val else ""
        btid = str(btid_val).strip() if btid_val else ""
        reexam_grade_value = str(grade_val).strip() if grade_val else ""

       
        if reexam_grade_value == "FF":
            if btid in control:
                
                if cource_title not in control[btid].get('subjects_set', set()):
                    newdict.setdefault(btid, []).append(cource_title)
            else:
               
                newdict.setdefault(btid, []).append(cource_title)

        
        print(f"Processed row: {r}")
        r += 1
def process_subject_file(file, control, newdict):
    subject_wb = load_workbook(file,data_only=True)
    
    for sheetname in subject_wb.sheetnames:
        sheet = subject_wb[sheetname]
        process_subject_sheet(sheet, control, newdict)
   
def main():
    control_file = "control.xlsx"
    control_wb  = load_workbook(control_file)
    control_sheet = control_wb.active
    control_dict  = get_control_dict(control_sheet)
    # print(control_dict)
    newdict = dict()

    subject_file = "csh2.xlsx"
    process_subject_file(subject_file, control_dict, newdict)
    process_subject_file("csh4.xlsx", control_dict, newdict)
    pprint(newdict)

if __name__ == "__main__":
    main()