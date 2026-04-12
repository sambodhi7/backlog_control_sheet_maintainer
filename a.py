from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from config import config

def get_control_dict(sheet):
    r = config.ROW_STARTING 
    control_dict = {}
    while sheet.cell(row=r, column=1).value is not None:
        name = sheet.cell(row=r, column=3).value
        btid = sheet.cell(row=r, column=2).value
        toenter = 4 
        subject_sets = set()
        while(sheet.cell(row=r, column=toenter).value is not None):    
            subject =  (sheet.cell(row=r,column=toenter).value).strip()
            if subject in codeLookup:
                subject = codeLookup[subject]
            subject_sets.add(subject)
            toenter+=1
        clear_size = toenter - 4
        
        toenter=4
        control_dict[btid] = {
            'row' : r,
            'name' : name,
            'toenter': toenter,
            'subjects_set': subject_sets,
            'clear_size': clear_size
        }
        r+=1
    control_dict["last_row"] = r
    return control_dict



def util_format_string(s):
    if s is None:
        return ""
    return s.strip().upper().replace(" ", "")

def get_sub_grade_from_sterm_str(s):
    s = s.upper().strip()
    s = s.replace("S.TERM", "")
    for txt in s.split(")"):
        if not txt:
            continue
        txt = txt.strip()

        grade, rest = txt.split("(")
        grade=grade.replace(",", "")
        grade = grade.strip()
        sub = rest.split(",")[0]
        sub=sub.strip()
        yield sub, grade



def process_subject_sheet2(sheet, control, newdict, codeLookup):
    row = config.HEADERR_STARTING
    c = 1
    cell = sheet.cell(row = row ,column = c)
    while util_format_string(cell.value)!='COURSECODE':
        c+=1
        cell = sheet.cell(row = row ,column = c)
        if c > 20:
            print("Course code header not found in sheet", sheet.title)
            return
    c+=1

    # print(sheet.cell(row = row ,column = c).value)

    def get_row_to_btid_dict(): 
        r = config.ROW_STARTING
        res = {}
        while True:
            value  = sheet.cell(row=r, column=1).value
            value = util_format_string(value)
            if "NO" in value:
                break 
            r+=1
        r+=1
        rs=r
        
        while True:
            value  = sheet.cell(row=r, column=1).value
            if value is None:
                break 
            btid = sheet.cell(row=r, column=2).value
            if btid is not None:
                res[r] = util_format_string(btid)
    
            r+=1
        return rs, res
    
    row_s, row_to_btid_dict = get_row_to_btid_dict()
    
     
    while(1):
        course_code_cell = sheet.cell(row = row ,column = c)
        if course_code_cell.value is None:
            break
        
        grade_cell = sheet.cell(row = row_s ,column = c).value
        ct = 0
        while(ct<config.MAX_NUMBER_OF_STUDENTS): 
            grade_cell = sheet.cell(row = row_s + ct ,column = c).value
            if row_to_btid_dict.get(row_s + ct) is None:
                break
            if grade_cell is None:
                ct+=1
                continue
            grade = util_format_string(grade_cell.value)
            btid = row_to_btid_dict[row_s + ct]
            
            course_title = sheet.cell(row=row+1 ,column = c).value
            course_code = util_format_string(course_code_cell.value)
            
            if course_code not in codeLookup:
                codeLookup[course_code] = course_title
                        
            ct+=1
            
        c+=1
        

def process_subject_file2(file, control, newdict, codeLookup):
    subject_wb = load_workbook(file, data_only=True)

    for sn in subject_wb.sheetnames:
        sn = sn.upper()
        if "SEM" in sn:
            process_subject_sheet2(subject_wb[sn], control, newdict, codeLookup)

def main():
    process_subject_file2("newdata.xlsx", {} , {}, {})

if __name__ == "__main__":
    main()