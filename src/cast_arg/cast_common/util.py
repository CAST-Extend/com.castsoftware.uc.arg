from pandas import DataFrame,json_normalize,concat,ExcelWriter
from os import makedirs,listdir
from os.path import exists,abspath,join,isfile
from subprocess import Popen,PIPE,STDOUT
from pandas.api.types import is_numeric_dtype
from time import sleep
import sys
from urllib.parse import urlparse

from inquirer import List,prompt,Text,Confirm

def yes_no_input(prmpt:str,default_value=True) -> bool:
    if default_value:
        default_value='Y'
    else:
        default_value='N'

    question = [
        Text('rslt', message=f'{prmpt} [{default_value}]?')
    ]

    while True:
        # val = input(f'{prompt} [{default_value}]?').upper()
        answer = prompt(question)
        val = answer['rslt'].upper()
        if len(val) == 0: val = default_value

        if val in ['Y','YES']:
            return True
        elif val in ['N','NO']:
            return False
        else:
            print ('Expecting a Y[es] or N[o] response. ',end='')

from glob import glob

def _string_question(prmpt:str, default_value=""):
    question = []
    if len(default_value)>0:
        question.append(Text('rslt', message=f'{prmpt} [{default_value}]?'))
        # i = input(f'{prmpt} [{default_value}]: ')
    else:
        question.append(Text('rslt', message=f'{prmpt}? '))
    return question

def folder_input(prmpt:str,folder:str='',file:str=None,combine:bool=False,create=False,use_default=False) -> str:
    if combine:
        folder=folder.replace(file,'')

    while True:
        test_folder = abspath(f'{folder}')
        if use_default and exists(abspath(f'{test_folder}/{file}')):
            folder = abspath(f'{folder}')
            break

        question = _string_question(prmpt, folder)
        answer = _get_answer(prompt(question), folder)
        if len(answer)==0:
            folder = abspath(folder)
        else:
            folder = abspath(answer)
        if exists(folder):
            if file is not None:
                if glob(abspath(f'{folder}/{file}')):
                    break
                else:
                    print (f'{folder} folder does not contain {file}, please try again')
                    continue
            else:
                break
        
        if create and yes_no_input(f'{folder} does not exist, create it?'):
            makedirs(folder)
            break

        print (f'{folder} not found, please try again')

    if combine:
        folder = abspath(f'{folder}/{file}')
    return folder

def file_input(prmpt:str, folder:str='', file:str=None, combine:bool=False) -> str:

    if not file is None:
        path = abspath(f'{folder}/{file}')
        files = [file for file in glob(path) if isfile(join(path,file))]
        if len(files)==1:
            return files[0]

    question = _string_question(prmpt, folder)
    while True:
        test_folder = abspath(f'{folder}')
        answer = prompt(question)
        if len(answer)==0:
            folder = abspath(folder)
        else:
            folder = abspath(answer)
        if exists(folder):
            if file is not None:
                if glob(abspath(f'{folder}/{file}')):
                    break
                else:
                    print (f'{folder} folder does not contain {file}, please try again')
                    continue
            else:
                break

        if create and yes_no_input(f'{folder} does not exist, create it?'):
            makedirs(folder)
            break

        print (f'{folder} not found, please try again')

    if combine:
        folder = abspath(f'{folder}/{file}')
    return folder

def _get_answer(answers:list, default_value=""):
    rslt = answers['rslt'].strip()
    if len(rslt)==0:
        rslt = default_value
    return rslt

def string_input(prmpt:str,default_value=""):
    question = _string_question(prmpt,default_value)

    while True:
        answer = _get_answer(prompt(question), default_value)
        if len(answer)==0:
            if len(default_value) == 0:
                print('Input required, please try again')
                continue
        else:
            break
    return answer

def secret_input(prmpt:str,default_value=""):
    if len(default_value)>0:
        prmpt = f'{prmpt} [***********]'

    question = _string_question(prmpt)

    while True:
        answer = _get_answer(prompt(question), default_value)
        if len(answer)==0:
            if len(default_value) == 0:
                print('Input required, please try again')
                continue
            else:
                answer = default_value
        return answer            

def uri_validator(x):
    try:
        result = urlparse(x)
        return all([result.scheme, result.netloc])
    except Exception as ex:
        return False

def url_input(prmpt:str,default_value:str):
    question = _string_question(prmpt, default_value)

    while True:
        answers = prompt(question)
        answer = answers['rslt'].strip()
        if len(answer)==0:
            answer = default_value
        if not uri_validator(answer):
            print ("Bad URL, please try again")      
        else:
            return answer

def get_between(txt,tag_start,tag_end,start_at=0):
    text = txt[start_at:]
    
    start = text[start_at:].find(f"{tag_start}")+len(f"{tag_start}")
    end = text[start:].find(f"{tag_end}")
    between = text[start:start+end].strip()

    return between,start,end,start-len(tag_start),end+len(tag_end)

def list_to_text(list):
    rslt = ""

    l = len(list)
    if l == 0:
        return ""
    elif l == 1:
        return list[0]

    data = list
    if l > 2:
        last_name = list[-2]
    else:
        last_name = ''

    for a in list:
        rslt = rslt + a + ", "

    rslt = rslt[:-2]
    rslt = rreplace(rslt,', ',' and ')

    return rslt

def rreplace(s, old, new, occurrence=1):
    li = s.rsplit(old, occurrence)
    return new.join(li)

def each_risk_factor(ppt, aip_data, app_id, app_no):
    # collect the high risk grades
    # if there are no high ris grades then use medium risk
    app_level_grades = aip_data.get_app_grades(app_id)
    risk_grades = aip_data.calc_health_grades_high_risk(app_level_grades)
    risk_catagory="high"
    if risk_grades.empty:
        risk_catagory="medium"
        risk_grades = aip_data.calc_health_grades_medium_risk(app_level_grades)
    
    #in the event all health risk factors are low risk
    if risk_grades.empty:
        ppt.replace_block(f'{{app{app_no}_risk_detail}}',
                          f'{{end_app{app_no}_risk_detail}}',
                          "no high-risk health factors")
    else: 
        ppt.replace_text(f'{{app{app_no}_risk_category}}',risk_catagory)
        ppt.copy_block(f'app{app_no}_each_risk_factor',["_risk_name","_risk_grade"],len(risk_grades.count(axis=1)))
        f=1
        for index, row in risk_grades.T.iteritems():
            ppt.replace_text(f'{{app{app_no}_risk_name{f}}}',index)
            ppt.replace_text(f'{{app{app_no}_risk_grade{f}}}',row['All'].round(2))
            f=f+1

        ppt.replace_text(f'{{app{app_no}_risk_detail}}','')
        ppt.replace_text(f'{{end_app{app_no}_risk_detail}}','')

    ppt.remove_empty_placeholders()
    return risk_grades

def format_table(writer, data, sheet_name,width=None,total_line:bool=False,decimal_columns=None):
    if decimal_columns is None:
        decimal_columns = []

    data.to_excel(writer, index=False, sheet_name=sheet_name, startrow=1,header=False)

    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    rows = len(data)
    if total_line:
        rows += 1
    cols = len(data.columns)-1
    columns=[]

    # Add a custom number format with commas to the workbook
    comma_format = workbook.add_format({'num_format': '#,##0'})         # No decimals
    decimal_format = workbook.add_format({'num_format': '#,##0.00'})    # With decimals

    first=True
    for col_num, value in enumerate(data.columns.values):
        json = {'header': value}

        if is_numeric_dtype(data[value]) and data[value].dtype != 'bool':
            if value in decimal_columns:
                json['format'] = decimal_format
            else:
                json['format'] = comma_format

        if first:
            first=False
            if total_line:
                json['total_string']='Totals'
        else:
            if is_numeric_dtype(data[value]) and data[value].dtype != 'bool':
                if total_line:
                    json['total_function']='sum'

        columns.append(json)

    table_options={
        'columns':columns,
        'header_row':True,
        'autofilter':True,
        'banded_rows':True,
        'total_row':total_line
    }

    worksheet.add_table(0, 0, rows, cols,table_options)

    header_format = workbook.add_format({'text_wrap':True,
                                        'align': 'center'})

    col_width = 10
    if width is None:
        width = []
        for col in data.columns:
            x = data[col].astype(str).str.len().max()
            if x < 15: x = 10
            if x > 100: x = 150
            width.append(x)

        # for i in range(1,len(data.columns)+1):
        #    width.append(col_width)
    for col_num, value in enumerate(data.columns.values):
        worksheet.write(0, col_num, value, header_format)
        col_format = None
        if is_numeric_dtype(data[value]) and data[value].dtype != 'bool':
            col_format = decimal_format if value in decimal_columns else comma_format
        worksheet.set_column(col_num, col_num, width[col_num], col_format)
        
    return worksheet

def find_nth(string, substring, n):
   if (n == 1):
       return string.find(substring)
   else:
       return string.find(substring, find_nth(string, substring, n - 1) + 1)

def no_dups(string, separator,add_count=False):
    alist = list(string.split(separator))
    alist.sort()
    nlist = []
    clist = []
    for i in alist:
        if i not in nlist:
            nlist.append(i)
            clist.append(1)
        else:
            idx = nlist.index(i)
            clist[idx]=clist[idx]+1

    if add_count:
        for val in nlist:
            idx = nlist.index(val)
            cnt = clist[idx]
            
            cval=''
            if cnt > 1:
                cval = f'({cnt})'
            val = f'{val}{cval}'
            nlist[idx]=val

    string = separator.join(nlist)
    return string

def toExcel(file_name, tabs):
    writer = ExcelWriter(file_name, engine='xlsxwriter')
    for key in tabs:
        format_table(writer,tabs.get(key),key)
    writer.save()

def resource_path(relative_path):
    "get the absolute path to resource, works for dev and for PyInstaller"
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = abspath('.')
    return join(base_path, relative_path)

def create_folder(folder):
    if not exists(folder):
        makedirs(folder)

def run_process(args,wait=True,output=True) -> int:
#    process = Popen(args, stdout=PIPE, stderr=PIPE)
    process = Popen(args, stdout=PIPE, stderr=STDOUT,text=True,bufsize=1)
    if wait:
        return check_process(process,output)
    else:
        return process

def check_process(process:Popen,output=True):
    ret = []
    while process.poll() is None:
        line = process.stdout.readline()
        line = line.lstrip("b'").rstrip('\n')
        if output == True and len(line.strip(' ')) > 0:
            print(line)
        ret.append(line)
    stdout, stderr = process.communicate()
    if not stdout is None and len(stdout): ret.append(stdout)
    if not stderr is None and len(stderr): ret.append(stderr)
    return process.returncode,ret

def track_process(proc):
    if proc.poll() is None:
        while line := proc.stdout.readline():
            print(line.replace('\n','                                               '),end='\r')
            sleep(.5)


def convert_LOC(total:int):
    unit = ''
    if 1000 <= total <= 1000000:
        unit = 'KLOC'
        total = int(total/1000)
    elif total > 1000000:
        unit = 'MLOC'
        total = round(total/1000000,1)
    return total,unit