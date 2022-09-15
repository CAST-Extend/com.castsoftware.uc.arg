import pandas as pd
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


def format_table(writer, data, sheet_name,width=None):
    
    data.to_excel(writer, index=False, sheet_name=sheet_name, startrow=1,header=False)

    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    rows = len(data)
    cols = len(data.columns)-1
    columns=[]
    for col_num, value in enumerate(data.columns.values):
        columns.append({'header': value})

    table_options={
                'columns':columns,
                'header_row':True,
                'autofilter':True,
                'banded_rows':True
                }
    worksheet.add_table(0, 0, rows, cols,table_options)
    
    header_format = workbook.add_format({'text_wrap':True,
                                        'align': 'center'})

    col_width = 10
    if width == None:
        width = []
        for i in range(1,len(data.columns)+1):
           width.append(col_width)
    for col_num, value in enumerate(data.columns.values):
        worksheet.write(0, col_num, value, header_format)
        w=width[col_num]
        worksheet.set_column(col_num, col_num, w)
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

