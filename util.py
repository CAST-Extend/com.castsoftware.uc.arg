import pandas as pd

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
        ppt.replace_block(f'{{app{app_no+1}_risk_detail}}',
                          f'{{end_app{app_no+1}_risk_detail}}',
                          "no high-risk health factors")
    else: 
        ppt.replace_text(f'{{app{app_no+1}_risk_category}}',risk_catagory)
        ppt.copy_block(f'app{app_no+1}_each_risk_factor',["_risk_name","_risk_grade"],len(risk_grades.count(axis=1)))
        f=1
        for index, row in risk_grades.T.iteritems():
            ppt.replace_text(f'{{app{app_no+1}_risk_name{f}}}',index)
            ppt.replace_text(f'{{app{app_no+1}_risk_grade{f}}}',row['All'].round(2))
            f=f+1

        ppt.replace_text(f'{{app{app_no+1}_risk_detail}}','')
        ppt.replace_text(f'{{end_app{app_no+1}_risk_detail}}','')

    ppt.remove_empty_placeholders()
    return risk_grades

def format_table(writer, data, sheet_name,width):
    
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

    col_width = 100
    for col_num, value in enumerate(data.columns.values):
        worksheet.write(0, col_num, value, header_format)
        w=width[col_num]
        worksheet.set_column(col_num, col_num, w)
    return worksheet

