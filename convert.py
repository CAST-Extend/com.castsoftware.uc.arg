from restCall import AipRestCall
from restCall import AipData
from pptx import Presentation
from powerpoint import PowerPoint


import pandas as pd



def each_risk_factor(ppt, aip_data, app_no):
    app_id = apps[app_no]

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

aip_rest = AipRestCall("http://sha-dd-console:8080/CAST-RESTAPI-integrated/rest/","cast","cast",True)

project = "Blackhawks"    
apps = ["mobile_doorman_ios","mobile_doorman_android"] 
app_cnt = len(apps)
aip_data = AipData(aip_rest,project, apps)
all_apps_avg_grade = aip_data.calc_grades_all_apps()

ppt = PowerPoint("c:\\work\\data\\template.pptx","..\\data\\test.pptx")

ppt.replace_text("{project}",project)
ppt.replace_text("{app_count}",len(apps))
ppt.replace_text("{all_apps}",aip_data.get_all_app_text())
ppt.replace_risk_factor(all_apps_avg_grade,search_str="{summary_")

risk_grades = aip_data.calc_health_grades_high_risk(all_apps_avg_grade)
if risk_grades.empty:
    risk_grades = aip_data.calc_health_grades_medium_risk(all_apps_avg_grade)

if risk_grades.empty:
    ppt.replace_block('{immediate_action}','{end_immediate_action}','')
else:
    ppt.replace_text("{immediate_action}","") 
    ppt.replace_text("{end_immediate_action}","") 
    ppt.replace_text("{high_risk_grade_names}",aip_data.text_from_list(risk_grades.index.values.tolist()))

ppt.duplicate_slides(app_cnt)

app_no=1
ppt.copy_block("each_app",["app"],app_cnt)
for app_no in range(0,app_cnt):
    each_risk_factor(ppt, aip_data,app_no)

    app_id = apps[app_no]
    grade_all = aip_data.get_app_grades(app_id)
    sizing = aip_data.get_loc_sizing(app_id)
    grade_by_tech_df = aip_data.get_grade_by_tech(app_id)



    snapshot = aip_data.snapshot(app=app_id)
    app_name = snapshot['name']
    loc = sizing['Number of Code Lines']

    ppt.replace_text(f'{{app{app_no+1}_name}}',app_name)
#    ppt.replace_text(f'{{app{app_no+1}_high_risk_grade_names}}',aip_data.get_high_risk_grade_text(grade_all).lower())
#    ppt.replace_text(f'{{app{app_no+1}_medium_risk_grade_names}}',aip_data.get_medium_risk_grade_text(grade_all).lower())
    ppt.replace_risk_factor(grade_all,app_no+1)
    ppt.replace_grade(grade_all,app_no+1)
    ppt.update_table(f'app{app_no+1}_grade_by_tech_table',grade_by_tech_df)
    ppt.update_chart(f'app{app_no+1}_sizing_pie_chart',grade_by_tech_df['LOC'])



ppt.save()

