from restCall import AipRestCall
from restCall import AipData
from pptx import Presentation
from powerpoint import PowerPoint
import pandas as pd
import numpy as np 

import util 

aip_rest = AipRestCall("http://sha-dd-console:8080/CAST-RESTAPI-integrated/rest/","cast","cast",True)

project = "accela"    
apps = ["accela"] 
app_cnt = len(apps)
aip_data = AipData(aip_rest,project, apps)
all_apps_avg_grade = aip_data.calc_grades_all_apps()

ppt = PowerPoint("c:\\work\\data\\template.pptx","deck\\test.pptx")

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
    app_id = apps[app_no]
    print (f'Working on {app_id}')

    risk_grades = util.each_risk_factor(ppt, aip_data,app_id, app_no)

    grade_all = aip_data.get_app_grades(app_id)
    sizing = aip_data.get_loc_sizing(app_id)
    grade_by_tech_df = aip_data.get_grade_by_tech(app_id)

    snapshot = aip_data.snapshot(app=app_id)
    app_name = snapshot['name']
    loc = sizing['Number of Code Lines']

    ppt.replace_text(f'{{app{app_no+1}_all_technogies}}',util.list_to_text(snapshot['technology']))
    ppt.replace_text(f'{{app{app_no+1}_name}}',app_name)
    ppt.replace_text(f'{{app{app_no+1}_high_risk_grade_names}}',util.list_to_text(risk_grades.index.values))

    loc_df = aip_data.get_loc_sizing(app_id)
    loc = loc_df['Number of Code Lines']
    ppt.replace_loc(loc,app_no+1)

    ppt.replace_risk_factor(grade_all,app_no+1)
    ppt.replace_grade(grade_all,app_no+1)

    ppt.update_table(f'app{app_no+1}_grade_by_tech_table',grade_by_tech_df)
    ppt.update_chart(f'app{app_no+1}_sizing_pie_chart',grade_by_tech_df['LOC'])

    sizing_df = pd.DataFrame(aip_data.tech_sizing(app_id),index=[0])
    sizing_df['LoC']=pd.Series(["{0:.0f} K".format(val / 1000) for val in sizing_df['LoC']])
    sizing_df['Files']=pd.Series(["{0:.0f}".format(val) for val in sizing_df['Files']])
    sizing_df['Classes']=pd.Series(["{0:.0f}".format(val) for val in sizing_df['Classes']])
    sizing_df['SQL Artifacts']=pd.Series(["{0:.0f}".format(val) for val in sizing_df['SQL Artifacts']])
    sizing_df['Tables']=pd.Series(["{0:.0f}".format(val) for val in sizing_df['Tables']])
    sizing_df = sizing_df.transpose()
    ppt.update_table(f'app{app_no+1}_tech_sizing',sizing_df)

    violation_df = pd.DataFrame(aip_data.violation_sizing(app_id),index=[0])
    violation_df['Violation Count']=pd.Series(["{0:.0f}".format(val) for val in violation_df['Violation Count']])
    violation_df[' per file']=pd.Series(["{0:.2f}".format(val) for val in violation_df[' per file']])
    violation_df[' per kLoC']=pd.Series(["{0:.2f}".format(val) for val in violation_df[' per kLoC']])
    violation_df['Complex objects']=pd.Series(["{0:.0f}".format(val) for val in violation_df['Complex objects']])
    violation_df[' With violations']=pd.Series(["{0:.0f}".format(val) for val in violation_df[' With violations']])
    ppt.update_table(f'app{app_no+1}_violation_sizing',violation_df.transpose())

    rules_df = aip_data.critical_rules(app_id)
    critical_rule_df = pd.json_normalize(rules_df['rulePattern'])
    critical_rule_df = critical_rule_df[['name','critical']]
    rule_summary_df=critical_rule_df.groupby(['name']).size().reset_index(name='counts').sort_values(by=['counts'],ascending=False)
    rule_summary_df=rule_summary_df.head(10)
    ppt.update_table(f'app{app_no+1}_top_violations',rule_summary_df,include_index=False)




    (ap_df,ap_summary_df)=aip_data.action_plan(app_id)
    file_name = f'deck/{apps[app_no]}.xlsx'
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    col_widths=[10,50,50,30,30,30,30,30,30,30,30,30,30]
    util.format_table(writer,ap_df,'Action Plan',col_widths)
    col_widths=[50,50,10,30,30,30,30,30,30,30,30,30,30]
    util.format_table(writer,ap_summary_df,'Summary',col_widths)
    writer.save()




ppt.save()

