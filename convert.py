from restCall import AipRestCall
from restCall import AipData
from pptx import Presentation
from powerpoint import PowerPoint
from jproperties import Properties 
from pptx.dml.color import RGBColor

import pandas as pd
import numpy as np 
import util 
import math
import argparse

class GeneratePPT:

    _app_list = []
    _ppt = None
    _aip_data = None
    _effort_df = None
    _output_folder = None

    def __init__(self,project,app_list,template,output_folder):
        self._app_list = app_list
        self._output_folder=output_folder

        out = f"{self._output_folder}/{project}.pptx"
        print (f'Generating {out}')

        self._rest = AipRestCall("http://sha-dd-console:8080/CAST-RESTAPI-integrated/rest/","cast","cast",True)
        self._aip_data = AipData(self._rest,project, self._app_list)

        self._ppt = PowerPoint(template,out)
        self._effort_df = pd.read_csv('./Effort.csv')

        all_apps_avg_grade = self._aip_data.calc_grades_all_apps()

        #project level work
        app_cnt = len(self._app_list)

        self._ppt.duplicate_slides(app_cnt)
        self._ppt.copy_block("each_app",["app"],app_cnt)

        self._ppt.replace_text("{project}",project)
        self._ppt.replace_text("{app_count}",app_cnt)
        self._ppt.replace_text("{all_apps}",self._aip_data.get_all_app_text())
        self._ppt.replace_risk_factor(all_apps_avg_grade,search_str="{summary_")

        risk_grades = self._aip_data.calc_health_grades_high_risk(all_apps_avg_grade)
        if risk_grades.empty:
            risk_grades = self._aip_data.calc_health_grades_medium_risk(all_apps_avg_grade)

        if risk_grades.empty:
            self._ppt.replace_block('{immediate_action}','{end_immediate_action}','')
        else:
            self._ppt.replace_text("{immediate_action}","") 
            self._ppt.replace_text("{end_immediate_action}","") 
            self._ppt.replace_text("{high_risk_grade_names}",self._aip_data.text_from_list(risk_grades.index.values.tolist()))

        for app_no in range(0,app_cnt):
            app_id = self._app_list[app_no]
            print (f'Working on {app_id}')

            risk_grades = util.each_risk_factor(self._ppt, self._aip_data,app_id, app_no)

            grade_all = self._aip_data.get_app_grades(app_id)
            self._ppt.replace_risk_factor(grade_all,app_no)
            grade_by_tech_df = self._aip_data.get_grade_by_tech(app_id)
            self._ppt.update_table(f'app{app_no+1}_grade_by_tech_table',grade_by_tech_df)
            self._ppt.update_chart(f'app{app_no+1}_sizing_pie_chart',grade_by_tech_df['LOC'])

            snapshot = self._aip_data.snapshot(app=app_id)
            app_name = snapshot['name']
            self._ppt.replace_text(f'{{app{app_no+1}_name}}',app_name)
            self._ppt.replace_text(f'{{app{app_no+1}_all_technogies}}',util.list_to_text(snapshot['technology']))
            self._ppt.replace_text(f'{{app{app_no+1}_high_risk_grade_names}}',util.list_to_text(risk_grades.index.values))


            loc_df = self._aip_data.get_loc_sizing(app_id)
            loc = loc_df['Number of Code Lines']
            self._ppt.replace_loc(loc,app_no+1)

            self._ppt.replace_grade(grade_all,app_no+1)

            self.fill_sizing(app_no)
            self.fill_critical_rules(app_no)
            self.fill_action_plan(app_no)
            self.fill_violations(app_no)

        self._ppt.save()


    def fill_critical_rules(self,app_no):
        app_id = self._app_list[app_no]
        rules_df = self._aip_data.critical_rules(app_id)
        if not rules_df.empty:
            critical_rule_df = pd.json_normalize(rules_df['rulePattern'])
            critical_rule_df = critical_rule_df[['name','critical']]
            rule_summary_df=critical_rule_df.groupby(['name']).size().reset_index(name='counts').sort_values(by=['counts'],ascending=False)
            rule_summary_df=rule_summary_df.head(10)
            self._ppt.update_table(f'app{app_no+1}_top_violations',rule_summary_df,include_index=False)

    def fill_action_plan(self,app_no):
        app_id = self._app_list[app_no]

        (ap_df,ap_summary_df)=self._aip_data.action_plan(app_id)
        if not ap_summary_df.empty:
            ap_summary_df = ap_summary_df.merge(self._effort_df, how='inner', on='Technical Criteria')
            #cost_col = (ap_summary_df['Eff Hours'] * ap_summary_df['No. of Actions'])/8
            ap_summary_df['Days Effort'] = (ap_summary_df['Eff Hours'] * ap_summary_df['No. of Actions'])/8
            ap_summary_df['Cost Est.'] = ap_summary_df['Days Effort'] * 600

            file_name = f'{self._output_folder}/{app_id}_action_plan.xlsx'
            writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
            col_widths=[50,40,10,10,10,50,10,10,10]
            summary_tab = util.format_table(writer,ap_summary_df,'Summary',col_widths)
            col_widths=[10,50,50,30,30,30,30,30,30,30,30,30,30]
            util.format_table(writer,ap_df,'Action Plan',col_widths)
            writer.save()

            #fill action plan related tags
            self.fill_action_plan_text(ap_summary_df,app_no,'extreme','security')
            self.fill_action_plan_text(ap_summary_df,app_no,'high')
            self.fill_action_plan_text(ap_summary_df,app_no,'moderate')
            self.fill_action_plan_text(ap_summary_df,app_no,'low')

            #configure action plan table background colors 
            ap_summary_df.loc[ap_summary_df['tag']=='extreme','RGB']='244,212,212'
            ap_summary_df.loc[ap_summary_df['tag']=='high','RGB']='255,229,194'
            ap_summary_df.loc[ap_summary_df['tag']=='moderate','RGB']='203,225,238'
            ap_summary_df.loc[ap_summary_df['tag']=='low','RGB']='254,254,255'

            ap_table = pd.concat([ap_summary_df[ap_summary_df['tag']=='extreme'],
                                  ap_summary_df[ap_summary_df['tag']=='high'],
                                  ap_summary_df[ap_summary_df['tag']=='moderate'],
                                  ap_summary_df[ap_summary_df['tag']=='low']])

            ap_table = ap_table.drop(columns=['comment','tag','Technical Criteria','Days Effort','Cost Est.','Eff Hours'],index=1)

            self._ppt.update_table(f'app{app_no+1}_action_plan',ap_table.head(29),include_index=False,background='RGB')

            sum = ap_summary_df['No. of Actions'].sum()
            self._ppt.replace_text(f"{{app{app_no+1}_total_violations}}",str(sum)) 




    def fill_action_plan_text(self,ap_summary_df,app_no,priority,default=''):
        (priority_text, violation_total, filtered) = self.common_business_criteria(ap_summary_df,priority,default)
        self._ppt.replace_text(f"{{app{app_no+1}_{priority}_business_criteria_text}}",priority_text.lower()) 
        self._ppt.replace_text(f"{{app{app_no+1}_{priority}_violation_total}}",violation_total) 
        self._ppt.replace_text(f"{{app{app_no+1}_{priority}_violation_text}}",self.list_violations(filtered)) 
        days_effort = math.ceil(filtered['Days Effort'].sum())
        cost_effort = (days_effort*600)/1000
        self._ppt.replace_text(f"{{app{app_no+1}_{priority}_cost}}",f'~${cost_effort}K-${cost_effort*2}K') 
        self._ppt.replace_text(f"{{app{app_no+1}_{priority}_days}}",f'~{days_effort}-{days_effort*2}')

    def list_violations(self,filtered):
        first = True
        text = ""
        for criteria in filtered['Technical Criteria'].unique():
            df = filtered[filtered['Technical Criteria']==criteria]
            total = df['No. of Actions'].sum()
            
            cases = 'for'
            if first:
                cases = 'cases of'
                first = False
            
            rule = criteria[criteria.find('-')+1:].strip().lower()
            if len(rule) == 0:
                rule = criteria
            text = f'{text}{total} {cases} {rule}, '
        return util.rreplace(text[:-2],', ',' and ')

    def common_business_criteria(self,summary_df,priority,default=''):
        filtered=summary_df[summary_df['tag']==priority]
        count = 0
        sum = 0
        list = []
        if not filtered.empty:
            sum = filtered['No. of Actions'].sum()
            for business in filtered['Business Criteria']:
                items = business.split(',')
                for t in items:
                    if t.strip() not in list:
                        list.append(t.strip())

        if sum==0:
            sum_txt = 'zero'
        else:
            sum_txt = str(sum)
        
        if not list:
            list.append(default)

        return util.list_to_text(list),sum_txt, filtered

    def fill_violations(self,app_no):
        app_id = self._app_list[app_no]
        violation_df = pd.DataFrame(self._aip_data.violation_sizing(app_id),index=[0])
        violation_df['Violation Count']=pd.Series(["{0:.0f}".format(val) for val in violation_df['Violation Count']])
        violation_df[' per file']=pd.Series(["{0:.2f}".format(val) for val in violation_df[' per file']])
        violation_df[' per kLoC']=pd.Series(["{0:.2f}".format(val) for val in violation_df[' per kLoC']])
        violation_df['Complex objects']=pd.Series(["{0:.0f}".format(val) for val in violation_df['Complex objects']])
        violation_df[' With violations']=pd.Series(["{0:.0f}".format(val) for val in violation_df[' With violations']])
        self._ppt.update_table(f'app{app_no+1}_violation_sizing',violation_df.transpose())

    def fill_sizing(self,app_no):
        app_id = self._app_list[app_no]

        sizing_df = pd.DataFrame(self._aip_data.tech_sizing(app_id),index=[0])
        sizing_df['LoC']=pd.Series(["{0:.0f} K".format(val / 1000) for val in sizing_df['LoC']])
        sizing_df['Files']=pd.Series(["{0:.0f}".format(val) for val in sizing_df['Files']])
        sizing_df['Classes']=pd.Series(["{0:.0f}".format(val) for val in sizing_df['Classes']])
        sizing_df['SQL Artifacts']=pd.Series(["{0:.0f}".format(val) for val in sizing_df['SQL Artifacts']])
        sizing_df['Tables']=pd.Series(["{0:.0f}".format(val) for val in sizing_df['Tables']])
        sizing_df = sizing_df.transpose()
        self._ppt.update_table(f'app{app_no+1}_tech_sizing',sizing_df)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Assessment Deck Generation Tool')
    parser.add_argument('-c','--config', required=True, help='Configuration properties file')
    args = parser.parse_args()

    config = Properties()
    with open(args.config, 'rb') as config_file:
        config.load(config_file)
        project_name = config.get('project').data
        template_name = config.get('template').data
        app_list = config.get('application.list').data.strip().split(',')
        output_folder = config.get('output.folder').data

    GeneratePPT(project_name,app_list,template_name,output_folder)


