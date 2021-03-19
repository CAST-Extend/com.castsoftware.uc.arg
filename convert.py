from restCall import AipRestCall
from restCall import AipData
from pptx import Presentation
from powerpoint import PowerPoint
from jproperties import Properties 
from pptx.dml.color import RGBColor
from actionPlan import ActionPlan

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

    def __init__(self,config):
        project = config.get('project').data
        template = config.get('template').data
        self._app_list = config.get('application.list').data.strip().split(',')
        self._output_folder = config.get('output.folder').data

        out = f"{self._output_folder}/{project}.pptx"
        print (f'Generating {out}')

        self._rest = AipRestCall("http://sha-dd-console:8080/CAST-RESTAPI-integrated/rest/","cast","cast",True)
        self._aip_data = AipData(self._rest,project, self._app_list)

        self._ppt = PowerPoint(template,out)

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
        self._ppt.replace_text("{summary_at_risk_factors}",util.list_to_text(risk_grades.index.tolist()).lower())

        
        if risk_grades.empty:
            self._ppt.replace_block('{immediate_action}','{end_immediate_action}','')
        else:
            self._ppt.replace_text("{immediate_action}","") 
            self._ppt.replace_text("{end_immediate_action}","") 
            self._ppt.replace_text("{high_risk_grade_names}",self._aip_data.text_from_list(risk_grades.index.values.tolist()))
                
        # create instance of action plan class 
        ap = ActionPlan (self._app_list,self._aip_data,self._ppt,self._output_folder)

        for app_no in range(0,app_cnt):
            app_id = self._app_list[app_no]
            if self._aip_data.has_data(app_id):
                print (f'Working on {app_id}')

                # do risk factors for the executive summary page
                risk_grades = util.each_risk_factor(self._ppt, self._aip_data,app_id, app_no)
                self._ppt.replace_text(f'{{app{app_no+1}_high_risk_grade_names}}',util.list_to_text(risk_grades.index.values))

                grade_all = self._aip_data.get_app_grades(app_id)
                self._ppt.replace_risk_factor(grade_all,app_no)
                grade_by_tech_df = self._aip_data.get_grade_by_tech(app_id)

                self._ppt.update_table(f'app{app_no+1}_grade_by_tech_table',grade_by_tech_df.drop(['Documentation'],axis=1))
                self._ppt.update_chart(f'app{app_no+1}_sizing_pie_chart',grade_by_tech_df['LOC'])

                snapshot = self._aip_data.snapshot(app=app_id)
                app_name = snapshot['name']
                self._ppt.replace_text(f'{{app{app_no+1}_name}}',app_name)
                self._ppt.replace_text(f'{{app{app_no+1}_all_technogies}}',util.list_to_text(snapshot['technology']))

                #calculate high and medium risk factors
                risk_grades = self._aip_data.calc_health_grades_high_risk(grade_all)
                if risk_grades.empty:
                    risk_grades = self._aip_data.calc_health_grades_medium_risk(grade_all)
                self._ppt.replace_text(f'{{app{app_no+1}_at_risk_grade_names}}',util.list_to_text(risk_grades.index.tolist()).lower())

                loc_df = self._aip_data.get_loc_sizing(app_id)
                loc = loc_df['Number of Code Lines']
                self._ppt.replace_loc(loc,app_no+1)

                loc_tbl = pd.DataFrame.from_dict(data=self._aip_data.get_loc_sizing(app_id),orient='index').drop('Critical Violations')
                loc_tbl = loc_tbl.rename(columns={0:'loc'})
                loc_tbl['percent'] = round((loc_tbl['loc'] / loc_tbl['loc'].sum()) * 100,2)
                loc_tbl['loc']=pd.Series(["{0:,.0f}".format(val) for val in loc_tbl['loc']], index = loc_tbl.index)
                loc_tbl['percent']=pd.Series(["{0:.2f}%".format(val) for val in loc_tbl['percent']], index = loc_tbl.index)
                self._ppt.update_table(f'app{app_no+1}_loc_table',loc_tbl,has_header=False)
                self._ppt.update_chart(f'app{app_no+1}_loc_pie_chart',loc_tbl['loc'])

                self._ppt.replace_grade(grade_all,app_no+1)

                self.fill_sizing(app_no)
                self.fill_critical_rules(app_no)
                ap.fill_action_plan(app_no)
                self.fill_violations(app_no)
            else:
                print (f'No snapshot available for {app_id}')

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



    def fill_violations(self,app_no):
        app_id = self._app_list[app_no]
        violation_df = pd.DataFrame(self._aip_data.violation_sizing(app_id),index=[0])
        violation_df['Violation Count']=pd.Series(["{0:,.0f}".format(val) for val in violation_df['Violation Count']])
        violation_df[' per file']=pd.Series(["{0:,.2f}".format(val) for val in violation_df[' per file']])
        violation_df[' per kLoC']=pd.Series(["{0:,.2f}".format(val) for val in violation_df[' per kLoC']])
        violation_df['Complex objects']=pd.Series(["{0:,.0f}".format(val) for val in violation_df['Complex objects']])
        violation_df[' With violations']=pd.Series(["{0:,.0f}".format(val) for val in violation_df[' With violations']])
        self._ppt.update_table(f'app{app_no+1}_violation_sizing',violation_df.transpose())
        self._ppt.replace_text(f'{{app{app_no+1}_critical_violations}}',violation_df['Violation Count'].loc[0])


    def fill_sizing(self,app_no):
        app_id = self._app_list[app_no]

        sizing_df = pd.DataFrame(self._aip_data.tech_sizing(app_id),index=[0])
        sizing_df['LoC']=pd.Series(["{0:,.0f} K".format(val / 1000) for val in sizing_df['LoC']])
        sizing_df['Files']=pd.Series(["{0:,.0f}".format(val) for val in sizing_df['Files']])
        sizing_df['Classes']=pd.Series(["{0:,.0f}".format(val) for val in sizing_df['Classes']])
        sizing_df['SQL Artifacts']=pd.Series(["{0:,.0f}".format(val) for val in sizing_df['SQL Artifacts']])
        sizing_df['Tables']=pd.Series(["{0:,.0f}".format(val) for val in sizing_df['Tables']])
        sizing_df = sizing_df.transpose()
        self._ppt.update_table(f'app{app_no+1}_tech_sizing',sizing_df)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Assessment Deck Generation Tool')
    parser.add_argument('-c','--config', required=True, help='Configuration properties file')
    args = parser.parse_args()

    config = Properties()
    with open(args.config, 'rb') as config_file:
        config.load(config_file)

    GeneratePPT(config)


