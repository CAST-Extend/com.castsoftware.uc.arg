from pandas.core.frame import DataFrame
import requests
import pandas as pd
import logging
import enum
import urllib.parse

from requests.auth import HTTPBasicAuth 
from time import perf_counter, ctime
from copy import copy

class RestCall:

    _base_url = None
    _auth = None
    _time_tracker_df  = pd.DataFrame()
    _track_time = True

    def __init__(self, base_url, user, password, track_time=False):
        self._logger = logging.getLogger(__name__)
        shandler = logging.StreamHandler()
        formatter = logging.Formatter('%(asctime)s - %(filename)s [%(funcName)30s:%(lineno)-4d] %(levelname)-8s - %(message)s')
        shandler.setFormatter(formatter)
        self._logger.addHandler(shandler)

        self._base_url = base_url
        self._track_time = track_time
        self._auth = HTTPBasicAuth(user, password)
        self.setLoggingLevel(logging.INFO)

    def setLoggingLevel(self,level):
        self._logger.setLevel(level)

    def get(self,url=""):
        start_dttm = ctime()
        start_tm = perf_counter()

        # TODO: Errorhandling
        
        headers = {'Accept': 'application/json'}
        u = f'{self._base_url}{url}'
        resp = requests.get(url= u, auth = self._auth, headers=headers)

        # Save the duration, if enabled.
        if (self._track_time):
            end_tm = perf_counter()
            end_dttm = ctime()
            duration = end_tm - start_tm

            #print(f'Request completed in {duration} ms')
            self._time_tracker_df = self._time_tracker_df.append({'Application': 'ALL', 'URL': url, 'Start Time': start_dttm \
                                                        , 'End Time': end_dttm, 'Duration': duration}, ignore_index=True)
        
        # TODO: Errorhandling

        return resp.status_code,resp.json()

#class HLRestCall(RestCall):


class AipRestCall(RestCall):
    _measures = {
        '60017':'TQI',
        '60013':'Robustness',
        '60014':'Efficiency',
        '60016':'Security',
        '60011':'Transferability',
        '60012':'Changeability',
        '60015':'SEI Maintainability',
        '66033':'Documentation'
    }

    _violations = {
        '67011':'Violation Count',
        '67012':' per file',
        '67013':' per kLoC'
    }

    def get_domain(self,schema_name):
        domain_id = None
        (status,json) = self.get()
        if status == requests.codes.ok:
            try: 
                domain_id = list(filter(lambda x:x["schema"]==schema_name,json))[0]['name']
            except IndexError:
                self._logger.error(f'Domain not found for schema {schema_name}')
                
        return domain_id

    def get_latest_snapshot(self,domain_id):
        snapshot = {}
        (status,json) = self.get(f'{domain_id}/applications/3/snapshots')
        if status == requests.codes.ok and len(json) > 0:
            snapshot['id'] = json[0]['href'].split('/')[-1]  
            snapshot['name'] = json[0]['name']
            snapshot['technology'] = json[0]['technologies']
            snapshot['module_href'] = json[0]['moduleSnapshots']['href']
            snapshot['result_href'] = json[0]['results']['href'] 
        return snapshot 

    def get_grades_by_technology(self,domain_id,snapshot):
        first_tech=True
        grade = pd.DataFrame(columns=list(self._measures.values()))
        for tech in snapshot['technology']:
            t={}
            a={}
            for key in self._measures: 
                url = f'{domain_id}/applications/3/results?quality-indicators={key}&technologies={tech}'
                (status,json) = self.get(url)
                if status == requests.codes.ok and len(json) > 0:
                    try:
                        t[self._measures[key]]=json[0]['applicationResults'][0]['technologyResults'][0]['result']['grade']
                    except IndexError:
                        self._logger.warning(f'{domain_id} no grade available for {key} {tech} setting it to 4')
                        t[self._measures[key]]=4

                    if first_tech==True:
                        a[self._measures[key]]=json[0]['applicationResults'][0]['result']['grade']
                else:
                    self._logger.error (f'Error retrieving technology information:  {url}')
            if first_tech==True:
                grade.loc['All'] = a
            grade.loc[tech] = t
            first_tech=False
        return grade

    def get_sizing_by_technology(self,domain_id,snapshot,sizing):
        first_tech=True
        size_df = pd.DataFrame(columns=list(sizing.values()))
        for tech in snapshot['technology']:
            t={}
            a={}
            for key in sizing: 
                url = f'{domain_id}/applications/3/results?sizing-measures={key}&technologies={tech}'
                (status,json) = self.get(url)
                if status == requests.codes.ok and len(json) > 0:
                    try:
                        t[sizing[key]]= json[0]['applicationResults'][0]['technologyResults'][0]['result']['value']
                        if first_tech==True:
                            a[sizing[key]]=json[0]['applicationResults'][0]['result']['value']
                    except IndexError:
                        self._logger.debug(f'{domain_id} no grade available for {key} {tech}')
            if first_tech==True:
                size_df.loc['All'] = a
            size_df.loc[tech] = t
            first_tech=False
        return size_df

    def get_distribution_sizing(self, domain_id, metric_id):
        rslt = DataFrame(columns=['name','value'])
        (status,json) = self.get(f'{domain_id}/applications/3/results?metrics={metric_id}&select=categories')
        if status == requests.codes.ok and len(json) > 0:
            cat = json[0]['applicationResults'][0]['result']['categories']
            for index, name in enumerate(cat):
                rslt.loc[name['key']]=[[name['name']],[name['value']]]

        return rslt

    def get_rules(self,domain_id,snapshot_id,business_criteria,critical=True,non_critical=True,start_row=1,max_rows=10000):
        rslt_df =  pd.DataFrame()
        critical_arg=non_critical_arg=''

        if critical:
           critical_arg=f'cc:{business_criteria}' 
        if non_critical:
           non_critical_arg=f'nc:{business_criteria}' 

        rule_arg=critical_arg
        if len(rule_arg) > 0:
            rule_arg = rule_arg + ','
        rule_arg=f'{rule_arg}{non_critical_arg}'

        url = f'{domain_id}/applications/3/snapshots/{snapshot_id}/violations?rule-pattern={rule_arg}&startRow={start_row}&nbRows={max_rows}'
        (status,json) = self.get(url)
        if status == requests.codes.ok and len(json) > 0:
            rslt_df = pd.DataFrame(json)
        return rslt_df


    def get_action_plan(self,domain_id,snapshot_id):
        business_criteria = ['Robustness','Efficiency','Security','Transferability','Changeability']
    
        catagory = ''
        tech_criteria = ''
        rslt_df =  pd.DataFrame()
        ap_summary_df =  pd.DataFrame()
        url = f'{domain_id}/applications/3/snapshots/{snapshot_id}/action-plan/issues?startRow=1&nbRows=100000'
        (status,json) = self.get(url)
        if status == requests.codes.ok and len(json) > 0:
            rslt_df = pd.DataFrame(json)
            rule_pattern = pd.json_normalize(rslt_df['rulePattern']).add_prefix('rule.')
            rule_pattern['rule.href'] = rule_pattern['rule.href'].str.split('/').str[-1]
            rule_pattern = rule_pattern.rename(columns={"rule.href":"rule.id"})

            component = pd.json_normalize(rslt_df['component']).add_prefix('component.') 
            component.drop(component.columns.difference(['component.name','component.shortName']),1,inplace=True)

            remediation = pd.json_normalize(rslt_df['remedialAction']) 
            rslt_df = rule_pattern.join([component,remediation])                                                  

            rslt_df.insert(3,'Business Criteria','')
            rslt_df.insert(3,'tech_criteria','')

            save_rule_id = ''
            for key, value in rslt_df.iterrows():
                rule_id=value['rule.id']
                if save_rule_id != rule_id:
                    save_rule_id = rule_id
                    url = f'{domain_id}/quality-indicators/{rule_id}/snapshots/{snapshot_id}'
                    (status,json) = self.get(url)
                    if status == requests.codes.ok and len(json) > 0:
                        catagory = ''
                        tech_criteria = ''
                        for g1 in json['gradeAggregators']:
                            tech_criteria = g1['name']
                            for g2 in g1['gradeAggregators']:
                                if g2['name'] in business_criteria:
                                    catagory = catagory + g2['name'] + ', '
                
                rslt_df.loc[key,'tech_criteria']=tech_criteria
                rslt_df.loc[key,'Business Criteria']=catagory[:-2]

            rslt_df = rslt_df.sort_values(by=['rule.id'])
            ap_summary_df = rslt_df.groupby(['rule.name']).count()
            business = pd.DataFrame(rslt_df,columns=['rule.name','tech_criteria','Business Criteria','tag','comment']).drop_duplicates()
            ap_summary_df.drop(ap_summary_df.columns.difference(['rule.name','component.name']),1,inplace=True)
            ap_summary_df = pd.merge(ap_summary_df,business, on='rule.name')
            ap_summary_df = ap_summary_df[['rule.name','Business Criteria','component.name','comment','tag','tech_criteria']]
            ap_summary_df = ap_summary_df.rename(columns={'component.name':'No. of Actions',
                                                          'rule.name':'Quality Rule',
                                                          'tech_criteria':'Technical Criteria'
                                                          })

            rslt_df = rslt_df.rename(columns={'rule.name':'Rule Name',
                                              'comment':'Action Plan Priority',
                                              'component.name':'Object Name Location'})
            rslt_df = rslt_df[['Action Plan Priority','Rule Name','Object Name Location','rule.id']]
        return (rslt_df, ap_summary_df)

    def getLOC(self,domain_id):
        loc = 0
        (status,json) = self.get(f'{domain_id}/applications/3/results?sizing-measures=10151&snapshots=-1')
        if status == requests.codes.ok and len(json) > 0:
            loc = json[0]['applicationResults'][0]['result']['value']
        return loc

    def get_sizing(self, domain_id, input):
        rslt = {}
        for key in input: 
            (status,json) = self.get(f'{domain_id}/applications/3/results?sizing-measures={key}&snapshots=-1')
            if status == requests.codes.ok and len(json) > 0:
                rslt[input[key]]=json[0]['applicationResults'][0]['result']['value']
        return rslt

    def get_violation_CR(self,domain_id):
        vs = self.get_sizing(domain_id,self._violations) 
        complexity = self.get_distribution_sizing(domain_id,'67001')
        vs['Complex objects']=complexity.loc['67002']['value'][0]+complexity.loc['67003']['value'][0]
        complexity = self.get_distribution_sizing(domain_id,'67030')
        vs[' With violations']=complexity.loc['67031']['value'][0]+complexity.loc['67032']['value'][0]
        return vs


class AipData():
    _data={}
    _base=[]
    _rest=None

    _sizing = {
       '10151':'Number of Code Lines', 
       '10107':'Number of Comment Lines', 
       '10109':'Number of Commented-out Code Lines',
       '67011':'Critical Violations'
    }

    _tech_sizing = {
        '10151':'LoC',
        '10154':'Files',
        '10155':'Classes',
        '10158':'SQL Artifacts',
        '10163':'Tables'
    }

    _health_grade_ids = ['Efficiency','Robustness','Security','Changeability','Transferability']

    def __init__(self, rest, project, schema):
        self._rest=rest
        self._base=schema
        for s in schema:
            print (f'Collecting data for {s}')
            self._data[s]={}
            central_schema = f'{s}_central'
            domain_id = rest.get_domain(central_schema)
            if domain_id is not None:
                self._data[s]['domain_id']=domain_id
                self._data[s]['snapshot']=rest.get_latest_snapshot(domain_id)
                self._data[s]['grades']=rest.get_grades_by_technology(domain_id,self._data[s]['snapshot'])
                self._data[s]['sizing']=rest.get_sizing_by_technology(domain_id,self._data[s]['snapshot'],self._sizing)
                self._data[s]['loc_sizing']=rest.get_sizing(domain_id,self._sizing) 
                self._data[s]['tech_sizing']=rest.get_sizing(domain_id,self._tech_sizing) 
                self._data[s]['violation_sizing']=rest.get_violation_CR(domain_id)
                self._data[s]['critical_rules']=rest.get_rules(domain_id,self._data[s]['snapshot']['id'],60017,non_critical=False)

                (ap_df,ap_summary_df) = rest.get_action_plan(domain_id,self._data[s]['snapshot']['id']) 
                self._data[s]['action_plan']=ap_df
                self._data[s]['action_plan_summary']=ap_summary_df

    def data(self,app):
        return self._data[app]

    def domain(self, app):
        return self.data(app)['domain_id']

    def snapshot(self, app):
        return self.data(app)['snapshot']

    def grades(self, app):
        return self.data(app)['grades']

    def sizing(self, app):
        return self.data(app)['sizing']

    def critical_rules(self, app):
        return self.data(app)['critical_rules']

    def action_plan(self, app):
        ap_df = self.data(app)['action_plan']
        ap_summary_df = self.data(app)['action_plan_summary']
        return (ap_df,ap_summary_df)
        
    def get_app_grades(self, app, sort=False):
        app_grades = self.grades(app).loc['All']
        if sort:
            return app_grades
        else:
            return app_grades.sort_values()

    def calc_grades_all_apps(self):
        all_app=pd.DataFrame()
        for row in self._data:
            app_name=self.snapshot(row)['name']
            grades=self.grades(row)
            all_app = pd.concat([all_app,grades[grades.index.isin(['All'])].rename(index={'All': app_name})]).drop_duplicates()
        return all_app[all_app.columns].mean(axis=0)

    def calc_grades_health(self,grade_all):
        grade_df = pd.DataFrame(grade_all)
        grade_health=grade_df[grade_df.index.isin(self._health_grade_ids)]
        return grade_health

    def calc_health_grades_high_risk(self,grade_all):
        grade_health = self.calc_grades_health(grade_all)
        grade_at_risk=grade_health[grade_health < 2.5].dropna()
        return grade_at_risk

    def calc_health_grades_medium_risk(self,grade_all):
        grade_health = self.calc_grades_health(grade_all)
        grade_at_risk=grade_health[grade_health > 2.5].dropna()
        grade_at_risk=grade_health[grade_health < 3].dropna()
        return grade_at_risk

    def get_loc_sizing(self,app):
        return self.data(app)['loc_sizing']

    def tech_sizing(self, app):
        return self.data(app)['tech_sizing']

    def violation_sizing(self, app):
        return self.data(app)['violation_sizing']

    def get_all_app_text(self):
        rslt = ""

        data = self._data
        l = len(self._base)
        if l == 1:
            return self.snapshot(self._base[0])['name']

        last_name = self._base[-2]
        for a in self._base:
            rslt = rslt + self.snapshot(a)['name']
            if l >= 2 and a == last_name:
                rslt = rslt + " and "
            elif a != self._base[-1]:
                rslt = rslt + ", "
        return rslt

    def get_grade_by_tech(self,app):
        grade_df = self.grades(app).round(2).applymap('{:,.2f}'.format)
        grade_df = grade_df[grade_df.index.isin(['All'])==False]

        sizing_df = pd.DataFrame(self.sizing(app))
        sizing_df = sizing_df[sizing_df.index.isin(['All'])==False]
        sizing_df = pd.DataFrame(sizing_df["Number of Code Lines"].rename("LOC")).dropna()
        sizing_df = sizing_df.applymap('{:,.0f}'.format)
        
        tech = sizing_df.join(grade_df) 

        sizing_df = pd.DataFrame(self.sizing(app)) 
        sizing_df = sizing_df[sizing_df.index.isin(['All'])==False]
        sizing_df = pd.DataFrame(sizing_df["Critical Violations"]).dropna()
        sizing_df = sizing_df.applymap('{:,.0f}'.format)
        tech = tech.join(sizing_df)

        return tech

    def get_high_risk_grade_text(self, grades):
        grade_at_risk=self.calc_health_grades_high_risk(grades)
        if grade_at_risk.empty:
            return None
        else:
            return self.text_from_list(grade_at_risk.index.values.tolist())

    def get_medium_risk_grade_text(self, grades):
        grade_at_risk=self.calc_health_grades_medium_risk(grades)
        if grade_at_risk.empty:
            return None
        else:
            return self.text_from_list(grade_at_risk.index.values.tolist())

    def text_from_list(self,list):
        rslt = ""
        total_items = len(list)
        if total_items == 1:
            return list[0]

        last_item = list[-1]
        sec_last_item = list[-2]
        for name in list:
            rslt = rslt + name
            if total_items >= 2 and name == sec_last_item:
                rslt = rslt + " and "
            elif name != last_item:
                rslt = rslt + ", "
        return rslt



"""
apps = ["actionplatform","intersect"] 
aip_data = AipData(aip_rest,"Florence", apps)

app_id = apps[0]
grade_all = aip_data.get_app_grades(app_id)
text = aip_data.get_grade_at_risk_text(grade_all)
"""



