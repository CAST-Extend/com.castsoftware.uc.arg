from pandas.core.frame import DataFrame
import requests
import pandas as pd
import logging
import enum
import urllib.parse
import numpy as np 

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

    def get(self, url = ""):
        start_dttm = ctime()
        start_tm = perf_counter()

        # TODO: Errorhandling
        
        headers = {'Accept': 'application/json'}
        #u = f'{self._base_url}{url}'
        u = urllib.parse.quote(f'{self._base_url}{url}',safe='/:&?=')
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

        return resp.status_code, resp.json()

    def get2(self, url = ""):
        start_dttm = ctime()
        start_tm = perf_counter()

        # TODO: Errorhandling
        
        headers = {'Accept': '*/*'}
        #u = f'{self._base_url}{url}'
        u = urllib.parse.quote(f'{self._base_url}{url}',safe='/:&?=')
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

        return resp.status_code, resp.json()


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

    def get_domain(self, schema_name):
        domain_id = None
        (status,json) = self.get()
        if status == requests.codes.ok:
            try: 
                domain_id = list(filter(lambda x:x["schema"].lower()==schema_name.lower(),json))[0]['name']
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
                if self._data[s]['snapshot']:
                    self._data[s]['has data'] = True
                    self._data[s]['grades']=rest.get_grades_by_technology(domain_id,self._data[s]['snapshot'])
                    self._data[s]['sizing']=rest.get_sizing_by_technology(domain_id,self._data[s]['snapshot'],self._sizing)
                    self._data[s]['loc_sizing']=rest.get_sizing(domain_id,self._sizing) 
                    self._data[s]['tech_sizing']=rest.get_sizing(domain_id,self._tech_sizing) 
                    self._data[s]['violation_sizing']=rest.get_violation_CR(domain_id)
                    self._data[s]['critical_rules']=rest.get_rules(domain_id,self._data[s]['snapshot']['id'],60017,non_critical=False)

                    (ap_df,ap_summary_df) = rest.get_action_plan(domain_id,self._data[s]['snapshot']['id']) 
                    self._data[s]['action_plan']=ap_df
                    self._data[s]['action_plan_summary']=ap_summary_df
                else:
                    self._data[s]['has data'] = False

    def has_data(self, app):
        return self._data[app]['has data']

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
            if self.has_data(row):
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
            if self.has_data(a):
                rslt = rslt + self.snapshot(a)['name']
                if l >= 2 and a == last_name:
                    rslt = rslt + " and "
                elif a != self._base[-1]:
                    rslt = rslt + ", "
            else:
                rslt = "NO SNAPSHOT INFORMATION AVAILABLE"
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

class HLRestCall(RestCall):
    """
    Class to handle HL REST API calls.
    """
    def __init__(self, hl_base_url, hl_user, hl_pswd, hl_instance, timer_on):
        super(HLRestCall, self).__init__(hl_base_url, hl_user, hl_pswd, timer_on)

        self._hl_instance = hl_instance
        self._hl_data_retrieved = False
    
    def _get_app_ids(self, instance_id):
        # Reeetrieve the HL app id for the application.

        try:
            # TODO: remove the hard-coding
            # TODO: Get the app id.
            url = f'domains/{instance_id}/applications'

            (status, json) = self.get2(url)

            # TODO: Handle exceptions
            if status == requests.codes.ok and len(json) > 0:
                pass
        except:
            # TODO
            print('Oopsi.. caught an exception')
            raise

        return json

    def _get_third_party(self, app_id):
        try:
            # TODO: remove the hard-coding
            # TODO: Get the app id.
            url = f'domains/{self._hl_instance}/applications/{app_id}/thirdparty'

            (status, json) = self.get(url)

            # TODO: Handle exceptions
            if status == requests.codes.ok and len(json) > 0:
                # TODO: TEMP
                for i in range(len(json['thirdParties'])):
                    # Fill-in the blanks. Not every third party entry has CVE and license info.
                    # If not found, create a blank entry for ease in moving the data over to a dataframe.

                    json['thirdParties'][i]['cve'] = json['thirdParties'][i].get('cve', np.nan)
                    json['thirdParties'][i]['licenses'] = json['thirdParties'][i].get('licenses', np.nan)

                    #print('j:', j)
                    #print('json:', json['thirdParties'][i]['cve'], json['thirdParties'][i]['licenses'])

                return json
        except:
            # TODO
            print('Oopsi.. caught an exception')
            raise

class HLData(HLRestCall):
    """
    """
    def __init__(self, rest, project, schema):
        self._rest = rest
        self._base = schema
        self._data_retrieved = False
        self._app_id = None
        self._got_data = False
        self._has_crit_sev_cves = False
        self._has_high_sev_cves = False
        self._has_med_sev_cves = False
        self._has_high_risk_lics = False
        self._has_med_risk_lics = False
        self._third_party_df = pd.DataFrame()
        self._cve_df = pd.DataFrame()
        self._lic_df = pd.DataFrame()

    def get_app_ids(self, instance_id):
        # TODO: try-except
        return self._rest._get_app_ids(instance_id)

    def _get_third_party(self, app_id):

        if self._app_id != app_id:
            self._app_id = app_id
            self._got_data = False

        if not self._got_data:
            # If we do not have the data already for this app, retrieve it first.

            self._third_party_df = pd.DataFrame(self._rest._get_third_party(app_id)['thirdParties'])
            # TODO: Confirm that data was retrieved before setting this to True
            self._got_data = True

            self._cve_df = self._third_party_df.loc[:, ['name', 'cve']]
            self._cve_df.dropna(axis = 0, how = 'any', inplace = True)
            self._lic_df = self._third_party_df.loc[:, ['name', 'licenses']]
            self._lic_df.dropna(axis = 0, how = 'all', inplace = True)

        return

    def get_cves(self, app_id, type, limit = 0):
        """
        Returns CVE info for a give app, if there are Critical, High and medium CVEs.
        Unless the 'all' argument is provided, High CVEs are returned only when there are no/not enough critical CVEs
        to fill the slide. Similarly, medium CVEs are returned only where there are no/limited critical or high CVEs.

        Note that low CVEs are ignored for assessment purposes.
        """

        _cve_df = pd.DataFrame()
        sev_type = type.lower()

        # If the request is for critical sev CVEs, override the limit and return all rows.
        if sev_type == 'critical':
            limit = 0

        # Do we have the data retrieved for the app? If not, auto-retrieve.
        # If all == False, limit the number of CVEs returned.

        try:
            self._get_third_party(app_id)
        except:
            print('ERROR - no thirdparty data')
            raise

        i = 0
        prev_comp = ''
        comp_changed = False

        for tup in self._cve_df.itertuples():
            crit_cve_str, high_cve_str, med_cve_str = '', '', ''
            comp_name = tup[1]

            for cve in tup[2]['vulnerabilities']:
                # If critical sev is the type requested, also grab high and medium sev CVEs for the same component.
                # Similarly, if high sev is the type requested, also grab medium sev CVEs for the same component.
                criticality = cve['criticity'].lower()

                # Ignore components with low criticality, as we do not need to report those.
                if criticality == 'low':
                    continue

                # Note that are not considering the type requested here, but simply storing the CVE in its
                # appropriate slot. That part is handled further below, before they are added to the dataframe.
                if criticality == 'critical':
                    crit_cve_str += cve['name'] + ', '
                if criticality == 'high':
                    high_cve_str += cve['name'] + ', '
                elif criticality == 'medium':
                    med_cve_str += cve['name'] + ', '

            # Before adding the results into the dataframe, ensure that we have matches for the
            # crticality requested. For example, if critical CVEs were requested, we should have
            # values in the crit_cve_str. # If not, do not add.
            if sev_type == 'critical' and len(crit_cve_str) > 0 or \
                sev_type == 'high' and len(high_cve_str) > 0    or \
                sev_type == 'medium' and len(med_cve_str) > 0:

                # Strip the comma-spaces at the end of the CVE strings.
                crit_cve_str = crit_cve_str[:-2]
                high_cve_str = high_cve_str[:-2]
                med_cve_str = med_cve_str[:-2]

                if len(crit_cve_str) == 0:
                    crit_cve_str = 'N/A'
                if len(high_cve_str) == 0:
                    high_cve_str = 'N/A'
                if len(med_cve_str) == 0:
                    med_cve_str = 'N/A'

                _cve_df = _cve_df.append({'Component Name': comp_name, 'Critical Sev CVEs': crit_cve_str,
                            'High Sev CVEs': high_cve_str, 'Medium Sev CVEs': med_cve_str}, ignore_index = True)

            i += 1

            # Return only the max number of rows requested. In case the sev type is critical, always return all rows.
            if (limit > 0) and (i == limit):
                break

        #print(_cve_df)
        return _cve_df

    def get_lics(self, app_id, type = 'high', limit = 0):
        """
        """

        compliance = ''
        _lic = []
        _lic_df = pd.DataFrame()

        # Do we have the data retrieved for the app? If not, auto-retrieve.
        # If all == False, limit the number of CVEs returned.

        try:
            self._get_third_party(app_id)
        except:
            print('ERROR - no thirdparty data')
            raise

        i = 0
        for tup in self._lic_df.itertuples():
            high_risk_lic_str = ''
            med_risk_lic_str = ''
            comp_name = tup[1]

            for dic in tup[2]:
                if type == 'high':
                    # A note on risk-complaince relation.
                    # A riskier license translates to lower compliance. The REST API returns compliance info.
                    # So, if the type being sought is high-risk license, we return low compliance license.
                    #
                    # If tbe requested type is high risk licenses and we also see
                    # medium risk licenses for the same component, include those as well.
                    # Leveraging the content of high_risk_lic_str for this purpose.

                    if dic['compliance'] == 'low':
                        high_risk_lic_str += dic['name'] + '\n'
                    elif dic['compliance'] == 'medium' and len(high_risk_lic_str) > 0:
                        med_risk_lic_str += dic['name'] + '\n'
                elif type == 'medium':
                    # A medium risk likcense has a compliance of medium.
                    if dic['compliance'] == 'medium': 
                        med_risk_lic_str += dic['name'] + '\n'
            
            if len(high_risk_lic_str) > 0 or len(med_risk_lic_str) > 0:
                # Strip the trailing newline.
                high_risk_lic_str = high_risk_lic_str[:-1]
                med_risk_lic_str = med_risk_lic_str[:-1]

                if len(high_risk_lic_str) == 0:
                    high_risk_lic_str = 'N/A'

                if len(med_risk_lic_str) == 0:
                    med_risk_lic_str = 'N/A'

                _lic_df = _lic_df.append({'Component Name': comp_name, 'High Risk': high_risk_lic_str,
                            'Medium Risk': med_risk_lic_str}, ignore_index = True)
                i += 1

            if (limit > 0) and (i == limit):
                break

        """
        for elem in _lics:
            print(elem.key, elem.value)
        for idx in range(len(_lics)):
            for key in _lics[idx]:
                print(key, _lics[idx][key])

            #for j in range(len(high_risk_lic_df['LicenseInfo'][i])):
                #print(high_risk_lic_df['LicenseInfo'][i][j])
                #_lic_str += high_risk_lic_df['name'][i][j] + '\n'

            #_lic_str = _lic_str[:-1]
            #_lics = _lics.append({'name:', high_risk_lic_df['licenses'], 'licenses:', _lic_str})
        """
        #print(_lic_df)
        return _lic_df

"""
apps = ["actionplatform","intersect"] 
aip_data = AipData(aip_rest,"Florence", apps)

app_id = apps[0]
grade_all = aip_data.get_app_grades(app_id)
text = aip_data.get_grade_at_risk_text(grade_all)
"""