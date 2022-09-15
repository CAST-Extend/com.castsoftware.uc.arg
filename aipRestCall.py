from restAPI import RestCall
from requests import codes

from pandas import DataFrame
from pandas import merge
from pandas import json_normalize

class AipRestCall(RestCall):
    _measures = {
        '60017':'TQI',
        '60013':'Robustness',
        '60014':'Efficiency',
        '60016':'Security',
        '60011':'Transferability',
        '60012':'Changeability',
#        '60015':'SEI Maintainability',
        '66033':'Documentation',
        '1061000':"ISO",
        '1061003':"ISO_REL",
        '1061001':"ISO_MAINT",
        '1061002':"ISO_EFF",
        '1061004':"ISO_SEC"
    }

    _violations = {
        '67011':'Violation Count',
        '67012':' per file',
        '67013':' per kLoC'
    }

    def get_domain(self, schema_name):
        domain_id = None
        (status,json) = self.get()
        if status == codes.ok:
            try: 
                domain_id = list(filter(lambda x:x["schema"].lower()==schema_name.lower(),json))[0]['name']
            except IndexError:
                self.error(f'Domain not found for schema {schema_name}')
        if status == 0:
            domain_id = -1        
        return domain_id

    def get_quality_indicators(self,domain_id,snapshot_id, key):
        url = f'{domain_id}/quality-indicators/{key}/snapshots/{snapshot_id}'

        (status,json) = self.get(url)
        if status == codes.ok and len(json) > 0:
            return json['gradeContributors']
        else:
            return None

    def get_violation_ratio(self,domain_id, key):
        url = f'{domain_id}/applications/3/results?quality-indicators=cc:{key},nc:{key}&select=violationRatio'
        (status,json) = self.get(url)
        if status == codes.ok and len(json) > 0:
            return json[0]['applicationResults']
        else:
            return None

    def get_grade(self,domain_id, key):
        url = f'{domain_id}/applications/3/results?quality-indicators={key}'
        (status,json) = self.get(url)
        if status == codes.ok and len(json) > 0:
            return json[0]['applicationResults']
        else:
            return None

    def get_latest_snapshot(self,domain_id):
        snapshot = {}
        (status,json) = self.get(f'{domain_id}/applications/3/snapshots')
        if status == codes.ok and len(json) > 0:
            snapshot['id'] = json[0]['href'].split('/')[-1]  
            snapshot['name'] = json[0]['name']
            snapshot['technology'] = json[0]['technologies']
            snapshot['module_href'] = json[0]['moduleSnapshots']['href']
            snapshot['result_href'] = json[0]['results']['href'] 
        return snapshot 

    def get_grades_by_technology(self,domain_id,snapshot):
        first_tech=True
        grade = DataFrame(columns=list(self._measures.values()))
        for tech in snapshot['technology']:
            t={}
            a={}
            for key in self._measures: 
                url = f'{domain_id}/applications/3/results?quality-indicators={key}&technologies={tech}'
                (status,json) = self.get(url)
                if status == codes.ok and len(json) > 0:
                    try:
                        t[self._measures[key]]=json[0]['applicationResults'][0]['technologyResults'][0]['result']['grade']
                    except IndexError:
                        self.warning(f'{domain_id} no grade available for {key} {tech} setting it to 4')
                        t[self._measures[key]]=4

                    if first_tech==True:
                        a[self._measures[key]]=json[0]['applicationResults'][0]['result']['grade']
                else:
                    self.error (f'Error retrieving technology information:  {url}')
            if first_tech==True:
                grade.loc['All'] = a
            grade.loc[tech] = t
            first_tech=False
            
        return grade

    def get_sizing_by_technology(self,domain_id,snapshot,sizing):
        first_tech=True
        size_df = DataFrame(columns=list(sizing.values()))
        for tech in snapshot['technology']:
            t={}
            a={}
            for key in sizing: 
                url = f'{domain_id}/applications/3/results?sizing-measures={key}&technologies={tech}'
                (status,json) = self.get(url)
                if status == codes.ok and len(json) > 0:
                    try:
                        t[sizing[key]]= json[0]['applicationResults'][0]['technologyResults'][0]['result']['value']
                        if first_tech==True:
                            a[sizing[key]]=json[0]['applicationResults'][0]['result']['value']
                    except IndexError:
                        self.debug(f'{domain_id} no grade available for {key} {tech}')
            if first_tech==True:
                size_df.loc['All'] = a
            size_df.loc[tech] = t
            first_tech=False
        return size_df

    def get_distribution_sizing(self, domain_id, metric_id):
        rslt = DataFrame(columns=['name','value'])
        (status,json) = self.get(f'{domain_id}/applications/3/results?metrics={metric_id}&select=categories')
        if status == codes.ok and len(json) > 0:
            cat = json[0]['applicationResults'][0]['result']['categories']
            for index, name in enumerate(cat):
                rslt.loc[name['key']]=[[name['name']],[name['value']]]

        return rslt

    def get_rules(self,domain_id,snapshot_id,business_criteria,critical=True,non_critical=True,start_row=1,max_rows=10000):
        rslt_df =  DataFrame()
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
        if status == codes.ok and len(json) > 0:
            rslt_df = DataFrame(json)
        return rslt_df


    def get_action_plan(self,domain_id,snapshot_id):
        business_criteria = ['Robustness','Efficiency','Security','Transferability','Changeability']
    
        catagory = ''
        tech_criteria = ''
        rslt_df =  DataFrame()
        ap_summary_df =  DataFrame()
        crow = 1

        url = f'{domain_id}/applications/3/snapshots/{snapshot_id}/action-plan/issues?startRow=1&nbRows=100000'
        self.debug('begin - get action plan records')
        (status,json) = self.get(url)
        self.debug('end - get action plan records')
        if status == codes.ok and len(json) > 0:
            rslt_df = DataFrame(json)
            rule_pattern = json_normalize(rslt_df['rulePattern']).add_prefix('rule.')
            rule_pattern['rule.href'] = rule_pattern['rule.href'].str.split('/').str[-1]
            rule_pattern = rule_pattern.rename(columns={"rule.href":"rule.id"})

            component = json_normalize(rslt_df['component']).add_prefix('component.') 
            component.drop(columns=component.columns.difference(['component.name','component.shortName']),axis=1,inplace=True)

            remediation = json_normalize(rslt_df['remedialAction']) 
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
                    if status == codes.ok and len(json) > 0:
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
            business = DataFrame(rslt_df,columns=['rule.name','tech_criteria','Business Criteria','tag','comment']).drop_duplicates()
            ap_summary_df.drop(columns=ap_summary_df.columns.difference(['rule.name','component.name']),axis=1,inplace=True)
            ap_summary_df = merge(ap_summary_df,business, on='rule.name')
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
        if status == codes.ok and len(json) > 0:
            loc = json[0]['applicationResults'][0]['result']['value']
        return loc

    def get_sizing(self, domain_id, input):
        rslt = {}
        for key in input: 
            (status,json) = self.get(f'{domain_id}/applications/3/results?sizing-measures={key}&snapshots=-1')
            if status == codes.ok and len(json) > 0:
                rslt[input[key]]=json[0]['applicationResults'][0]['result']['value']
        return rslt

    def get_violation_CR(self,domain_id):
        vs = self.get_sizing(domain_id,self._violations) 
        complexity = self.get_distribution_sizing(domain_id,'67001')
        vs['Complex objects']=complexity.loc['67002']['value'][0]+complexity.loc['67003']['value'][0]
        complexity = self.get_distribution_sizing(domain_id,'67030')
        vs[' With violations']=complexity.loc['67031']['value'][0]+complexity.loc['67032']['value'][0]
        return vs
