from requests import codes
from pandas import DataFrame,json_normalize

from cast_common.logger import Logger, INFO,DEBUG
from cast_common.mri import MRI

class ActionPlan(MRI):

    _action_plan = {}
    def load_data(self,domain_id:str,snapshot_id:int):
        key=f'{domain_id}:{snapshot_id}'
        if not key in ActionPlan._action_plan.keys():    
            url = f'{domain_id}/applications/3/snapshots/{snapshot_id}/action-plan/issues?startRow=1&nbRows=100000'
            (status,json) = self.get(url)
            if status == codes.ok and len(json) > 0:
                rslt_df = DataFrame(json)

                rule_pattern = json_normalize(rslt_df['rulePattern']).add_prefix('rule.')
                rule_pattern['rule.href'] = rule_pattern['rule.href'].str.split('/').str[-1]
                rule_pattern = rule_pattern.rename(columns={"rule.href":"rule.id"})

                component = json_normalize(rslt_df['component']).add_prefix('component.') 
                remediation = json_normalize(rslt_df['remedialAction']) 
                rslt_df = rule_pattern.join([component,remediation])                                                  

                rslt_df.insert(4,'tech_criteria','')
                rslt_df.insert(4,'Business Criteria','')
                rslt_df.insert(4,'component.tech','')
            else:
                raise ValueError('No Data Returned')
            pass


        pass

    def export_excel(self):
        pass

    def export_json(self):
        pass


ap = ActionPlan(base_url='http://arch-ps-2:8087/rest',
                token=True,user='admin',pswd='85a39ae85a156b2ab184ce0a709e964e')
ap.load_data('458018d7-7186-46aa-8234-cdc8019f1b7d',0)
pass


    