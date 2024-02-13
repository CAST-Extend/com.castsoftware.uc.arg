
from cast_common.logger import Logger, INFO,DEBUG
#from cast_common.aipRestCall import AipRestCall
#from cast_arg import AipData
from cast_common.mri import MRI
from cast_arg.powerpoint import PowerPoint
from cast_arg.pages.mri_report import MRIPage
from pandas import json_normalize,DataFrame
from tqdm import tqdm



class TechDetailTable(MRIPage):
    description = 'Technical Detail Table'

    """
    This class is used to fill in the AIP Technical detail table
    """
    def report(self,app_name:str,app_no:int) -> bool:
        # app_tag = f'app{app_no}'

        sizing = {
        '10151':'Number of Code Lines', 
        '67011':'Critical Violations'
        }

        domain_id = self.get_domain(app_name)
        snapshot = self.get_latest_snapshot(domain_id)
        (ap_df,ap_summary_df)=self.get_action_plan(domain_id,snapshot['id'])        

        sizing_df = self.get_sizing_by_module(domain_id,snapshot,sizing)
        sizing_df['Fix Now']=0
        sizing_df['Near Term']=0
        sizing_df['Mid Term']=0
        # sizing_df['Long Term']=0
        if ap_df.empty:
            for key, value in sizing_df.iterrows():
                sizing_df.at[key,'Fix Now'] = 0
                sizing_df.at[key,'Near Term'] = 0
                sizing_df.at[key,'Mid Term'] = 0
        else:
            for key, value in sizing_df.iterrows():
                if key=='All':
                    sizing_df.at[key,'Fix Now'] = len(ap_df[ap_df['Action Plan Priority']=='Fix Now'])
                    sizing_df.at[key,'Near Term'] = len(ap_df[ap_df['Action Plan Priority']=='Near Term'])
                    sizing_df.at[key,'Mid Term'] = len(ap_df[ap_df['Action Plan Priority']=='Mid Term'])
                    # sizing_df.at[key,'Long Term'] = len(ap_df[ap_df['Action Plan Priority']=='Long Term'])
                else:
                    sizing_df.at[key,'Fix Now'] = self._get_counts(ap_df,'Fix Now',key)
                    sizing_df.at[key,'Near Term'] = self._get_counts(ap_df,'Near Term',key)
                    sizing_df.at[key,'Mid Term'] = self._get_counts(ap_df,'Mid Term',key)
                    # sizing_df.at[key,'Long Term'] = self._get_counts(ap_df,'Long Term',key)
                pass

        sizing_df = sizing_df.astype('int')

        # sizing_df=sizing_df[['Number of Code Lines','Fix Now','Near Term','Mid Term','Long Term','Critical Violations']]
        sizing_df['Number of Code Lines'] = sizing_df['Number of Code Lines'].map('{:,.0f}'.format)
        PowerPoint.ppt.update_table(f'app{app_no}_technical_details_table',sizing_df,header_rows=2)

    def _get_counts(self,data:DataFrame,priority:str,tech:str) -> int:
        t = tech.replace('+',r'\+').lower()
        fn = data[data['Action Plan Priority']==priority]
        fn = fn[fn['Technology'].str.lower().str.contains(t)]
        return len(fn)
        


        pass

    # def replace_text(self,prs, prefix, item, data):
    #     tag = f'{{{prefix}_{item}}}'
    #     self.log.debug(f'{tag}: {data}')
    #     prs.replace_text(tag,data)

