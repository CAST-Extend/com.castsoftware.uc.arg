import math
import util
import pandas as pd



class ActionPlan:
    _app_list = []
    _ppt = None
    _aip_data = None
    _effort_df = None
    _output_folder = None

    _day_rate=1200

    __extrm_effort = 0
    __high_effort = 0
    __med_effort = 0
    __low_effort = 0

    __extrm_cost = 0
    __high_cost = 0
    __low_cost = 0
    __med_cost = 0

    __extrm_vio_total = 0
    __high_vio_total = 0
    __med_vio_total = 0
    __low_vio_total = 0

    __extrm_data = ''
    __high_data = ''
    __med_data = ''
    __low_data = ''

    def get_extreme_costing(self):    
        return  self.__extrm_effort, \
                self.__extrm_cost, \
                self.__extrm_total, \
                self.__extrm_data

    def get_high_costing(self):  
        return  self.__high_effort, \
                self.__high_cost, \
                self.__high_total, \
                self.__high_data

    def get_med_costing(self):    
        return  self.__med_effort, \
                self.__med_cost, \
                self.__med_total, \
                self.__med_data

    def get_low_costing(self):    
        return  self.__low_effort, \
                self.__low_cost, \
                self.__low_total, \
                self.__low_data

    def __init__(self,app_list,aip_data,ppt,output_folder):
        self._app_list = app_list
        self._output_folder=output_folder
        self._ppt = ppt
        self._aip_data=aip_data
        self._effort_df = pd.read_csv('./Effort.csv')

    def fill_action_plan(self,app_no):
        app_id = self._app_list[app_no]

        (ap_df,ap_summary_df)=self._aip_data.action_plan(app_id)
        if not ap_summary_df.empty:
            ap_summary_df = ap_summary_df.merge(self._effort_df, how='inner', on='Technical Criteria')
            #cost_col = (ap_summary_df['Eff Hours'] * ap_summary_df['No. of Actions'])/8
            ap_summary_df['Days Effort'] = (ap_summary_df['Eff Hours'] * ap_summary_df['No. of Actions'])/8
            ap_summary_df['Cost Est.'] = ap_summary_df['Days Effort'] * self._day_rate

            file_name = f'{self._output_folder}/{app_id}_action_plan.xlsx'
            writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
            col_widths=[50,40,10,10,10,50,10,10,10]
            summary_tab = util.format_table(writer,ap_summary_df,'Summary',col_widths)
            col_widths=[10,50,50,30,30,30,30,30,30,30,30,30,30]
            util.format_table(writer,ap_df,'Action Plan',col_widths)
            writer.save()

            #fill action plan related tags
            (self.__extrm_effort,self.__extrm_cost,self.__extrm_total,self.__extrm_data) = \
                self.fill_action_plan_text(ap_summary_df,app_no,'extreme','security')
            
            (self.__high_effort,self.__high_cost,self.__high_total,self.__high_data) = \
                self.fill_action_plan_text(ap_summary_df,app_no,'high')

            (self.__med_effort,self.__med_cost,self.__med_total,self.__med_data) = \
                self.fill_action_plan_text(ap_summary_df,app_no,'moderate')

            (self.__low_effort,self.__low_cost,self.__low_total,self.__low_data) = \
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

#            ap_table = ap_table.drop(columns=['comment','tag','Technical Criteria','Days Effort','Cost Est.','Eff Hours'],index=0)
            ap_table = ap_table.drop(columns=['comment','tag','Technical Criteria','Days Effort','Cost Est.','Eff Hours'])

            self._ppt.update_table(f'app{app_no+1}_action_plan',ap_table.head(29),include_index=False,background='RGB')

            sum = ap_summary_df['No. of Actions'].sum()
            self._ppt.replace_text(f"{{app{app_no+1}_total_violations}}",str(sum))

            violation_table = self._ppt.get_shape_by_name(f'app{app_no+1}_action_plan') 
            if violation_table:
                page_no = self._ppt.get_page_no(violation_table)
                self._ppt.replace_text(f"{{app{app_no+1}_violation_page}}",str(page_no))

            violation_table = self._ppt.get_shape_by_name(f'app{app_no+1}_table_type_name') 
            if violation_table:
                page_no = self._ppt.get_page_no(violation_table)
                self._ppt.replace_text(f"{{app{app_no+1}_HL_violation_page}}",str(page_no))
        else:
            self._ppt.replace_text(f"{{app{app_no+1}_extreme_violation_total}}",'TBD') 
            self._ppt.replace_text(f"{{app{app_no+1}_high_violation_total}}",'TBD') 
            self._ppt.replace_text(f"{{app{app_no+1}_moderate_violation_total}}",'TBD') 
            self._ppt.replace_text(f"{{app{app_no+1}_low_violation_total}}",'TBD') 

    def fill_action_plan_text(self,ap_summary_df,app_no,priority,default=''):
        (priority_text, violation_total, filtered) = self.common_business_criteria(ap_summary_df,priority,default)
        violation_text = self.list_violations(filtered)
        days_effort = math.ceil(filtered['Days Effort'].sum())
        cost_effort = (days_effort*self._day_rate)/1000

        # self._ppt.replace_text(f"{{app{app_no+1}_{priority}_business_criteria_text}}",priority_text.lower()) 
        # self._ppt.replace_text(f"{{app{app_no+1}_{priority}_violation_total}}",violation_total) 
        # self._ppt.replace_text(f"{{app{app_no+1}_{priority}_violation_text}}",self.list_violations(filtered)) 
        # self._ppt.replace_text(f"{{app{app_no+1}_{priority}_cost}}",f'${cost_effort*2}K') 
        # self._ppt.replace_text(f"{{app{app_no+1}_{priority}_days}}",f'{days_effort*2}')

        return days_effort, cost_effort, violation_total, filtered

    def common_business_criteria(self,summary_df,priority,default=''):
        filtered=summary_df[summary_df['tag']==priority]
        count = 0
        sum = 0
        list = []
        if not filtered.empty:
            sum = filtered['No. of Actions'].sum()
            list = self.business_criteria(filtered)
        sum_txt = str(sum)
        
        if not list:
            list.append(default)

        return util.list_to_text(list),sum_txt, filtered

    def business_criteria(self,filtered):
        list = []
        if not filtered.empty:
            for business in filtered['Business Criteria']:
                items = business.split(',')
                for t in items:
                    if t.strip() not in list:
                        list.append(t.strip())
        return list


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
