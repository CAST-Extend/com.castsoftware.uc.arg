from stats import AIPStats
from cast_common.logger import Logger,INFO
from cast_common.aipRestCall import AipRestCall
from cast_common.util import format_table,list_to_text
from os import getcwd
from os.path import abspath,dirname,exists
from site import getsitepackages


import math
import pandas as pd

__author__ = "Nevin Kaplan"
__email__ = "n.kaplan@castsoftware.com"
__copyright__ = "Copyright 2022, CAST Software"


class ActionPlan(AipRestCall):
    """
        This class is used to collect action plan information and add them to the 
        the proper tags
        
        ERROR HANDLING IMPROVEMENTS (v1.8.0):
        - Fixed IndexError when accessing DataFrame rows that don't exist
        - Added comprehensive error handling for all DataFrame operations
        - Added graceful degradation when data is missing or incomplete
        - Added fallback values ("N/A", "0") for missing data
        - Added column existence checks before operations
        - Added NaN value handling for business criteria
        - Added file operation error handling
        - Added type safety for data conversions
    """

    _app_list = []
    _ppt = None
    _aip_data = None
    _effort_df = None
    _output_folder = None

    def __init__(self,app_list,aip_data,ppt,output_folder,day_rate,logger_level=INFO):
        self._app_list = app_list
        self._output_folder=output_folder
        self._ppt = ppt
        self._aip_data=aip_data

        #ef_name = abspath(f'{dirname(__file__)}/Effort.csv')
        # self.ef_name = abspath(f'{getsitepackages()[-1]}/cast_arg/Effort.csv')
        self.ef_name = abspath(f'src/cast_arg/Effort.csv')

        if not exists(self.ef_name):
            raise RuntimeError(f'Required file not found: {self.ef_name}')
        self._effort_df = pd.read_csv(self.ef_name)

        self._day_rate = day_rate
        self._fix_now = AIPStats(day_rate)
        self._high = AIPStats(day_rate)
        self._med = AIPStats(day_rate)
        self._low = AIPStats(day_rate)


    @property
    def fix_now(self): return self._fix_now
    @property
    def high(self): return self._high
    @property
    def medium(self): return self._med
    @property
    def low(self): return self._low
    @property
    def day_rate(self): return self._day_rate


    def fill_action_plan(self,app_id,app_no):
        """
        Fill action plan data with comprehensive error handling.
        
        ORIGINAL ISSUE: IndexError when accessing mid_term_action_top2.iloc[1] 
        when there were fewer than 2 "moderate" actions available.
        
        FIXES APPLIED:
        1. Added bounds checking before DataFrame access
        2. Added fallback values ("N/A") for missing data
        3. Added comprehensive try-catch blocks for all operations
        4. Added column existence checks
        5. Added graceful degradation for missing data
        """

        (ap_df,ap_summary_df)=self._aip_data.action_plan(app_id)
        if not ap_summary_df.empty:
#            ap_summary_df = ap_summary_df.merge(self._effort_df, how='inner', on='Technical Criteria')
            ap_summary_df = ap_summary_df.merge(self._effort_df, how='left', on='Technical Criteria')
            missing_effort = ap_summary_df[ap_summary_df['Eff Hours'].isnull()]['Technical Criteria']
            if len(missing_effort) > 0:
                self._log.warning('*******************************************')
                self._log.warning('Action plan effort will not be accurate!')
                self._log.warning(f'Missing entries for: {missing_effort.tolist()}')
                self._log.warning(f'Update Effort Table located at: {self.ef_name}')
                self._log.warning('*******************************************')
                ap_summary_df.loc[ap_summary_df['Technical Criteria'].isin(missing_effort),'Eff Hours']=0


            #cost_col = (ap_summary_df['Eff Hours'] * ap_summary_df['No. of Actions'])/8
            ap_summary_df['Days Effort'] = (ap_summary_df['Eff Hours'] * ap_summary_df['No. of Actions'])/8
            ap_summary_df['Cost Est.'] = ap_summary_df['Days Effort'] * self._day_rate

            ap_summary_df = ap_summary_df.sort_values(by='No. of Actions', ascending=False)

            # print(ap_summary_df)

            # ERROR HANDLING: Excel file creation with graceful degradation
            # If Excel file creation fails, program continues execution
            try:
                file_name = f'{self._output_folder}/{app_id}_action_plan.xlsx'
                writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
                col_widths=[50,40,10,10,10,50,10,10,10]
                # ERROR HANDLING: Check if 'comment' column exists before renaming
                if 'comment' in ap_summary_df.columns:
                    ap_summary_df.rename(columns={'comment': 'Priority'}, inplace=True)
                summary_tab = format_table(writer,ap_summary_df[['Quality Rule','Business Criteria','No. of Actions','Priority']],'Summary',col_widths)
                col_widths=[10,50,50,30,30,30,30,30,30,30,30,30,30]
                # ERROR HANDLING: Check if 'Rule Id' column exists before dropping
                if 'Rule Id' in ap_df.columns:
                    ap_df = ap_df.drop('Rule Id', axis=1)
                format_table(writer,ap_df,'Action Plan',col_widths)
                writer.close()
            except Exception as e:
                self._log.warning(f"Error creating Excel file: {e}")
                # Continue execution even if Excel file creation fails

            #fill action plan related tags
            self._fix_now = self.calc_action_plan_effort(ap_summary_df,app_no,'extreme','security')
            self._high = self.calc_action_plan_effort(ap_summary_df,app_no,'high')
            self._med = self.calc_action_plan_effort(ap_summary_df,app_no,'moderate')
            self._low = self.calc_action_plan_effort(ap_summary_df,app_no,'low')
            

            # ERROR HANDLING: RGB color setting with column existence check
            # If 'tag' column doesn't exist, skip RGB color assignment
            try:
                ap_summary_df.loc[ap_summary_df['tag']=='extreme','RGB']='244,212,212'
                ap_summary_df.loc[ap_summary_df['tag']=='high','RGB']='255,229,194'
                ap_summary_df.loc[ap_summary_df['tag']=='moderate','RGB']='203,225,238'
                ap_summary_df.loc[ap_summary_df['tag']=='low','RGB']='217,217,217'
            except KeyError as e:
                self._log.warning(f"Error setting RGB colors: {e}")

            # ERROR HANDLING: Safe DataFrame concatenation and column dropping
            # Creates fallback empty DataFrame if concatenation fails
            try:
                ap_table = pd.concat([ap_summary_df[ap_summary_df['tag']=='extreme'],
                                      ap_summary_df[ap_summary_df['tag']=='high'],
                                      ap_summary_df[ap_summary_df['tag']=='moderate'],
                                      ap_summary_df[ap_summary_df['tag']=='low']])

                # ERROR HANDLING: Only drop columns that actually exist
                columns_to_drop = ['Priority','Technical Criteria','Days Effort','Cost Est.','Eff Hours']
                existing_columns = [col for col in columns_to_drop if col in ap_table.columns]
                if existing_columns:
                    ap_table = ap_table.drop(columns=existing_columns)
            except Exception as e:
                self._log.warning(f"Error creating ap_table: {e}")
                ap_table = pd.DataFrame()  # Create empty DataFrame as fallback

            # ERROR HANDLING: Safe table row processing
            # Continues with existing ap_table if processing fails
            try:
                starting_rows = ap_table[ap_table['tag'] == 'extreme']

                # 6 rows for each of C, D, B
                other_tags = ['high', 'moderate', 'low']
                other_rows = ap_table[ap_table['tag'].isin(other_tags)].groupby('tag').head(5)

                # Combine and reorder
                ap_table = pd.concat([starting_rows] + [other_rows[other_rows['tag'] == tag] for tag in other_tags])
            except Exception as e:
                self._log.warning(f"Error processing table rows: {e}")
                # Continue with existing ap_table if processing fails

            # print(ap_table)

            # CRITICAL FIX: Safe top indices processing with comprehensive error handling
            # This fixes the original IndexError by checking DataFrame bounds before access
            try:
                top_indices = ap_table[ap_table["tag"] != "moderate"].groupby("tag")["No. of Actions"].idxmax()
                
                # ERROR HANDLING: Check if "extreme" key exists and value is not NaN
                if "extreme" in top_indices and not pd.isna(top_indices["extreme"]):
                    top1_immediate_action = ap_table.loc[top_indices["extreme"]]
                    self._ppt.replace_text(f"{{top_1_immediate_action}}", str(top1_immediate_action['Quality Rule']))
                else:
                    self._ppt.replace_text(f"{{top_1_immediate_action}}", "N/A")
                    
                # ERROR HANDLING: Check if "high" key exists and value is not NaN
                if "high" in top_indices and not pd.isna(top_indices["high"]):
                    top1_near_term_action = ap_table.loc[top_indices["high"]]
                    self._ppt.replace_text(f"{{top_1_near_term_action}}", str(top1_near_term_action['Quality Rule']))
                else:
                    self._ppt.replace_text(f"{{top_1_near_term_action}}", "N/A")
            except (KeyError, IndexError, ValueError) as e:
                self._log.warning(f"Error processing top indices: {e}")
                self._ppt.replace_text(f"{{top_1_immediate_action}}", "N/A")
                self._ppt.replace_text(f"{{top_1_near_term_action}}", "N/A")

            # CRITICAL FIX: Safe mid-term action processing with bounds checking
            # This is the main fix for the original IndexError issue
            try:
                mid_term_df = ap_table[ap_table["tag"] == "moderate"]
                if not mid_term_df.empty:
                    mid_term_action_top2 = mid_term_df.nlargest(2, "No. of Actions")
                    
                    # ERROR HANDLING: Check DataFrame length before accessing iloc[0]
                    if len(mid_term_action_top2) > 0:
                        top1_mid_term_action = mid_term_action_top2.iloc[0]
                        self._ppt.replace_text(f"{{top_1_mid_term_action}}", str(top1_mid_term_action['Quality Rule']))
                    else:
                        self._ppt.replace_text(f"{{top_1_mid_term_action}}", "N/A")

                    # ERROR HANDLING: Check DataFrame length before accessing iloc[1]
                    # This prevents the original IndexError when < 2 rows exist
                    if len(mid_term_action_top2) > 1:
                        top2_mid_term_action = mid_term_action_top2.iloc[1]
                        self._ppt.replace_text(f"{{top_2_mid_term_action}}", str(top2_mid_term_action['Quality Rule']))
                    else:
                        self._ppt.replace_text(f"{{top_2_mid_term_action}}", "N/A")
                else:
                    # ERROR HANDLING: No moderate actions found, use fallback values
                    self._ppt.replace_text(f"{{top_1_mid_term_action}}", "N/A")
                    self._ppt.replace_text(f"{{top_2_mid_term_action}}", "N/A")
            except (KeyError, IndexError, ValueError) as e:
                self._log.warning(f"Error processing mid term actions: {e}")
                self._ppt.replace_text(f"{{top_1_mid_term_action}}", "N/A")
                self._ppt.replace_text(f"{{top_2_mid_term_action}}", "N/A")

            # ERROR HANDLING: Safe column dropping with existence check
            if 'tag' in ap_table.columns:
                ap_table = ap_table.drop(columns=['tag'])

            try:
                self._ppt.update_table(f'app{app_no}_action_plan',ap_table,app_id,include_index=False,background='RGB')
            except ValueError as ex:
                self._log.warning('Action plan table missing from tamplate')
                
            # ERROR HANDLING: Safe sum calculation with type checking
            try:
                sum = ap_summary_df['No. of Actions'].sum()
                self._ppt.replace_text(f"{{app{app_no}_total_violations}}",str(sum))
            except (KeyError, TypeError) as e:
                self._log.warning(f"Error calculating total violations: {e}")
                self._ppt.replace_text(f"{{app{app_no}_total_violations}}","0")

            # ERROR HANDLING: Safe page number retrieval for action plan table
            try:
                violation_table = self._ppt.get_shape_by_name(f'app{app_no}_action_plan') 
                if violation_table:
                    page_no = self._ppt.get_page_no(violation_table)
                    self._ppt.replace_text(f"{{app{app_no}_violation_page}}",str(page_no))
            except Exception as e:
                self._log.warning(f"Error getting violation page number: {e}")

            # ERROR HANDLING: Safe page number retrieval for HL table
            try:
                violation_table = self._ppt.get_shape_by_name(f'app{app_no}_table_type_name') 
                if violation_table:
                    page_no = self._ppt.get_page_no(violation_table)
                    self._ppt.replace_text(f"{{app{app_no}_HL_violation_page}}",str(page_no))
            except Exception as e:
                self._log.warning(f"Error getting HL violation page number: {e}")
        else:
            # ERROR HANDLING: Empty DataFrame case - use fallback values
            self._ppt.replace_text(f"{{app{app_no}_extreme_violation_total}}",'TBD') 
            self._ppt.replace_text(f"{{app{app_no}_high_violation_total}}",'TBD') 
            self._ppt.replace_text(f"{{app{app_no}_moderate_violation_total}}",'TBD') 
            self._ppt.replace_text(f"{{app{app_no}_low_violation_total}}",'TBD') 

    def calc_action_plan_effort(self,ap_summary_df,app_no,priority,default='') -> AIPStats: 
        """
        Calculate action plan effort with comprehensive error handling.
        
        IMPROVEMENTS:
        - Added try-catch block for all operations
        - Added column existence check for 'Days Effort'
        - Added safe type conversion for violations count
        - Added fallback values when calculations fail
        """
        try:
            rslt = AIPStats(self._day_rate)
            (priority_text, vio_cnt, rslt.data) = self.common_business_criteria(ap_summary_df,priority,default)
            # ERROR HANDLING: Check if 'Days Effort' column exists before calculation
            if not rslt.data.empty and 'Days Effort' in rslt.data.columns:
                rslt.effort = math.ceil(rslt.data['Days Effort'].sum()*2)
            else:
                rslt.effort = 0
            # ERROR HANDLING: Safe type conversion for violations count
            rslt.violations = int(vio_cnt) if vio_cnt.isdigit() else 0
            return rslt
        except Exception as e:
            self._log.warning(f"Error in calc_action_plan_effort: {e}")
            rslt = AIPStats(self._day_rate)
            rslt.effort = 0
            rslt.violations = 0
            return rslt

    # def fill_action_plan_tags(self,app_no,type,effort,cost,vio_cnt,bus_txt,vio_txt):
    #     self._ppt.replace_text(f'{{app{app_no}_{type}_eff}}',effort)
    #     self._ppt.replace_text(f'{{app{app_no}_{type}_cost}}',cost)
    #     self._ppt.replace_text(f'{{app{app_no}_{type}_vio_cnt}}',vio_cnt)

    #     self._ppt.replace_text(f'{{app{app_no}_{type}_bus_txt}}',bus_txt,tbd_for_blanks=False)
    #     self._ppt.replace_text(f'{{app{app_no}_{type}_vio_txt}}',vio_txt)

    def common_business_criteria(self,summary_df,priority,default=''):
        """
        Get common business criteria with comprehensive error handling.
        
        IMPROVEMENTS:
        - Added try-catch block for all operations
        - Added column existence check for 'No. of Actions'
        - Added fallback return values when operations fail
        """
        try:
            filtered=summary_df[summary_df['tag']==priority]
            count = 0
            sum = 0
            list = []
            if not filtered.empty:
                # ERROR HANDLING: Check if 'No. of Actions' column exists
                if 'No. of Actions' in filtered.columns:
                    sum = filtered['No. of Actions'].sum()
                list = self.business_criteria(filtered)
            sum_txt = str(sum)
            
            if not list:
                list.append(default)

            return list_to_text(list),sum_txt, filtered
        except Exception as e:
            self._log.warning(f"Error in common_business_criteria: {e}")
            return default, "0", pd.DataFrame()

    def business_criteria(self,filtered):
        """
        Extract business criteria with comprehensive error handling.
        
        IMPROVEMENTS:
        - Added try-catch block for all operations
        - Added column existence check for 'Business Criteria'
        - Added NaN value checking before string operations
        - Added safe string splitting and processing
        """
        list = []
        try:
            if not filtered.empty and 'Business Criteria' in filtered.columns:
                for business in filtered['Business Criteria']:
                    # ERROR HANDLING: Check for NaN values before string operations
                    if pd.notna(business):  # Check for NaN values
                        items = business.split(',')
                        for t in items:
                            if t.strip() not in list:
                                list.append(t.strip())
        except (KeyError, AttributeError, TypeError) as e:
            self._log.warning(f"Error processing business criteria: {e}")
        return list


    # def list_violations(self,filtered):
    #     first = True
    #     text = ""
    #     try:
    #         for criteria in filtered['Technical Criteria'].unique():
    #             df = filtered[filtered['Technical Criteria']==criteria]
    #             total = df['No. of Actions'].sum()
                
    #             cases = 'for'
    #             if first:
    #                 cases = 'cases of'
    #                 first = False
                
    #             rule = criteria[criteria.find('-')+1:].strip().lower()
    #             if len(rule) == 0:
    #                 rule = criteria
    #             text = f'{text}{total} {cases} {rule}, '
    #         return util.rreplace(text[:-2],', ',' and ')
    #     except (KeyError):
    #         return ""
            
