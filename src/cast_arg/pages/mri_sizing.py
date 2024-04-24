from cast_arg.pages.mri_report import MRIPage
from cast_arg.powerpoint import PowerPoint
from pandas import DataFrame,Series

class MRISizing(MRIPage):
    description = 'Calculating MRI Sizing'

    def report(self,app_name:str,app_no:int) -> bool:
        loc_df = self.get_loc_sizing(app_name)
        if len(loc_df) > 0:
            loc = loc_df['Number of Code Lines']
            self._ppt.replace_loc(loc,app_no)

            loc_tbl = DataFrame.from_dict(data=self.get_loc_sizing(app_name),orient='index').drop('Critical Violations')
            loc_tbl = loc_tbl.rename(columns={0:'loc'})
            loc_tbl['percent'] = round((loc_tbl['loc'] / loc_tbl['loc'].sum()) * 100,2)
            loc_tbl['loc']=Series(["{0:,.0f}".format(val) for val in loc_tbl['loc']], index = loc_tbl.index)

            percent_comment = loc_tbl.loc['Number of Comment Lines','percent']
            percent_comment_out = loc_tbl.loc['Number of Commented-out Code Lines','percent']

            if percent_comment < 15:
                comment_level='low'
            elif percent_comment > 15 and percent_comment <= 20:
                comment_level='good'
            else:
                comment_level='high'
        
            self.ppt.replace_text(f'{{app{app_no}_comment_hl}}',comment_level)
            self.ppt.replace_text(f'{{app{app_no}_comment_level}}',comment_level)
            self.ppt.replace_text(f'{{app{app_no}_comment_pct}}',percent_comment)
            self.ppt.replace_text(f'{{app{app_no}_comment_out_pct}}',percent_comment_out)

            loc_tbl['percent']=Series(["{0:.2f}%".format(val) for val in loc_tbl['percent']], index = loc_tbl.index)
            self.ppt.update_table(f'app{app_no}_loc_table',loc_tbl,app_name,header_rows=0)
            self.ppt.update_chart(f'app{app_no}_loc_pie_chart',DataFrame(loc_tbl['loc']))

            # self._ppt.replace_grade(grade_all,app_no)

            self.sizing_table(app_name,app_no)
            self.violations_table(app_name,app_no)

        # self.sizing_pie_chart(app_name,app_no)

    def sizing_table(self,app_name:str,app_no:int):
        self._log.info('Filling technical sizing table')
        sizing_df = DataFrame(self.tech_sizing(app_name),index=[0])
        sizing_df['LoC']=Series(["{0:,.0f} K".format(val / 1000) for val in sizing_df['LoC']])
        sizing_df['Files']=Series(["{0:,.0f}".format(val) for val in sizing_df['Files']])
        sizing_df['Classes']=Series(["{0:,.0f}".format(val) for val in sizing_df['Classes']])
        sizing_df['SQL Artifacts']=Series(["{0:,.0f}".format(val) for val in sizing_df['SQL Artifacts']])
        sizing_df['Tables']=Series(["{0:,.0f}".format(val) for val in sizing_df['Tables']])
        sizing_df = sizing_df.transpose()
        self._ppt.update_table(f'app{app_no}_tech_sizing',sizing_df,app_name)

    def violations_table(self,app_id,app_no):
        self._log.info('Filling critical violation table')
        violation_df = DataFrame(self.violation_sizing(app_id),index=[0])
        violation_df['Violation Count']=Series(["{0:,.0f}".format(val) for val in violation_df['Violation Count']])
        violation_df[' per file']=Series(["{0:,.2f}".format(val) for val in violation_df[' per file']])
        violation_df[' per kLoC']=Series(["{0:,.2f}".format(val) for val in violation_df[' per kLoC']])
        violation_df['Complex objects']=Series(["{0:,.0f}".format(val) for val in violation_df['Complex objects']])
        violation_df[' With violations']=Series(["{0:,.0f}".format(val) for val in violation_df[' With violations']])
        self._ppt.update_table(f'app{app_no}_violation_sizing',violation_df.transpose(),app_id)
        self._ppt.replace_text(f'{{app{app_no}_critical_violations}}',violation_df['Violation Count'].loc[0])

    # def sizing_pie_chart(self,app_id,app_no):
    #     self._log.info('Filling sizing pie chart')
    #     grade_by_tech_df = self.get_grade_by_tech(app_id)
    #     if not grade_by_tech_df.empty:
    #         #add appmarq technology
    #         self._ppt.replace_text(f'{{app{app_no}_largest_tech}}',grade_by_tech_df.index[0])

    #         self._log.info('Filling Technical Overview')
    #         # Technical Overview - Lines of code by technology GRAPH

    #         loc_df = DataFrame(grade_by_tech_df['LOC'])
    #         loc_df=loc_df.reset_index()
    #         loc_df=loc_df[loc_df['LOC'].str.replace('[^0-9]', '', regex=True).astype('int64')>0]
    #         self._ppt.update_chart(f'app{app_no}_sizing_pie_chart',loc_df)

