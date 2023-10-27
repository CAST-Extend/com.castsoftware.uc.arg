from cast_common.powerpoint import PowerPoint as common_ppt
from cast_common.logger import INFO
from pptx.chart.data import CategoryChartData
from pandas import Series

__author__ = "Nevin Kaplan"
__email__ = "n.kaplan@castsoftware.com"
__copyright__ = "Copyright 2022, CAST Software"


class PowerPoint(common_ppt):
    """
    This module holds techniques explicitly designed for 
    constructing PowerPoint slides. The majority of its capabilities 
    come from the PowerPoint module in the com.castsoftware.uc.python.common library. 

    Args:
        common_ppt: PowerPoint class found in com.castsoftware.uc.python.common
    """

    def __init__(self,template:str=None,output:str=None,log_level=INFO) -> None:
        super().__init__(template,output,log_level)

    def replace_risk_factor(self, grades:Series, app_no:int=0, search_str:str=None):
        if search_str == None:
            search_str=f'{{app{app_no}_risk_'
        
        for slide in self._prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        if paragraph.text.find(search_str)!=-1:
                            run = self._merge_runs(paragraph)
                            # cur_text=''
                            # first=True
                            # for run in paragraph.runs:
                            #     cur_text = cur_text + run.text
                            #     if first != True:
                            #         self.delete_run(run)
                            #     first=False
                            # run = paragraph.runs[0]
                            # run.text = cur_text

                            while True:
                                t = cur_text.find(search_str)
                                if t == -1: 
                                    break
                                g = cur_text[t+len(search_str):]
                                g = g[:g.find("}")]
                                if g:
                                    try:
                                        grade=grades[g]
                                        risk = ''
                                        if grade < 2:
                                            risk = 'high'
                                        elif grade < 3:
                                            risk = 'medium'
                                        else:
                                            risk = 'low'

                                        cur_text = cur_text.replace(f'{search_str}{g}}}',risk)
                                        run.text = cur_text
                                    except KeyError:
                                        self.debug(f'invalid key: {g}')
                                        break;

    def update_grade_slider(self, shape,data):
        #shape = self.get_shape_by_name(name)
        try:
            if shape.has_chart:
                chart_data = CategoryChartData()
                # chart_data.categories = titles
                
                chart_data.categories = ['grade']
                chart_data.add_series('Series 1', data)
                shape.chart.replace_data(chart_data)
        except AttributeError as ae:
            self.warning(f'Attribute Error, invalid template configuration: {shape.name} ({ae})')
        except KeyError as ke:
            self.warning(f'Key Error, invalid template configuration: {shape.name} ({ke})')
        except TypeError as t:
            self.warning(f'Type Error, invalid template configuration: {shape.name} ({t})')

    def remove_empty_placeholders(self):
        for slide in self._prs.slides:
            for placeholder in slide.shapes.placeholders:
                if placeholder.has_text_frame and placeholder.text_frame.text == "":
                    sp = placeholder._sp
                    sp.getparent().remove(sp)

    def replace_loc(self, loc, app_no):
        loc_short = "{0:,.0f} LoC".format(loc) 
        if loc > 1000000:
            loc_short = "~{0:,.2f} MLoC".format(loc/1000000) 
        elif loc < 1000000 and loc > 1000:
            loc_short = "~{0:,.0f} KLoC".format(loc/1000) 
        self.replace_text(f'{{app{app_no}_loc}}',f'{loc:,.0f}')
        self.replace_text(f'{{app{app_no}_loc_short}}',loc_short)

        size_catagory = 'small'
        if loc > 1000000:
            size_catagory = 'very large'
        elif loc > 500000:
            size_catagory = 'large'
        self.replace_text(f'{{app{app_no}_loc_category}}',size_catagory)

