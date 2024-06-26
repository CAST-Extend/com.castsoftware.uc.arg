from cast_common.highlight import Highlight
from cast_common.powerpoint import PowerPoint
from cast_common.logger import DEBUG,INFO
from pptx.chart.data import CategoryChartData
from pandas import DataFrame
from bisect import bisect_left

class HighlightBenchmark(Highlight):

    quartile_txt = ['4th','3rd','2nd','1st']

    _benchmark = {}

    def __init__(self,log_level=DEBUG):
        self._log_level=log_level
        super().__init__()

        #get the benchmark data
        self._benchmark = DataFrame(self._get(r'/benchmark'))
        pass


    def report(self,app_list:str|list=None,app_no:int=0) -> bool:
        overall_health_txt = {'high':'excellent','medium':'adequate','low':'poor'}

        if type(app_list) is list and len(app_list)>1:
            self._tag_prefix = 'port'
        else:
            self._tag_prefix = f'app{app_no}'
            app_list = [app_list]
        self._tag_prefix = f'{self._tag_prefix}_hl'

        #get data
        t_scores = self.calc_scores(app_list)

        bm = self._benchmark
        sample_size = int(bm['sampleSize'].iloc[0])
        PowerPoint.ppt.replace_text(f'{{bm_sample_size}}',sample_size)

        quart_cols = [f for f in bm.columns if f.startswith('quartile')]
        quartile = bm[quart_cols].dropna().reset_index()

        #loop through the Highlight grades
        for key in self.grades:
            score = t_scores[key] 
            self.replace_text(f'{key}_score',score)

            ind_avg = round(float(bm.loc[key]['avg'])*100,2)
            PowerPoint.ppt.replace_text(f'{{bm_industry_{key}_score}}',ind_avg)

            #calculat the High/Medium/Low (HML) values for each grade
            threshold = self.grades[key]['threshold']
            if len(threshold)>0:
                if score < threshold[0]:
                    hml = 'low'
                elif score > threshold[1]:
                    hml = 'high'
                else:
                    hml = 'medium'
                
                #now set the color according to the HML caluculation 
                color = self.get_hml_color(hml)
                PowerPoint.ppt.fill_text_box_color(f'{self._tag_prefix}_{key}_tile',color)

                #calculate the quartile boundries
                qtr = quartile[quartile['index']==key][quart_cols]
                if not qtr.empty:
                    qtr = [round(x*100,2) for x in list(qtr.iloc[0])]
                    qtr.append(100)

                #add the quartile bountries table as chart data to the stacked barchart
                self.fill_slider(key,score,qtr.copy())
                #self.fill_slider(f'{key}_tech',0,qtr.copy())

                #calculate witch quartile the score falls in and fill in the label 
                idx = bisect_left(qtr,score)
                qtr_txt = self.quartile_txt[idx]           
                self.replace_text(f'{key}_quartile',qtr_txt)

        if self._tag_prefix != 'port_hl':
            tech = self.get_technology(app_list[0])   
            top_tech = tech.iloc[0]['technology']
            PowerPoint.ppt.replace_text(f'{{{self._tag_prefix}_top_tech}}',top_tech)

        pass

    def replace_text(self,item, data):
        tag = f'{{{self._tag_prefix}_{item}}}'
        self.log.debug(f'{tag}: {data}')
        PowerPoint.ppt.replace_text(tag,data)

    def fill_slider(self,name,score:float,quartiles:list):
        #self._ppt.update_chart(f'app{app_no}_sizing_pie_chart',DataFrame(grade_by_tech_df['LOC']))

        #update the score barchart
        shape = PowerPoint.ppt.get_shape_by_name(f'{self._tag_prefix}_bm_{name}_score')
        if shape is not None and shape.has_chart:
            chart = shape.chart
            chart_data = CategoryChartData()
            chart_data.categories=['Score']
            chart_data.add_series('series',[score])
            chart.replace_data(chart_data)

        #update the stacked barchart with the quartile values
        shape = PowerPoint.ppt.get_shape_by_name(f'{self._tag_prefix}_bm_{name}_slider')
        if shape is not None and shape.has_chart:
            try:
                chart = shape.chart

                first = True
                prev = 0
                for idx,q in enumerate(quartiles):
                    quartiles[idx] = q - prev
                    prev = quartiles[idx]+prev

                chart_data = CategoryChartData()
                chart_data.categories=['Score']
                first = True
                idx = len(quartiles)

                for q in quartiles:
                    chart_data.add_series(f'q{idx}',[q])
                chart.replace_data(chart_data)


                # chart_data.categories=['Quartile','Score']
                # chart_data.add_series(f'q{idx}',tuple(quartiles))

                # s=0
                # for q in quartiles:
                #     if idx == 1:
                #         s = score
                #     chart_data.add_series(f'q{idx}',(q,s))
                #     idx -= 1

            except AttributeError as ae:
                self.warning(f'Attribute Error, invalid template configuration: {name} ({ae})')
            except KeyError as ke:
                self.warning(f'Key Error, invalid template configuration: {name} ({ke})')
            except TypeError as t:
                self.warning(f'Type Error, invalid template configuration: {name} ({t})')
        pass


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
