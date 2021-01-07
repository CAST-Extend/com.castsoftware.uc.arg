from pptx import Presentation
from pptx.chart.data import ChartData

class PowerPoint:
    _input = None
    _output = None
    _prs = None

    def __init__(self,input, output):
       self._input=input 
       self._output=output

       self._prs = Presentation(self._input)

    def replace_risk_factor(self, grades):
        search_str='{risk_'
        for slide in self._prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        if paragraph.text.find(search_str)!=-1:
                            cur_text=''
                            first=True
                            for run in paragraph.runs:
                                cur_text = cur_text + run.text
                                if first != True:
                                    self.delete_run(run)
                                first=False
                            run = paragraph.runs[0]
                            run.text = cur_text

                            while(1==1):
                                t = cur_text.find(search_str)
                                if t == -1: 
                                    break
                                g = cur_text[t+len(search_str):]
                                g = g[:g.find("}")]
                                if g:
                                    grade=grades[g]
                                    risk = ''
                                    if grade < 2:
                                        risk = 'very high'
                                    elif grade < 2.5:
                                        risk = 'high'
                                    elif grade < 3:
                                        risk = 'medium'
                                    else:
                                        risk = 'low'

                                    cur_text = cur_text.replace(f'{search_str}{g}}}',risk)
                                    run.text = cur_text


    def replace_grade(self, grades):
        search_str='{grade_'

        for slide in self._prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        if paragraph.text.find(search_str)!=-1:
                            cur_text=''
                            first=True
                            for run in paragraph.runs:
                                cur_text = cur_text + run.text
                                if first != True:
                                    self.delete_run(run)
                                first=False
                            run = paragraph.runs[0]
                            run.text = cur_text

                            while(True):
                                t = cur_text.find(search_str)
                                if t == -1: 
                                    break
                                g = cur_text[t+len(search_str):]
                                g = g[:g.find("}")]
                                if g:
                                    grade=grades[g]
                                    cur_text = cur_text.replace(f'{search_str}{g}}}',str(round(grade,2)))
                                    run.text = cur_text

    def replace_text (self, search_str, repl_str):
        for slide in self._prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        if paragraph.text.find(search_str)!=-1:
                            cur_text=''
                            first=True
                            for run in paragraph.runs:
                                cur_text = cur_text + run.text
                                if first != True:
                                    self.delete_run(run)
                                first=False
                            run = paragraph.runs[0]
                            run.text = cur_text.replace(str(search_str), str(repl_str))
    
    def update_chart(self, chart_name,categories,data):
        for slide in self._prs.slides:
            for shape in slide.shapes:
                if shape.has_chart and shape.name == chart_name:
                    chart_data = ChartData()
                    chart_data.categories = categories
                    chart_data.add_series('Series 1', data)
                    shape.chart.replace_data(chart_data)

    def get_shape_by_name(self, name):
        rslt = None
        for slide in self._prs.slides:
            for shape in slide.shapes:
                if shape.name == name:
                    rslt = shape
        return rslt

    def merge_runs(self, paragraph):
        cur_text=''
        first=True
        for run in paragraph.runs:
            cur_text = cur_text + run.text
            if first != True:
                self.delete_run(run)
            first=False
        run = paragraph.runs[0]
        run.text = cur_text
        return run

    def fill_table(self, name, df, orient):
        table_shape = self.get_shape_by_name(name)
        if table_shape.has_table:
            table=table_shape.table

            if orient=='index':
                for r in range (0,len(table.rows)):
                    text = data_top = df.head().index[r]
                    run = self.merge_runs(table.cell(r,0).text_frame.paragraphs[0])
                    run.text = run.text.replace(run.text,text)

                for r in range (0,len(table.rows)-1):
                    row = table.rows[r]
                    for c in range (0,len(row.cells)-1):
                        cell = table.cell(r,c+1)
                        run = self.merge_runs(cell.text_frame.paragraphs[0]) 
                        text = df[c][r]
                        run.text = run.text.replace(run.text,text)

    def delete_run(self,run):
        r = run._r
        r.getparent().remove(r)

    def save(self):
        self._prs.save(self._output)

"""
from restCall import AipRestCall
import pandas as pd

ppt = PowerPoint("..\\data\\template.pptx","..\\data\\test.pptx")
aip = AipRestCall("http://sha-dd-console:8080/CAST-RESTAPI-integrated/rest/","cast","cast",True)
domain_id = aip.getDomain("actionplatform_central")

sizing = aip.getSizing(domain_id,aip._sizing)
sizing_df = pd.DataFrame.from_dict(data=sizing, orient='index')
sizing_df[1] = None
sizing_df[1] = round((sizing_df[0]/sizing_df[0].sum()) * 100)
sizing_df[0] = sizing_df[0].map('{:,.0f}'.format)
sizing_df[1] = sizing_df[1].map('{:,.0f}%'.format)

ppt.fill_table('DocTable',sizing_df,orient='index')
ppt.save()

table_shape = ppt.get_shape_by_name('DocTable')
if table_shape.has_table:
    table=table_shape.table

    for r in range (0,len(table.rows)):
        text = data_top = sizing_df.head().index[r]
        run = ppt.merge_runs(table.cell(r,0).text_frame.paragraphs[0])
        run.text = run.text.replace(run.text,text)

    for r in range (0,len(table.rows)-1):
        row = table.rows[r]
        for c in range (0,len(row.cells)-1):
            cell = table.cell(r,c+1)
            run = ppt.merge_runs(cell.text_frame.paragraphs[0]) 
            text = sizing_df[c][r]
            run.text = run.text.replace(run.text,text)

        row1 = sizing_df[0][0]

        for key in sizing.keys():
            cell = table.cell(row,0)
            run = ppt.merge_runs(cell.text_frame.paragraphs[0]) 
            run.text = run.text.replace(run.text,key)
            row=row+1
"""
    


