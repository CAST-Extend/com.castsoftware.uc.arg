from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.chart.plot import BarPlot
from pptx.parts.chart import ChartPart
from pptx.parts.embeddedpackage import EmbeddedXlsxPart
from pptx.dml.color import ColorFormat, RGBColor
from pptx.enum.dml import MSO_THEME_COLOR

import util

from copy import deepcopy
import six
import pandas as pd
import logging


class PowerPoint:
    _input = None
    _output = None
    _prs = None

    def __init__(self,input, output):
       self._input=input 
       self._output=output

       self._prs = Presentation(self._input)

       self._logger = logging.getLogger(__name__)
       shandler = logging.StreamHandler()
       formatter = logging.Formatter('%(asctime)s - %(filename)s [%(funcName)30s:%(lineno)-4d] %(levelname)-8s - %(message)s')
       shandler.setFormatter(formatter)
       self._logger.addHandler(shandler)

    def replace_risk_factor(self, grades, app_no=0, search_str=None):
        if search_str == None:
            search_str=f'{{app{app_no+1}_risk_'
        
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
                                        self._logger.debug(f'invalid key: {g}')
                                        break;

    def replace_grade(self, grades,app_no=0, search_str=None):
        if search_str == None:
            search_str=f'{{app{app_no}_grade_'

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
                                    grade_str = str(round(grade,2))
                                    if len(grade_str) < 4:
                                        grade_str = grade_str + '0'
                                    cur_text = cur_text.replace(f'{search_str}{g}}}',grade_str)
                                    run.text = cur_text

                                    if cur_text == grade_str:
                                        color = self.get_grade_color(grade)
                                        run.font.color.rgb = color
                                        shape.line.width=0
                                        box_name = f'{search_str[1:]}{g}'
                                        box = self.get_shape_by_name(box_name,slide)
                                        if box != None:
                                            box.line.color.rgb = color
                                            self.update_grade_slider(f'{search_str[1:]}chart_{g}', [grade])

    def update_grade_slider(self, name,data):
        shape = self.get_shape_by_name(name)
        try:
            if shape.has_chart:
                chart_data = CategoryChartData()
                # chart_data.categories = titles
                
                chart_data.categories = 'grade'
                chart_data.add_series('Series 1', data)
                shape.chart.replace_data(chart_data)
        except AttributeError:
            self._logger.debug(f'Invalid template configuration: {name}')

    def get_grade_color(self,grade):
        rgb = 0
        if grade > 3:
            rgb = RGBColor(0,176,80) # light green
        elif grade <3 and grade > 2:
            rgb = RGBColor(214,142,48) # yellow
        else:
            rgb = RGBColor(255,0,0) # red
        return rgb


    def replace_text (self, search_str, repl_str):
        if len(str(repl_str)) == 0:
            repl_str = 'TBD'
        for slide in self._prs.slides:
            self.replace_slide_text(slide, search_str, repl_str)

    def replace_slide_text (self, slide, search_str, repl_str):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    self.replace_paragraph_text(paragraph,search_str,repl_str)
            elif shape.has_table:
                tbl=shape.table
                row_count = len(tbl.rows)
                col_count = len(tbl.columns)
                for r in range(0,row_count):
                    for c in range(0, col_count):
                        cell = tbl.cell(r,c)
                        for paragraph in cell.text_frame.paragraphs:
                            self.replace_paragraph_text(paragraph,search_str,repl_str)


    def replace_shape_name (self, slide, search_str, repl_str):
        for shape in slide.shapes:
            if shape.name.find(search_str) != -1: 
                shape.name = shape.name.replace(search_str,repl_str)



    def replace_paragraph_text (self, paragraph, search_str, repl_str):
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

    def get_shape_by_name(self, name, use_slide=None):
        rslt = None

        slides = self._prs.slides
        if use_slide != None:
            slides = [use_slide] 

        for slide in slides:
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
        if len(paragraph.runs) > 0:   
            run = paragraph.runs[0]
            run.text = cur_text
        else: 
            run = paragraph.add_run()
        return run

    def update_chart(self, name,df):
        shape = self.get_shape_by_name(name)
        if shape != None:
            titles = list(df.index.values)
            data = df.to_numpy()
            for i in range(0,len(data)):
                if (isinstance(data[i],str)):
                    data[i] = int(data[i].replace(',',''))

            if shape.has_chart:
                chart_data = CategoryChartData()
                chart_data.categories = titles
                chart_data.add_series('Series 1', data)
                shape.chart.replace_data(chart_data)


    def update_table(self, name, df, include_index=True, background=None, has_header=True):
        table_shape = self.get_shape_by_name(name)
        if table_shape != None and table_shape.has_table:
            table=table_shape.table

            colnames = list(df.columns)

            # Insert the row zero names
            if include_index:
                for col_index, col_name in enumerate(df.index):
                    try:
                        if has_header:
                            cell = table.cell(col_index+1,0)
                        else:
                            cell = table.cell(col_index,0)
                        text = str(col_name)
                        run = self.merge_runs(cell.text_frame.paragraphs[0])
                        run.text = text
                    except IndexError:
                        self._logger.warn('index error in update_table while setting df index')
            rows, cols = df.shape
            if background:
                cols = cols-1

                colors = df[background]
                for row in range(rows):
                    rgb = colors.iloc[row].split(",")
                    for col in range(cols):
                        try:
                            if has_header:
                                cell = table.cell(row+1,col)
                            else:
                                cell = table.cell(row,col)
                            cell.fill.fore_color.rgb = RGBColor(int(rgb[0]), int(rgb [1]), int(rgb[2]))
                        except IndexError:
                            self._logger.warn('index error in update_table while setting background color')



            m = df.values
            for row in range(rows):
                for col in range(cols):
                    val = m[row, col]
                    text = str(val)
                    
                    if include_index:
                        tbl_col=col+1
                    else:
                        tbl_col=col

                    try:
                        if has_header:
                            cell = table.cell(row+1,tbl_col)
                        else:
                            cell = table.cell(row,tbl_col)

                        run = self.merge_runs(cell.text_frame.paragraphs[0])
                        run.text = text
                    except IndexError:
                        self._logger.warn('index error in update_table while setting values')

    def replace_block(self, begin_tag, end_tag, repl_text):
        for slide in self._prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        text = paragraph.text
                        if text.find(begin_tag)!=-1:
                            run=self.merge_runs(paragraph)
                            run_text = run.text
                            text_prefix = text[:run_text.find(begin_tag)]
                            text_suffix = text[run_text.find(end_tag)+len(end_tag):]
                            new_text = text_prefix + repl_text + text_suffix
                            run.text = run.text.replace(run_text,new_text)

    def copy_block(self, tag, prefix, count):
        search_start = f'{{{tag}}}'
        search_end = f'{{end_{tag}}}'

        block = []

        found=False
        for slide in self._prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        text=paragraph.text
                        if not found and text.find(search_start)!=-1:
                            #is the end in the same paragraph?
                            if text.find(search_end)!=-1:
                                run=self.merge_runs(paragraph)
                                old_text = run.text
                                text_prefix = text[:old_text.find(search_end)]
                                text_suffix = text[old_text.find(search_end):]

                                sub_text = ""
                                for app_no in range (2,count+1):
                                    temp = old_text[old_text.find(search_start)+len(search_start):old_text.find(search_end)]
                                    for p in prefix:
                                        temp = temp.replace(f'{p}1',f'{p}{app_no}')
                                    sub_text = sub_text + ", " + temp
                                
                                new_text = text_prefix.replace(search_start,'') + sub_text + text_suffix.replace(search_end,'')
                                #new_text = text_prefix + sub_text + text_suffix

                                run.text = run.text.replace(old_text,new_text)
                            else:
                                found=True
                        elif found and text.find(search_end)!=-1:
                            found=False
                            for app_no in range(2,count+1):
                                self.paste_block(block, shape.text_frame,app_no)
                            if paragraph.text==search_end:
                                self.delete_paragraph(paragraph)
                            block=[]
                        if found:
                            if paragraph.text==search_start:
                                self.delete_paragraph(paragraph)
                            else:
                                block.append(paragraph)
#        self.replace_text(search_start,"")
#        self.replace_text(search_end,"")

    def paste_block(self,block, text_frame,app_no):
        start_at = block[-1]

        for b in block:
            p = text_frame.add_paragraph()
            p.alignment = b.alignment
            p.line_spacing = b.line_spacing
            p.level = b.level

            for r in b.runs:    
                run = p.add_run()
                run.text = deepcopy(r.text)
                font = run.font
                font.name = r.font.name
                font.size = r.font.size
                font.bold = r.font.bold
                font.italic = r.font.italic
                if hasattr(r.font.color, 'rgb'):
                    font.color.rgb = r.font.color.rgb
                # else:
                #     font.color.theme_color = r.font.color.theme_color 

            run = self.merge_runs(p)
            run.text = run.text.replace("{app1_",f'{{app{app_no}_')
            run.text = run.text.replace("{end_app1_",f'{{end_app{app_no}_')

    def replace_loc(self, loc, app_no):
        loc_short = "{0:,.0f}".format(loc) 
        if loc > 1000000:
            loc_short = "~{0:,.2f} MLOC".format(loc/1000000) 
        elif loc < 100000:
            loc_short = "~{0:,.0f} KLOC".format(loc/1000) 
        self.replace_text(f'{{app{app_no}_loc}}',f'{loc:,.0f}')
        self.replace_text(f'{{app{app_no}_loc_short}}',loc_short)

        size_catagory = 'small'
        if loc > 1000000:
            size_catagory = 'very large'
        elif loc > 500000:
            size_catagory = 'large'
        self.replace_text(f'{{app{app_no}_loc_category}}',size_catagory)

    def duplicate_slides(self, app_cnt):
        for cnt in range(2,app_cnt+1):
            for idx, slide in enumerate(self._prs.slides):
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            if paragraph.text == "{app_per_page}":
    #                            self.replace_slide_text(slide,"{app_per_page}","")
                                new_slide = self.copy_slide(idx)
                                self.replace_slide_text(new_slide,"{app_per_page}","")
                                self.replace_slide_text(new_slide,"{app1_",f'{{app{cnt}_')
                                self.replace_shape_name(new_slide,"app1_",f'app{cnt}_')
        for idx, slide in enumerate(self._prs.slides):
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        if paragraph.text == "{app_per_page}":
                            self.replace_slide_text(slide,"{app_per_page}","")

    
    def copy_slide(self,index):
        source = self._prs.slides[index]
        blank_slide_layout = source.slide_layout
        dest = self._prs.slides.add_slide(blank_slide_layout)

        for shp in source.shapes:
            el = shp.element
            newel = deepcopy(el)
            dest.shapes._spTree.insert_element_before(newel, 'p:extLst')

        for key, value in source.part.rels.items():
            # Make sure we don't copy a notesSlide relation as that won't exist
            if "notesSlide" not in value.reltype:
                target = value._target
                # if the relationship was a chart, we need to duplicate the embedded chart part and xlsx
                if "chart" in value.reltype:
                    partname = target.package.next_partname(
                        ChartPart.partname_template)
                    xlsx_blob = target.chart_workbook.xlsx_part.blob
                    target = ChartPart(partname, target.content_type,
                                    deepcopy(target._element), package=target.package)

                    target.chart_workbook.xlsx_part = EmbeddedXlsxPart.new(
                        xlsx_blob, target.package)

                dest.part.rels.add_relationship(value.reltype,
                                                target,
                                                value.rId)

        return dest

    def remove_empty_placeholders(self):
        for slide in self._prs.slides:
            for placeholder in slide.shapes.placeholders:
                if placeholder.has_text_frame and placeholder.text_frame.text == "":
                    sp = placeholder._sp
                    sp.getparent().remove(sp)

    def get_shape_parent(self,shape):
        rslt = None
        if hasattr(shape,'_parent'):
            rslt = shape._parent
        return rslt

    def delete_paragraph(self,paragraph):
        p = paragraph._p
        parent_element = p.getparent()
        parent_element.remove(p)

    def get_page_no(self,shape):
        page_no = 0
        if shape:
            while True:
                shape = self.get_shape_parent(shape)
                if type(shape).__name__ == 'Slide':
                    break;
            
            page_no = self._prs.slides.index(shape) 
        return page_no


    def delete_run(self,run):
        r = run._r
        r.getparent().remove(r)

    def save(self):
        self._prs.save(self._output)

"""
from restCall import AipRestCall
from restCall import AipData

aip_rest = AipRestCall("http://sha-dd-console:8080/CAST-RESTAPI-integrated/rest/","cast","cast",True)
project = "Blackhawks"    
apps = ["mobile_doorman_android","mobile_doorman_ios"] 
app_cnt = len(apps)
aip_data = AipData(aip_rest,project, apps)
grade_all = aip_data.get_app_grades(app_id)

app_no=0
ppt.replace_grade(grade_all,app_no+1)

ppt.save()
"""