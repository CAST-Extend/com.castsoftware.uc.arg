from cast_common.highlight import Highlight
from cast_arg.powerpoint import PowerPoint
from cast_common.logger import Logger,INFO
from pandas import DataFrame
from json import load,JSONDecodeError
from os.path import exists,abspath


class HLPage(Highlight):

    _ppt=None
    _log = None
    _output = None
    _benchmark = None

    _tag_prefix = None

    _text_replace_json = {}
    _text_replace_file = abspath('text_replace.json')

    @property
    def ppt(self) -> PowerPoint:
        return HLPage._ppt

    @property
    def tag_prefix(self) -> str:
        if self._tag_prefix is None:
            raise ValueError(f'tag prefix is not set')
        return self._tag_prefix
    @tag_prefix.setter
    def tag_prefix(self,value):
        self._tag_prefix=f'{value}_hl'
   
    @property
    def output(self) -> str:
        return HLPage._output
    
    @property
    def log(self) -> Logger:
        return HLPage._log
    
    @property
    def replace_options(self) -> dict:
        return HLPage._text_replace_json

    def __init__(self,output:str=None,ppt:PowerPoint=None,  
                 hl_user:str=None, hl_pswd:str=None,hl_basic_auth=None, hl_instance:int=0,
                 hl_apps:str=[],hl_tags:str=[], 
                 hl_base_url:str=None, 
                 log_level=INFO):
        
        if HLPage._log is None:
            HLPage._log = Logger('HLPage',level=log_level)    
        if self.ppt is None: 
            if ppt is None:
                raise AttributeError('PPTX must be defined in the first instance of HLPage')
            else:
                HLPage._ppt=ppt

        if self.output is None: 
            if output is None:
                raise AttributeError('output must be defined in the first instance of HLPage')
            else:
                HLPage._output=output

        super().__init__(hl_user, hl_pswd,hl_basic_auth, hl_instance,hl_apps,hl_tags, hl_base_url, log_level)

        if HLPage._benchmark is None:
            HLPage._benchmark = DataFrame(self._get(r'/benchmark'))

        if HLPage._text_replace_json == {}:
            if not exists(HLPage._text_replace_file):
                raise AttributeError(f'Text Replacement file not found: {HLPage._text_replace_file}')
            try:
                HLPage._text_replace_json = load(open(HLPage._text_replace_file))
            except JSONDecodeError as ex:
                raise AttributeError(f'Malformed text replacement file {HLPage._text_replace_file}: {ex}')



    def replace_from_options(self,tag,level):
        #loop through the text replacment json file     
        for t in self.replace_options:
            repl = f'{tag}_{t}'
            try:
                self.replace_text(repl,self.replace_options[t][level])
            except KeyError as ke:
                self.log.warning(f'Text replacement failure in {self.__class__} for [{repl}][{level}]')

    def replace_text(self, find_tag:str, data, shape=False, slide=None):
        """
                            Review PowerPoint Presentations                     
                             Performing Text Replacement                       
                                                                               
            The find_tag parameter is used to identify the text to be replaced
            The data parameter contains the replacement text, or data.    

            Use the shape and slide parameters to limit the search accordingly            
        """
        # if slide is None:
        #     slide = PowerPoint.ppt._prs.slides[3]
        replace_tag = f'{self.tag_prefix}_{find_tag}'
        self.log.debug(f'{replace_tag}: {data}')
        if shape:
            PowerPoint.ppt.replace_textbox(replace_tag, data, slide=slide)
        else:
            replace_tag_with_braces = f'{{{replace_tag}}}'
            PowerPoint.ppt.replace_text(replace_tag_with_braces, data, slide=slide)



