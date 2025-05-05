import logging,sys
from logging import LogRecord, DEBUG, INFO, WARN, ERROR, DEBUG, CRITICAL

class Logger():
    loggers = set()

    def __init__(self,
                 name:str="general", 
                 level:int=INFO, 
                 format:str = '%(asctime)s [%(levelname)-s] %(message)s',
                 file_name:str=None,
                 file_mode:str='a',
                 console_output:bool=True):

        # Initial construct.
        self.format = format
        self.level = level
        self.name = name

        # Logger configuration.
        self.console_formatter = logging.Formatter(self.format)
        if console_output:
            self.console_logger = logging.StreamHandler(sys.stdout)
            self.console_logger.setFormatter(self.console_formatter)

        # Complete logging config.
        self.logger = logging.getLogger(name)
        if name not in self.loggers:
            self.loggers.add(name)
            self.logger.setLevel(self.level)
            if console_output:
                self.logger.addHandler(self.console_logger)
            if not file_name is None:
                fileHandler = logging.FileHandler(file_name,file_mode)
                fileHandler.setFormatter(self.console_formatter)
                self.logger.addHandler(fileHandler)

    @property    
    def is_debug(self):
        return self.logger.isEnabledFor(logging.DEBUG)

    def set_level(self,level):
        self.logger.setLevel(level)

    def debug(self, msg, *args, **kwargs):
        self.logger.debug(msg, *args, **kwargs)

    def info(self,msg, *args, **kwargs):
        self.logger.info(msg, *args, **kwargs)

    def warning(self, msg, *args, **kwargs):
        self.logger.warning(msg, *args, **kwargs)

    def error(self, msg, *args, **kwargs):
        self.logger.error(msg, *args, **kwargs)
