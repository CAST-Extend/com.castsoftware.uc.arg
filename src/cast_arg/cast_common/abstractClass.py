from inspect import signature,getmembers,getsourcelines

__author__ = "Nevin Kaplan"
__copyright__ = "Copyright 2022, CAST Software"
__email__ = "n.kaplan@castsoftware.com"

class abstract():
    def validate_abstract(self,name):

        method={}
        sourcelines = getsourcelines(name)[0]
        for i,line in enumerate(sourcelines):
            line = line.strip()
            if line.split('(')[0].strip() == '@abstractclassmethod': # leaving a bit out
                nextLine = sourcelines[i+1]
                
                #get the function name
                name = nextLine.split('def')[1].split('(')[0].strip()
                #now the parameters
                params = []
                param_section = nextLine.split('(')[1].split(')')[0]
                for p in param_section.split(','):
                    if p == 'self': continue
                    params.append(p.strip())
                method[name]=params

        for m in method:
            ldc = hasattr(self, m) and callable(getattr(self, m))
            if ldc:
                for param in method[m]:
                    t = signature(getattr(self, m))
                    param_parts = param.split(':')
                    if '*' in param_parts[0]:
                        param_parts[0] = param_parts[0].split('*')[1]

                    if param_parts[0] in t.parameters:
                        concrete_param = t.parameters[param_parts[0]]
                    else:
                        concrete_param=None
                    if concrete_param ==None or concrete_param.name not in param:
                        raise AttributeError(f'\"{self.__name__}\" class \"{m}\" method is missing \"{param}\" parameter')
                    if len(param_parts) > 1 and param_parts[1].strip() not in str(concrete_param.annotation):
                        raise AttributeError(f'\"{self.__name__}\" class \"{m}\" method \"{param}\" parameter is type {concrete_param.annotation} not {param_parts[1]}')
