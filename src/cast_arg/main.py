from argparse import ArgumentParser
from cast_arg.convert import GeneratePPT
from cast_arg.config import Config
from pkg_resources import get_distribution

# from convert import GeneratePPT
# from config import Config

__author__ = "Nevin Kaplan"
__email__ = "n.kaplan@castsoftware.com"
__copyright__ = "Copyright 2023, CAST Software"

if __name__ == '__main__':
    version = get_distribution('com.castsoftware.uc.arg').version
    print(f'\nCAST Assessment Deck Generation Tool (ARG), v{version}')
    print(f'com.castsoftware.uc.python.common v{get_distribution("com.castsoftware.uc.python.common").version}')
    print('Copyright (c) 2023 CAST Software Inc.')
    print('If you need assistance, please contact oneclick@castsoftware.com')

    parser = ArgumentParser(description='Assessment Report Generation Tool')
    parser.add_argument('-c','--config', required=True, help='Configuration properties file')
    args = parser.parse_args()
    config=Config(args.config)
    GeneratePPT(config)._ppt.save()
    
