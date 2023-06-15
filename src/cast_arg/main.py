from argparse import ArgumentParser
from cast_arg.convert import GeneratePPT
from cast_arg.config import Config
# from convert import GeneratePPT
# from config import Config

__author__ = "Nevin Kaplan"
__email__ = "n.kaplan@castsoftware.com"
__copyright__ = "Copyright 2023, CAST Software"

if __name__ == '__main__':
    print('\nCAST Assessment Deck Generation Tool')
    print('Copyright (c) 2023 CAST Software Inc.\n')
    print('If you need assistance, please contact Nevin Kaplan (NKA) from the CAST US PS team\n')

    parser = ArgumentParser(description='Assessment Report Generation Tool')
    parser.add_argument('-c','--config', required=True, help='Configuration properties file')
    args = parser.parse_args()
    ppt = GeneratePPT(Config(args.config))
    ppt.save_ppt()

