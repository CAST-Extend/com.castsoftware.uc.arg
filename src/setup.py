from setuptools import setup,find_packages

setup (
    name="com.castsoftware.uc.arg",
    url='https://github.com/CAST-Extend/com.castsoftware.uc.arg',
    description="Assessment Report Generator",
#    install_requires=['pandas','python-pptx==0.6.18','com.castsoftware.uc.python.common>=0.1.6','IPython','requests','Jinja2'],
    package_data={'':['cause.json','Effort.csv']},
    packages=find_packages()
)
