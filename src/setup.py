from setuptools import setup,find_packages

setup (
    name="com.castsoftware.uc.arg",
    url='https://github.com/CAST-Extend/com.castsoftware.uc.arg',
    author_email='n.kaplan@castsoftware.com',
    description="Assessment Report Generator",
    install_requires=['pandas','python-pptx==0.6.18','com.castsoftware.uc.python.common>=0.1.6','IPython','requests','Jinja2'],
    package_data={'':['cause.json','Effort.cvs']},
    packages=find_packages()
)
