# Assessment Deck Generation Tool (ARG)
Leveraging the power of CAST MRI and Highlight REST API's, ARG generates a PowerPoint deck in accordance with a designated template.  

## First time installation
1.	Create a folder to install the software, for example c:\ARG.
2.	In a command prompt, navigate to the recently established folder: c:\ARG.
3.	Create a virtual environment by running the command: PY -m venv .venv.
4.	Activate the virtual environment with this command: .\.venv\Scripts\activate.
5.	Finally, install ARG from the PYPI environment with this command: pip install com.castsoftware.uc.arg.

### Upgrade Process
1.	In a command prompt, navigate to the folder created during installation: c:\ARG.
2.	Activate the virtual environment with this command: .\.venv\Scripts\activate.
3.	Finally, install ARG from the PYPI environment with this command: pip install com.castsoftware.uc.arg==<full-version-id> --upgrade

## Usage
The script is designed to run on the command line using one parameter, --config, used to identify the project and applications to include in the presentation.

    <ARG install>\.venv\Scripts\activate
    py -m cast_arg.main --config config.properties 

# Configuration file
The Assessment Deck Generation Tool is configured using a json formatted file.  The file is divided into three sections, general configuration, application list and rest configuration.  
#### General Configuration
The general configuration section contains four parts,  
* company - the name of the company as it will appear on the report
* project name - the name of the project as it will appear on the report
* template - the absolute location of the report template
* output - the absolute location of the report output folder

#### Application list
This section contains a list of all applications that will be used in the report.  It is divided into three parts,  
* aip - the application schema triplet base
* highlight - the name of the application as it appears in the Highlight portal
* title - the name of the application to be used in the report

aip is the only required in this section.  If Highlight or title are empty or left out the aip value will be used.

#### REST Configuration
The REST configuration is divided into two parts, AIP and Highlight.  Both configurations contain 
* Active - toggle to turn REST service on/off
* URL - the service base URL 
* user - login user id
* password - login password (non-encrypted)

The Highlight configuration contains an additional field, instance, which refers to the login instance id.

#### Sample configuration

    {
        "company":"MMA Company Name",
        "project":"Project Name",
        "template":"C:/com.castsoftware.uc.arg/Template.pptx",
        "output":"C:/output/folder",
        "application":[
            {
                "aip":"triplet_base_name",
                "highlight":"highlight_name",
                "title":"application title"
            },
            {
                "aip":"triplet_base_name",
                "highlight":"highlight_name",
                "title":"application title"
            }, ...
        ],
        "rest":{
            "AIP":{
                "Active":true,
                "URL":"http://<URL>:<Port>/rest/",
                "user":"user_name",
                "password":"user_pasword"
            },
            "Highlight":{
                "Active":true,
                "URL":"https://<URL>/WS2/",
                "user":"user_name",
                "password":"user_pasword",
                "instance":"instance id"
            }
        }
    }


### Output
A single PowerPoint deck and one excel spread sheet is generated for each application configured in the properties file. The Deck is organized with an executive summary, one section per application containing detailed information and an appendix.  The excel sheets are hold the application action plan information and divided into two tabs, summary, and detail.  In the event an action plan is not configured, for the application, the sheets will be empty.    

This version of the script focus on CAST AIP content but not formatting.  For example, the executive summary section currently has a bullet point for only the first application.  This section concentrates on what manual formatting is required to make the deck client ready.  

#### <ins>Manual Formatting Required</ins>
The current version of this script only includes CAST AIP.  All items related to Open Source and Appmarq must be manually added.  This shortfall will be addressed in future releases. 

The following is a list of all known formatting issues:  

| Template Section | Issue | Workaround |
| -----------------| ----- | ---------- |
| Executive Summary | In the first paragraph, the number of applications is currently hard coded to “two”.  | Update his information manually. |
| Executive Summary | When generating a document containing more than one application the bullet points do not extend past the first application. | Use the paintbrush on the first bullet point to format the remaining applications. |
| Executive Summary | When generating a document containing more than one application the text is overrunning the page  | Select the textbox, right click the mouse, and select “Format Shape”.  In the “Text Options” tab and click on Textbox option.  Finally click on the “Shrink text on overflow” option. |
| Application Health | The second and third bullet points are not including grade improvements.  | Get the grade improvement scores from Action Plan Optimizer.  |
| Appendix | The appendix is not being included  | Manually insert the appendix from the appendix deck found in teams |


