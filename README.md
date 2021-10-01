# Assessment Deck Generation Tool
The script is used to generate a PowerPoint deck from the provided template. This version includes part of the executive summary, the health and action plan pages.  

## Installation
* Download and unpack the latest release of the Assessment Deck  Generation Tool from https://github.com/CAST-Extend/com.castsoftware.uc.arg/releases
   * Source Code (zip) 
   * com.castsoftware.uc.python.common.zip 

* Unpack the Source Code zip file (arg)
* Unpack the com.castsoftware.uc.python.common.zip into a separate folder
* Install/update required third party packages. 
    * Open a command prompt
    * Go to the Source Code folder 
    * run: pip install -r requrements.txt
* make sure the <com.castsoftware.uc.python.common>/src folder is included in the PYTHONPATH enviroment variable

## Usage
The script is designed to run on the command line using one parameter, --config, used to identify the project and applications to include in the presentation.

    py convert.py --config config.properties 

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


