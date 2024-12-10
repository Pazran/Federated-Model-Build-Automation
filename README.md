# Federated-Model-Build-Automation
Using Powershell to automate the process of model federation (NWF and NWD files) for the Intel Pelican Project (Exyte Malaysia).
Disclaimer: This is published on Github for learning.

## SYNOPSIS
Builds the Models of the project S96029 (Pelican):
*	Design Models
*	Construction Models
*	Existing Models
*	Federated Models (all the three above combined)

## REQUIREMENTS
Below list is what is required to run the script properly.
NOTE: All requirements, except Exyte BIM360 Integration App, are automatically checked by the script once you run it.

## HARDWARE
### MEMORY CAPACITY (RAM)
Minimum: 32GB  
Recommended: 128GB  
With less than the recommended value, updating/creating big models (e.g., PG-FM  models) may not be possible.

### DISK FREE SPACE
Minimum: 20GB  
NOTE: Automatically checked once you run the script.  
With less than the recommended value, updating/creating big models (e.g., PG-FM  models) may not be possible.

## SOFTWARE
### EXYTE BIM360 INTEGRATION APP (CURRENTLY SUSPENDED)
For: Automatic synchronization between BIM 360 and Exyte server  
The Exyte BIM360 Integration App is an external tool (not installed on your machine) that must be running with the proper configuration.

### AUTODESK NAVISWORKS
For: Model Builds  
Tested versions: Manage 2021, Manage 2022, Manage 2023 (Not properly working), Manage 2024  
NOTE: Automatically checked once you run the script.

## NETWORK
\\\A1300564\D\S13_BIM-VDC  
For: Building the models  
NOTE: Automatically checked once you run the script.  
Network path must be accessible.
 
## DOCUMENT
### MASTER MODEL TRACKER
For: Checking new and retired model files  
Spreadsheet must be available and accessible at the link  TEMPORARY NAME FOR JANET’S MODEL TRACKER

## USAGE
To start the Pelican Federated Model build automation, run the script Pelican_build_latest_federated_model.bat.  
Other useful scripts:
*	More conservative script Pelican_build_latest_federated_model_conservative.bat will rebuild affected NWF and NWD for modified and retired model. It will take around 5-7hrs average.
*	To force rebuild all Federated Models for maintenance use Pelican_force_build_all_nwf_and_models.bat. It will take around 6-8hrs average.
*	To generate text files to reflect the content of each NWF/NWD files, run Pelican_generate_latest_nwf_textfile.bat

## SCRIPT WORKFLOW (WHAT IT DOES)
All the steps below are done automatically when you run the script. They are only a detailed description of its workflow.  
1.	Backup previous NWC models  
    - Check if Archived folder exist. If not, the script will create the folder
    - Create folder with the date the script running as the folder name
    - Backup all the NWC models into this folder

2.	Copy new and updated NWC models
    - Check if there is new updated NWC models in BIM Machine
    - Copy NWC models into the temporary build folder

3.	Move models with wrong naming convention  
    - Check if Rejected folder exist. If not, the script will create one
    - Move all the models

4.	Generate NWF files  
    - Only appends a file if it is following the naming convention
    - Only generates a building model (NWF) if really needed (e.g., at least one of the below is true):
       - There is a new version of a file previously appended
       - A file previously appended does not exist anymore
       - A file new file is available to be appended (not part of previous build)
    - Update viewpoints and Search sets for newly generated NWF

5.	Generate Federated Model (NWD)  
    - Export all list of federated models to be rebuild into text files
    - Read generated text files from previous step.
    - Generate all federated model in the list.

## AUTHOR
Lawrenerno Jinkim
Janet Jasintha Lopez
