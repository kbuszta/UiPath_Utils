﻿# UiPath_Utils

Tiny utilities for RPA developers working with UiPath.

## CreateNewFile

Script to create new file based on Sample.xaml, where you can put necessary activities needed for every new workflow (i.e. Try Catch, Annotation, Comments etc.).

* Run script, put the name of new workflow and script will make new .xaml file by copying Sample.xaml file (in this case with starting and ending log messages). 
* Type '#' before new filename to automatically add folder name and "_" before the new name. 

Example: 

1. Put script and Sample file into folder 'Salesforce'. 
2. Run script, type '#AddCase' - it will create file 'Salesforce_AddCase.xaml' in the Salesforce folder.
