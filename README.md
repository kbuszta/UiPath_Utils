# UiPath_Utils
Tiny utilities for RPA developers working with UiPath.

## CreateNewFile
This script creates a new file based on Sample.xaml. You can include the necessary activities required for every new workflow, such as Try Catch, Annotations, Comments, etc.

To use the script, run it and provide the name for the new workflow. 
The script will generate a new .xaml file by duplicating the Sample.xaml file, including starting and ending log messages. 
If you want to add the folder name and an underscore before the new name automatically, prefix the filename with '#' symbol.

Example:

1. Place the script and the Sample file in the 'Salesforce' folder.
2. Run the script and type '#AddCase'.
3. It will create a file named 'Salesforce_AddCase.xaml' in the Salesforce folder.
