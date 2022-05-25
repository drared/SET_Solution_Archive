### Solution details: 

### Name: Get File List Workbook
### Description/purpose: 
>A simple Excel tool that creates a list of files in a specified folder. 
This tool is useful when an editable file list is required. 
A renaming file feature is also included. Use excel to modify the file list as desired and then run
the macro to do the name changes. 



#NOTE: 
> File name changes are NOT reversible. Double check names are as desired and will not be duplicated. 

## Instructions
There are currently two functions this Excel tool performs: 
1) In the 'main' excel worksheet, the target folder is selected and the 'list files' button is clicked. 
This runs the the 'list file' routine and the file list is created in the same worksheet

2) In the 'Rename' worksheet, a list of filenames can be added along with a list of new names. 
Running the routine applies the new names. No validation is done and no 'undo' button exists so steps should be taken to ensure the new names are correct and can be applied without error. 
### Location (url):  

### Author: Dale Anderson
### Author page: https://sharedexceltools.com/contributor-dale-anderson/

# Future additions excpected
1. Apply some checks to make sure renaming is safe
2. Allow the renaming to be reversed
3. Add logs to keep track of changes
4. Allow files other than 'normal' to be selected (ie directories, hidden etc)
5. Allow subdirectory files to be listed
6. Show all file types
7. Create worksheet showing folder structure
8. Show file attributes
### Version 1.3
- Added renaming file feature on separate worksheet (requires copy and past of file names)


### Version 1.2
 - Added renaming worksheet and module to update file names

### Version 1.0.0
- Original version










