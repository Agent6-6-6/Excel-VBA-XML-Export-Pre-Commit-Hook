# Excel VBA & XML Export Pre-Commit Hook

To be added to further & updated when I get a chance to add more detail!

### Do this for now:-
Basically put the three pre-commit files within the .git\hooks directory of a repository

Save/copy any excel file into the repositories root directory, commit any changes to tracked files and the scripts take care of extracting the VBA (forms/modules/class modules) and customUI XML into *.XML & *.VBA subdirectories and adds the extracted VBA & XML files to the commit. Effectively adding version control for VBA modules and the customUI.xml files (Ribbon) within Excel spreadsheets or add-ins. If no VBA or customUI is present, then no subdirectories are created for that particular component. Every subsequent change to the Excel file is picked up by git once you save. When you commit the changes, the pre-commit hook runs the scripts again to extract the VBA & XML, then they are added/removed from the commit, rinse and repeat to infinity. 

Some Excel test files with varying VBA and XML content are provided to test the functionality if you clone the repository, simply unzip the files and do a commit to see what is going on.

My first time using python so if it doesn't work for you, I'm probably not going to be of much help. But if its any consolation its exactly what I was after for my own workflow.

### A bit about my Excel workflow
.gitignore file is setup to only commit files within the root directory of the repository, thats how I roll with Excel. More to follow on this I guess so you can customise the scripts to your own workflow.

### Prerequisites
python 3.x (My first time using python so no idea if it will work under python 2.x, test it and let me know)

oletools (for extracting the VBA)

To get the VBA script working you'll need to enable programmatic access to the VBA project within Excel. You can do this by going -> File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings -> activate checkmark on 'Trust access to the VBA project object model' 

### Next steps...
Make/steal/beg/borrow some way to import the extracted modules back into Excel file.

Make a better readme (obviously)......
