# Excel VBA & XML Export Pre-Commit Hook

To be added to further & updated when I get a chance to add more detail!

### What do these Pre-commit scripts do?
When you commit an Excel file, these scripts extract all VBA normal modules, class modules & forms, and the customUI.xml & customUI14.xml files. 

### Why did I create these scripts
Basically I was looking for a way to automate a way of using the version control offered by Git/Github with Excel files and the VBA code modules stored within an Excel file. 

I'm a heavy user of VBA in engineering templates and often tweak code in individual templates that must then be transferred to other templates. Exporting VBA modules by hand or copying sections of code within the VBE is pretty cumbersome, extracting the xml files from the excel archives by hand is even more cumbersome.

The only thing I could find using good old Google was this one page at [xltrail](https://www.xltrail.com/blog/auto-export-vba-commit-hook), which sort of did at a high level what I was after. So with the ideas given in the code provided and some tweaking, I set about creating something that worked for my particular workflow. 

### How to use these scripts:-
Basically put the three pre-commit files within the `.git\hooks` directory of a repository.

Save/copy any Excel file (or multiple Excel files) into the repositories root directory, commit any changes to tracked files and the scripts take care of extracting the VBA (forms/modules/class modules) and customUI XML files into *.XML & *.VBA subdirectories and adds the extracted VBA & XML files to the commit. 

This effectively adds version control for VBA modules and the customUI.xml files (Ribbon) within Excel spreadsheets or add-ins. If no VBA or customUI is present, then no subdirectories are created for that particular component. Every subsequent change to the Excel file is picked up by git once you save as a change. When you commit these changes, the pre-commit hook runs the scripts again to extract the VBA & XML, then they are added/removed from the commit, rinse and repeat to infinity. If you delete or rename an excel file the associated subdirectories are also removed. 

If you want every new repository to use the pre-commit files, copy them into the `\Program Files\Git\mingw64\share\git-core\templates\hooks` directory in Windows.

Some Excel test files with varying VBA and XML content are provided to test the functionality if you clone the repository, simply unzip the files and do a commit to see what is going on.

My first time using python so if it doesn't work for you, I'm probably not going to be of much help. But if its any consolation its exactly what I was after for my own workflow, want it to do something different then adapt it to your own needs.

### A bit about my Excel workflow
.gitignore file is setup to only commit files within the root directory of the repository, thats how I roll with Excel. More to follow on this I guess so you can customise the scripts to your own workflow.

### Prerequisites
python 3.x (My first time using python so no idea if it will work under python 2.x, test it and let me know)

oletools (for extracting the VBA)

To get the VBA script working you'll need to enable programmatic access to the VBA project within Excel. You can do this by going -> File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings -> activate checkmark on `Trust access to the VBA project object model`

### Next steps...
Make/steal/beg/borrow some way to import the extracted modules and customUI xml files back into an Excel file.

Make a better readme (obviously)......
