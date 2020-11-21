[![GitHub issues](https://img.shields.io/github/issues-raw/Agent6-6-6/Excel-VBA-XML-Export-Pre-Commit-Hook.svg?color=RED&style=flat-square)](https://GitHub.com/Agent6-6-6/Excel-VBA-XML-Export-Pre-Commit-Hook/issues)
[![GitHub issues-closed](https://img.shields.io/github/issues-closed-raw/Agent6-6-6/Excel-VBA-XML-Export-Pre-Commit-Hook.svg?color=brightgreen&style=flat-square)](https://GitHub.com/Agent6-6-6/Excel-VBA-XML-Export-Pre-Commit-Hook/issues?q=is%3Aissue+is%3Aclosed)
[![Black Code](https://img.shields.io/badge/code%20style-black-000000.svg?style=flat-square)](https://github.com/ambv/Black)


# Excel VBA & XML Export Pre-Commit Hook

### What do these Pre-commit scripts do?
When you commit an Excel file, these scripts extract all VBA normal modules, class modules & forms, and the customUI.xml & customUI14.xml files.

### Blog post explaining use/setup
See this blog post link for further information
[Excel…. Version control…. Git the hell out of here!](https://engineervsheep.com/2020/excel-git/)

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

If you want it to do something different then adapt it to your own needs.

### A bit about my Excel workflow
The included .gitignore file is setup to only commit files within the root directory of the repository, thats how I roll with Excel.

I usually store my excel spreadsheet templates in a filename format like `templatename - Rev X.xxx.xltm` in their own repository/directory. This helps me identify at a glance templates based on revision. I usually keep things like verification information (hand calculations/example outputs), other resources, etc in further subdirectories so all of the relevant information is together with the template. These sub-directories are not tracked by git.

I also have an addin which contains generic code which can be used within individual templates, this approach is taken so that if I update VBA code (say to reflect changes in structural engineering standards/design code equations) then this updated code is available to all workbooks that used the previous version of this VBA code (no updating of individual templates VBA code required). This centralised approach to storing common code saves considerable development time. The addin is distributed to users of my templates (stored/run from network drive).

The addin/templates contains a few custom ribbon tabs that contain buttons for executing custom code and groups together existing excel ribbon functions in a manner that reduces development time, especially with respect to formatting spreadsheets to styles consistent with company policies, etc.

If I am making small changes I'll just do it in a master branch and commit to github to update the code modules stored on github, if doing major changes I'll do a branch until things are finalised. You just need to keep in mind because of the binary nature of the Excel files, you cannot work in two branches as there is no way to merge the code from two competing branches into the excel file itself. This pre-commit only extracts the code as its stored within a file. If changes are made in both branches then you need to reconcile them by hand before committing/merging one of the branches.

### Prerequisites
Python 3.x (Tested as working with Python 3.8.x)

Excel (Tested with Excel/Microsoft 365. Excel is used for exporting the VBA modules)

An earlier version of this tool used/required oletools for extracting the VBA. However this was unable to extract both files for userforms.

To get the VBA script working you'll need to enable programmatic access to the VBA project within Excel. You can do this by going -> File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings -> activate checkmark on `Trust access to the VBA project object model`

### Next steps...
Make/steal/beg/borrow some way to import the extracted modules and customUI xml files back into an Excel file... in progress
