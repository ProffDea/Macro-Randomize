# Macro-Randomize
Libreoffice Calc Macro using the Basic programming language. Once the macro runs, it reads content of a specified text file then pastes a random line of content into a selected cell. It repeats this function for each selected cell. The **races.txt** file is an example file with the contents of some D&D races. This file can be replaced with a different text file with different content as long as the new text file's name replaces the old name in the **config.txt** file.

# Set Up Guide
- Create a new or open an existing Libreoffice Calc Spreadsheet
- Save the spreadsheet into a directory you wish to use
- Drop the **config.txt** and **races.txt** files into the directory the spreadsheet file is in. Keep the **macro_randomize.bas** in mind
- Have the spreadsheet in the directory opened in Libreoffice Calc
- In Libreoffice Calc, go to Tools > Macros > Organize Macros > Basic... a dialog box should open
- On the right panel, click New. This should open a new window which is the IDE
- In the script editor panel of the IDE, erase everything in it to where it's blank
- Find the button called **Import Basic** along the top of the window and click it
- Select the **macro_randomize.bas** file from wherever you saved it from downloading. You should see the code for the macro flood the script editor

The set up is complete. In the Libreoffice Calc spreadsheet, you should be able to now go to Tools > Macros > Run Macro... to open the dialog box then go to wherever you created the new macro in which according to this guide would be in My Macros > Standard > Module1. Select **RandomizeContent** and click Run in order to use the macro.

Make sure to edit the contents of the **races.txt** file to your liking for the macro to use. If you change the name of the file, make sure you change the name in **config.txt** as well.
