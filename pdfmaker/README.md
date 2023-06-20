# How to use the pdfmaker module

## Prerequesites 

- Your completed gradesheets as xlsx files all in one folder
- Close all other excel files

## Start
- Open the module by double clicking the module name on the right

![](img/img2.png)

- You will need to edit the code in the editor that should be open on your screen.

- Change the path of the file with the completed gradesheets to the correct location on your device (line 15; tip: go to the file explorer and navigate to where the template is saved, right click the address bar at the top click copy address as text). Make sure there is a backslash at the end.

- Line 37 is currently set that it output the pdfs to the folder that the excel files are in. If you want a different one, comment out (with a ') line 37 and uncomment (by deleting the ') and edit line 39. Also Change line 42 to `FilePath2`

- Press play or F5 to run

![](img/img3.png)
