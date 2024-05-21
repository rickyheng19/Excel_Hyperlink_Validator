Excel Hyperlink Validator
-A simple python script to go through an excel file and validate if the hyperlinks work or 404s.
Uses libraries:
-pandas - for handing excel files
-requests - for checking the link statuses
-tkinter - for creating a simple UI

Prerequisites:
-An excel file containing a list of hyperlinks. The name of the links field must be "Link", and the column over must be "Status".

Example:

-----------------------------------------------------
| Link                    | Status                  | 
-----------------------------------------------------
| https://www.google.com/ |                         |
-----------------------------------------------------

Simply run the script, press the "Load and Validate" button, and wait.
Once the file is finished validating, it'll prompt you to save the new file.
