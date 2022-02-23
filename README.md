# Process Code

The python code (Process.py) process a given **xlsx** file, it can be provided as input or defined in the source code in the section below (input, output):
```
newExec = OrgTable("./Heinz GALTest.xlsx","./output.xlsx")
```
Both of the provided files need to be a valid .xlsx excel file.
The current code supports up to 10 levels of management.

## Install
The code is prepared to get installed automatically just running the following command:
```
sudo python3 setup.py install
```

## Dependencies
python >= 3.6 installed.
The code needs the following libraries intalled:
- seaborn
- openpyxl

These are installed automatically by the setup.py command. If a manuall install is needed use the requirements file provided:
```
pip3 install -r requirements.txt
```
## Output
The code will process the given file and order the subordinates in the output file, if the code is ran with the same output parameter it will overwrite the file.

## Contact
If you have questions or need assistance please reach the following email:
oscar.ruiz@softtek.com

![Python Logo](https://www.python.org/static/community_logos/python-logo.png "Sample inline image")