#### Requirements:
* [Git for Windows](https://git-scm.com/download/win) 
* Latest version of [Python](https://www.python.org/downloads/windows/)

#### Installation:

```bash
git clone https://github.com/dag-hammarskjold-library/excel-marc
cd excel-marc
python -m venv env
.\env\Scripts\activate
pip install -r requirements.txt
deactivate
setx MDB <MongoDB connection string>
powershell Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser
```

* Right-click on the "run.ps1" file and set it to with Powershell by default
* Optional: right-click again on the "run.ps1" file and create a shortcut file. The shortcut file can be moved to any other directory, or the desktop for convenience.

#### Usage:

* Double-click "run.ps1" 
* Select the Excel file to convert

###### Spreadsheet format example:

|1.191a|1.191b|1.191c|2.191a|2.191b|2.191c|245a|245b|269a| 
|-|-|-|-|-|-|-|-|-|
|A/RES/1|A/|1|S/RES/1|S/|1|title|subtitle|date|
|A/RES/2|A/|1|S/RES/2|S/|1|title|subtitle|date|

###### Setting default values:

Default values can be set by using a spreadheet of values in a file called "defaults.xlsx" and placing the file in the root of this repo. Any values found in the defaults file will be added to all the records in a converted spreadsheet if the given field is not already in the record being converted.
