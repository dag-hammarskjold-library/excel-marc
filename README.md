#### Requirements:

* Install the latest version of [Python](https://www.python.org/downloads/windows/)

#### Installation:

```bash
setx MDB <MongoDB connection string>
git clone https://github.com/dag-hammarskjold-library/excel-marc
cd excel-marc
python -m venv env
.\env\Scripts\activate
pip install -r requirements.txt
deactivate
```

* Right-click on the "run.ps1" file and set its properties to open with Powershell by default

#### Usage:

* Double-click "run.ps1"
* Select the Excel file to convert
