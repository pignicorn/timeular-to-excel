# Info

The script extracts entries via the Timeular API (v3) and writes these into an Excel file.

# Installation

1. Clone this directory: `git clone https://github.com/DrRetsiemsuah/timeular-to-excel`

2. Install requirements via python package manager `pip3 install -r requirements.txt`.

	```
	requests
	json
	xlsxwriter
	argparse
	```

3. Rename the `sample-api-key.json` file: `mv sample-api-key.json api-key.json`
4. Go to the profile page of Timeular to generate api key and api secret: https://app.timeular.com/#/settings/account. Copy & paste the values into the `api-key.json` file.

# Usage

Should be self explaining - therefore, refering to help message only.

```
usage: timeular-api.py [-h] -p PROJECT [-lm] [-cm] [-sd STARTDAY] [-ld LASTDAY]

Timeular export...

optional arguments:
  -h, --help            show this help message and exit
  -p PROJECT, --project PROJECT
                        Define the project name, as it is in timeular. Use "all" to extract all projects.
  -lm, --lastmonth      Extract all times from last month.
  -cm, --currentmonth   Extract all times from current month.
  -sd STARTDAY, --startday STARTDAY
                        First day of the period to capture.
  -ld LASTDAY, --lastday LASTDAY
                        Last day of the period to capture.
```