# Overview

This example python script takes the json output of a report created by [Sysreptor](https://github.com/Syslifters/sysreptor) and puts all findings into an xlsx file.

The worksheet is named findings, the client can use a dropdown to mark a finding as "Closed", "Ignored" or "Open". Various styling / formatting have been applied.

Please use as an example to create your own branded findings tracking file.

# Example
Output for the `Margherita Report Demo` Report.

![Example output](/example_xlsx.png "Findings")

# Installation

Create a virtualenv

```
virtualenv .venv
```

Install requirements

```
pip install -r requirements.txt
```

Export a report from the SysReptor projects and extracts its contents.

Rename the json file to:
`report_export.json`

Run the python script

```
python create_xlsx.py
```

# Notes
This is not a full CLI app where you can supply some parameter and everything is done for you.

It is just a demonstration how easy it is to create a customised, branded XLSX based findings list.

It is a great way to create a checklist/tracking of findings from a pentest report.

My suggestions: add another worksheet, put your company details in there, add a logo, add information from the JSON file such as the report date or similar.

Own it! Your client will appreciate it.

# License
[MIT](LICENSE)