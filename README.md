This is created utilizing various Python libraries created by other individuals, including:
  - Pandas
  - sqlite3
  - argparse
  - numpy
  - datetime
  - plotly

In order to run this program the following command line process needs to be run in an environment with those packages installed.
**python Report_Generic.py -db "Path\To\DBFile.db**

Optional Flags:
  -op | --output                 Indicate Output Directory, Default is C:\ProgramData\Generic\Reports\{JobID}\{EVDNUM}
  -nodt | --nodirectorytree      Exclude Directory Tree Report
