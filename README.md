# Extracting Connections from Excel Files

A repo that allows for the extraction of any connections within an excel spreadsheet, 
or a floder that contains excel spreadsheets.

In order to store results as a dataframe please use pd.DataFrame.from_dict(results, orient="index").
You can also store any files that failed to open using the return array.

## How to use

### v1 

Version 1 uses win32 to open the excel file locally and then tries to find the connections within that file
using the Conncections property of the workbook. 
Here is an example on how to use v1:

```python

from search_connections import getConnectionsFromExcelFiles
import pandas as pd

example_path = "This is a path to an excel file or a folder containing excel files"
extractor = getConnectionsFromExcelFiles(path=example_path)
test, failed_files = extractor.get_connections_from_excel()

pd.DataFrame.from_dict(test, orient="index").to_csv("test_1.csv")
pd.Series(data=failed_files, name="failed_files").to_csv("test_1_failed_files.csv")

```

### v2

Version 2 uses the idea that xlsx files are just a collection of underlying xml files that are ziped together.
Thus this version copies the xlsx file in as a zip file, checks if there is a connections.xml file
within the new zip file, and then reads the connections file as a string that is then stored in a dictionary.

Here is an example on how to use v2:

```python
import pandas as pd
from search_connections_v2 import getConnectionsFromExcelFiles

example_path = "This is a path to an excel file or a folder containing excel files"
extractor = getConnectionsFromExcelFiles(path=example_path)
test, failed_files = extractor.get_connections_from_excel()

pd.DataFrame.from_dict(test, orient="index").transpose().to_csv("test_1-v2.csv")
pd.Series(data=failed_files, name="failed_files", dtype=object).to_csv("test_1_failed_files.csv")
```