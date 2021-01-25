# Extracting Connections from Excel Files

A repo that allows for the extraction of any connections within an excel spreadsheet, 
or a floder that contains excel spreadsheets.

In order to store results as a dataframe please use pd.DataFrame.from_dict(results, orient="index").
You can also store any files that failed to open using the return array.

## How to use

```python
    from search_connections import getConnectionsFromExcelFiles
    import pandas as pd

    example_path = "This is a path to an excel file or a folder containing excel files"
    extractor = getConnectionsFromExcelFiles(path=example_path)
    test, failed_files = extractor.get_connections_from_excel()

    pd.DataFrame.from_dict(test, orient="index").to_csv("test_1.csv")
    pd.Series(data=failed_files, name="failed_files").to_csv("test_1_failed_files.csv")

```