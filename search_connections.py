import win32com.client
from pywintypes import com_error
import time
import logging
from helper_functions_ea import Logger
from pathlib import Path

_DELAY = 0.06  # seconds
_TIMEOUT = 70.0  # seconds


def _com_call_wrapper(f, *args, **kwargs):
    """
    COMWrapper support function.
    Repeats calls when 'Call was rejected by callee.' exception occurs.
    """
    # Unwrap inputs
    args = [arg._wrapped_object if isinstance(arg, ComWrapper) else arg for arg in args]
    kwargs = dict([(key, value._wrapped_object)
                   if isinstance(value, ComWrapper)
                   else (key, value)
                   for key, value in dict(kwargs).items()])

    start_time = None
    while True:
        try:
            result = f(*args, **kwargs)
        except com_error as e:
            if e.strerror == 'Call was rejected by callee.':
                if start_time is None:
                    start_time = time.time()
                    logging.warning('Call was rejected by callee.')

                elif time.time() - start_time >= _TIMEOUT:
                    raise

                time.sleep(_DELAY)
                continue

            raise

        break

    if isinstance(result, win32com.client.CDispatch) or callable(result):
        return ComWrapper(result)
    return result


class ComWrapper(object):
    """
    Class to wrap COM objects to repeat calls when 'Call was rejected by callee.' exception occurs.
    """

    def __init__(self, wrapped_object):
        assert isinstance(wrapped_object, win32com.client.CDispatch) or callable(wrapped_object)
        self.__dict__['_wrapped_object'] = wrapped_object

    def __getattr__(self, item):
        return _com_call_wrapper(self._wrapped_object.__getattr__, item)

    def __getitem__(self, item):
        return _com_call_wrapper(self._wrapped_object.__getitem__, item)

    def __setattr__(self, key, value):
        _com_call_wrapper(self._wrapped_object.__setattr__, key, value)

    def __setitem__(self, key, value):
        _com_call_wrapper(self._wrapped_object.__setitem__, key, value)

    def __call__(self, *args, **kwargs):
        return _com_call_wrapper(self._wrapped_object.__call__, *args, **kwargs)

    def __repr__(self):
        return 'ComWrapper<{}>'.format(repr(self._wrapped_object))


class getConnectionsFromExcelFiles(object):

    logger = Logger("Extract Connections").logger
    xl = None

    def __init__(self, path):
        self.path = Path(path)
        self.create_excel_wrapper()

    def create_excel_wrapper(self):
        _xl = win32com.client.Dispatch('Excel.Application')
        self.xl = ComWrapper(_xl)

    def get_excel_files(self):
        self.logger.info("Finding all excel files in the provided path")
        excel_list = []
        if self.path.is_file():
            if self.path.suffix == ".xlsx":
                excel_list.append(self.path)
            else:
                raise TypeError(f"The file {str(self.path)} is not an xlsx file. ")
        else:
            if self.path.is_dir():
                excel_list = [i for i in self.path.glob("*.xlsx")]
            else:
                raise NotADirectoryError("Please provide a valid directory")
        return excel_list

    def extract_command_text(self, wb, ):
        self.logger.info("Extracting the command text for each connection in excel workbook.")
        commands = []
        for x in wb.Connections:
            commands.append(x.OLEDBConnection.CommandText)
        return commands

    def get_connections_from_excel(self):
        all_connections = {}
        failed_files = []
        excel_files = self.get_excel_files()
        for excel in excel_files:
            try:
                wb = self.xl.workbooks.open(str(excel))
            except Exception as e:
                self.logger.error(f"Failed to open file {str(excel)}")
                failed_files.append(str(excel))
            else:
                try:
                    commands = self.extract_command_text(wb)
                    all_connections[str(excel)] = commands
                except Exception as e:
                    raise Exception(f"Could not read the commands in {str(excel)}")
                else:
                    wb.Close()
        return all_connections, failed_files

    def close_excel_wrapper(self):
        self.xl.Application.quit()


if __name__ == '__main__':
    import pandas as pd
    test, failed_files = getConnectionsFromExcelFiles(r'Z:\1. Research - Oil products\3. Spreadsheets').get_connections_from_excel()
    pd.DataFrame.from_dict(test, orient="index").transpose().to_csv("refinery_spreadsheets.csv")
    pd.Series(data=failed_files, name="failed_files").to_csv("refinery_spreadsheets_failed_files.csv")