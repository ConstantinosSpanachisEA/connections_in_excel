from helper_functions_ea import Logger
from pathlib import Path
import shutil
import zipfile
import os


class getConnectionsFromExcelFiles(object):

    logger = Logger("Extract Connections").logger

    def __init__(self, path):
        self.path = Path(path)

    def create_zip_copy(self, excel_path: Path):
        shutil.copy(str(excel_path), str(excel_path).replace("xlsx", "zip"))

    def open_zip_file(self, zip_file_path:str):
        if zip_file_path.endswith("zip"):
            zip_ = zipfile.ZipFile(zip_file_path)
        else:
            raise TypeError(f"{zip_file_path} is not a zip file")
        return zip_

    def read_content(self, zip_folder: zipfile.ZipFile, zipfile_name:str):
        self.logger.info("Reading content of connections.xml")
        with zip_folder.open(zipfile_name) as f:
            connections = f.read().decode("UTF-8")
        f.close()
        return connections

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
                list_of_files = [i for i in self.path.glob("**/*")]
                excel_list = [i for i in list_of_files if i.suffix == ".xlsx"]
                del list_of_files
            else:
                raise NotADirectoryError("Please provide a valid directory")
        return excel_list

    def extract_command_text(self, wb, ):
        self.logger.info("Extracting the command text for each connection in excel workbook.")
        commands = []
        return commands

    def get_connections_from_excel(self):
        all_connections = {}
        failed_files = []
        excel_files = self.get_excel_files()
        connection_file = 'xl/connections.xml'
        number_of_excel_files = len(excel_files)
        self.logger.info(f"Found {number_of_excel_files} excel files")
        i = 0
        for excel in excel_files:
            i += 1
            progress = i / number_of_excel_files
            self.logger.info(f"File {i}: {str(excel)}. Progress {progress}")
            try:
                self.create_zip_copy(excel)
            except Exception as e:
                self.logger.error(f"Failed to create zip file for {str(excel)}. Error was {e}")
                failed_files.append(str(excel))
            else:
                zip_file_path = str(excel).replace("xlsx", "zip")
                try:
                    zip_folder = self.open_zip_file(zip_file_path)
                except Exception as e:
                    self.logger.error(f"Failed to open zip file {zip_file_path}. Error was {e}")
                    failed_files.append(str(excel))
                else:
                    try:
                        if connection_file in zip_folder.namelist():
                            connections = self.read_content(zip_folder=zip_folder, zipfile_name=connection_file)
                        else:
                            connections = []
                    except Exception as e:
                        self.logger.error(f"Could not read the commands in {str(excel)}. The error was {e}")
                        failed_files.append(str(excel))
                    else:
                        zip_folder.close()
                        if len(connections) > 0:
                            all_connections[str(excel)] = connections
                        self.delete_zip_file(zip_file_path)
                        del connections
        return all_connections, failed_files

    def delete_zip_file(self, zip_file_path: str) -> None:
        os.remove(zip_file_path)


if __name__ == '__main__':
    import pandas as pd

    test, failed_files = getConnectionsFromExcelFiles(r'C:\Users\c.spanachis\Downloads').get_connections_from_excel()
    pd.DataFrame.from_dict(test, orient="index").transpose().to_csv("test_1-v2.csv")
    pd.Series(data=failed_files, name="failed_files", dtype=object).to_csv("test_1_failed_files.csv")
