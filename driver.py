import os
import os.path
import glob
import datetime
from typing import Dict, Literal, List
import xlwings as xw
import time
from pathlib import Path
import json

def main() -> None:
  """Entry way into the program."""
  run_time: int = 0
  i: int = 0
  args: str = read_args()
  data: Dict = json_read(args)
  #Setup only runs for the first time program is opened, look at source code in driver.py to see how.
  setup(data)
  #this reads the number of files in the rawData directory. Runs the program that many times.
  run_time: int = num_files(data)
  while(i < run_time):
      run(data, i)
      i = i + 1


def run(data: Dict, run_time: int) -> None:
    """Runs driver code"""
    file_type: str = newest_template(data)
    macro_template(file_type, data)
    macro_do(file_type, data)
    rename_spec(file_type, data, run_time)
    remove_old_file(data)


def json_read(args: str) -> Dict:
    """Reads config file"""
    if os.path.exists(args):
        f = open(args)
        data: Dict = json.load(f)
        return data
    else:
        raise Exception("Please enter the file location for config.json")


def read_args() -> str:
    """Provides file location of config.json"""
    config_loc: str = os.getcwd() + "\\config.json"
    config_loc: str ="C:\\Users\\FWMya\\Desktop\\Script WV4-Personal\\config.json"
    return(config_loc)


def num_files(data: Dict) -> int:
    """Returns the number of files in the rawData directory"""
    run_time = len([1 for x in list(os.scandir(os.getcwd() + data["rawDataSpotNoStar"]))])
    return run_time


def setup(data: Dict) -> None:
    """Setups the necessary directories for program. User must move templates into correct folders. Only runs the first time the program is opened on the computer."""
    os.chdir(os.getcwd())
    check_file = Path(os.getcwd() + data["rawDataSpotNoStar"])
    if check_file.is_dir() == False:
        os.mkdir(f'{data["rawDataFolder"]}')
        os.mkdir(f'{data["excelDump"]}')
        print("Necessary directories added. Please place templates in the templates file on your desktop and raw data into the rawdata folder on your desktop.")
        quit()

        
def macro_template(file_type, data: Dict) -> None:
    """Opens macro template"""
    if file_type == "verizon":
        os.chdir(os.getcwd() + data["templateSpot"])
        os.system(f'start excel.exe {data["verizonTemplateName"]}')
        #Necessary to allow time for macro to start
        time.sleep(data["macroStartTime"])
    elif file_type == "atandt":
        os.chdir(os.getcwd()+data["templateSpot"])
        os.system(f'start excel.exe {data["attTemplateName"]}')
        #Necessary to allow time for macro to start
        time.sleep(data["macroStartTime"])


def newest_template(data: Dict) -> str:
    """Finds newest excel sheet in the subdirectory and opens it"""
    file_type: str = ""
    list_of_files = glob.glob(os.getcwd()+data["rawDataSpot"])
    latest_file = sorted(list_of_files)[len(list_of_files)-1]
    raw_latest_file = '"{}"'.format(latest_file)
    os.system(f'start excel.exe {raw_latest_file}')
    #Will need to change this if raw data name conventions change
    if (data["verizonNameStart"]) in latest_file:
        file_type = "verizon"
    elif (data["attNameStart"]) in latest_file:
        file_type = "atandt"
    else:
        raise Exception("Error in original file name. Please ensure that file contains either HAP or fileExport depending on type and try again.")
    return file_type


def macro_do(file_type: str, data: Dict) -> None:
    """Makes macros run"""
    if file_type == "verizon":
        #Need this to allow time for finding the workbook
        time.sleep(data["macroPauseTime"])
        wb = xw.books.active
        macro = wb.macro(data["verizonMacroName"])
        macro()
        wb.close()
    elif file_type == "atandt":
         #Need this to allow time for finding the workbook
        time.sleep(data["macroPauseTime"])
        wb = xw.books.active
        macro = wb.macro(data["attMacroName"])
        macro()
        wb.close()


def rename_spec(file_type: str, data: Dict, run_time: int) -> str:
    """Renames file to correct date."""
    if file_type == "verizon":
        cor_time = datetime.datetime.now() - datetime.timedelta(days = run_time)
        name = data["FinalSpot"] + data["1stPartVerizonPythonRename"] + cor_time.strftime("%m%d%y") + f' {data["2ndPartVerizonPythonRename"]}'
        os.chdir(Path(os.getcwd()).parent)
        os.rename(os.getcwd() + "\\" + data["excelDumpSpot"] + data["verizonVBArename"], name)
        return name
    elif file_type == "atandt":
        #If passing more than 1 AT&T file, change this days value to match the amount of AT&T files.
        cor_time = datetime.datetime.now() - datetime.timedelta(days = 1)
        name = data["FinalSpot"] + data["1stPartATTPythonRename"]  + cor_time.strftime("%m%d%y") + f' {data["2ndPartATTPythonRename"]}'
        os.chdir(Path(os.getcwd()).parent)
        os.rename(os.getcwd() + "\\" + data["excelDumpSpot"] + data["attVBArename"], name)
        return name
    else:
        raise Exception("Error in save name.")


def remove_old_file(data: Dict) -> None:
    """Removes Raw Data file from the rawdata directory."""
    list_of_files = glob.glob(os.getcwd() + data["rawDataSpot"])
    latest_file = sorted(list_of_files)[len(list_of_files)-1]
    os.remove(latest_file)


def cur_day() -> str:
    """Finds the current day."""
    return_day: str = ""
    cur_time = datetime.datetime.now()
    return_day = cur_time.strftime("%w")
    return return_day


if __name__ == "__main__":
    main()