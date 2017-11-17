# SAP-XLSX-PyConvert

This project was bourne of the need to convert XLSX Files to CSV for processing in a Leads Management Solution based on a SAP Netweaver BW System. This solution is meant to run in batch, meaning that there is a directory with XLSX files and this script is meant to iterate through the directory and convert all of the files in the directory in a back ground fashion.

There is various methods to read XLSX files into ABAP, but due to the various version restrictions and  limitations on the SAP Netweaver platform, this solution was choosen, in which to use python to read and convert XLSX files.

# Known Issues / Limitations
The current script will only do a basic conversion of the Microsoft Excel file, meaning that it will simply read and convert. I fyou have any objects such as pivot tables, charts, etc. You need to amend the script accodingly to apply your own logic.

# Usage of the script
The script can be executed with the following command line parameters.

*Optional arguments:*

##### -h, Show this help message and exit
  
##### -v, Show program's version number and exit
  
##### -f, Convert single file to CSV
  
*Mandatory arguments:*

##### -l, This is the folder path for the rotating log file

##### -s, Source folder of XLSX files

##### -d, Archive / Destination folder for original XLSX file




# Example Usage of Script

**This command will display all of the help associated the program**

python xlsx_convert.py -h


**This command will display the version of the script**

python xlsx_convert.py -v


**This usage will execute the script and convert all files in the -s '/SAP-XLSX-PyConvert/Source/' folder, write all log messages to file 'xls_conv.log' in the -l '/SAP-XLSX-PyConvert/Log_Folder/' folder and move the orginal XLSX files to the -d '/SAP-XLSX-PyConvert/Archive/' folder.**

python xlsx_convert.py -l '/SAP-XLSX-PyConvert/Log_Folder/' -s '/SAP-XLSX-PyConvert/Source/' -d '/SAP-XLSX-PyConvert/Archive/'


**Please Note: This will be enabled as part of version 1.2 of the script**

This usage will execute the script and convert a single file using the -f 'VafBM_20161021.xlsx' in the -s '/SAP-XLSX-PyConvert/Source/' folder, write all log messages to file 'xls_conv.log' in the -l '/SAP-XLSX-PyConvert/Log_Folder/' folder and move the orginal XLSX files to the -d '/SAP-XLSX-PyConvert/Archive/' folder.

python xlsx_convert.py -l '/SAP-XLSX-PyConvert/Log_Folder/' -s '/SAP-XLSX-PyConvert/Source/' -d '/SAP-XLSX-PyConvert/Archive/' -f 'VafBM_20161021.xlsx'

# Runtime Environment 
This script was developed and tested with the following runtime environment(s):

**Local development environment**
* Python version -  2.7.10
* Mac OS Sierra - 10.12.6

**SAP Application Server Environment**

SUSE Enterprise Linux 

* Linux version 3.0.101-107-default
* gcc version 4.3.4  

**SAP Application Environment Infomation**

* SAP Netweaver 7.40 SP 13
* Python version -  2.6.9

[For more detailed runtime information](https://github.com/jacintod/SAP-XLSX-PyConvert/wiki/Requirements-and-Versioning) 


# ABAP Usage  
There is several ways in which to call the python script, the easiest way is to wrap it in a function module / class and then call it in your report program.

For the purposes of this project, it has been wrapped in a function module and in the main ABAP Program, the function module is called which passes the parameters to the python script.

```ABAP
*DATA IV_SRC_FOLDER TYPE CHAR0064.
*DATA IV_LOG_FOLDER TYPE CHAR0064.
*DATA IV_DST_FOLDER TYPE CHAR0064.

CALL FUNCTION 'Z_XLS_CONV'
  EXPORTING
    iv_src_folder       = '/usr/sap/SAPBW/intf/in/'
    iv_log_folder       = '/usr/sap/SAPBW/intf/in/'
    iv_dst_folder       = '/usr/sap/SAPBW/intf/out/'.
```