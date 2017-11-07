# SAP-XLSX-PyConvert

This project was bourne of the need to convert XLSX Files to CSV for processing in a Leads Management Solution based on a SAP Netweaver BW System. This solution is meant to run in batch, meaning that there is a directory with XLSX files and this script is meant to iterate through the directory and convert all of the files in the directory in a back ground fashion.

There is various methods to read XLSX files into ABAP, but due to the various version restrictions and  limitations on the SAP Netweaver platform, this solution was choosen, in which to use python to read and convert XLSX files.

# Known Issues / Limitations


# Usage of the script
The script can be executed with the following command line parameters.

Optional arguments:
  -h, Show this help message and exit
  
  -v, Show program's version number and exit
  
  -f, Convert single file to CSV
  
Mandatory arguments:
  -l, This is the folder path for the rotating log file

  -s, Source folder of XLSX files

  -d, Archive / Destination folder for original XLSX file

Example Usage of Script

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

The script was developed using the following libraries versions:
* BaseHTTPServer 0.3
* SimpleHTTPServer 0.6
* SocketServer 0.4
* _csv 1.0
* _ctypes 1.1.0
* _elementtree 1.0.6
* _struct 0.2
* cgi 2.6
* csv 1.0
* ctypes 1.1.0
* decimal 1.70
* distutils 2.7.10
* email 4.0.3
* et_xmlfile 1.0.1
* jdcal 1.2
* json 2.0.9
* logging 0.5.1.2
* openpyxl 2.3.4
* parser 0.5
* pickle $Revision: 72223 $
* pkg_resources._vendor.packaging 15.3
* pkg_resources._vendor.packaging.__about__ 15.3
* platform 1.0.7
* pyexpat 2.7.10
* re 2.2.1
* setuptools 18.5
* setuptools.version 18.5
* urllib 1.17
* urllib2 2.7
* xml.parsers.expat $Revision: 17640 $
* zlib 1.0

# ABAP Information 
The following application components were in use for the development and deployment of the script
SAP Application Environment Information
* SAP_BASIS	740	0013 - SAP Basis Component
* SAP_ABA	740	0013 - Cross-Application Component
* SAP_GWFND	740	0013 - SAP Gateway Foundation 7.40
* SAP_UI	740	0015 - User Interface Technology 7.40
* PI_BASIS	740	0013 - Basis Plug-In
* ST-PI	740	0006 - SAP Solution Tools Plug-In
* BI_CONT	747	0015 - Business Intelligence Content
* SAP_BW	740	0013 - SAP Business Warehouse
* GRCPINW	V1100_731 - SAP GRC NetWeaver Plug-In
* DMIS	2011_1_731 - DMIS 2011_1
* SAS	930	0000	-	SAS/ACCESS Interface to R/3
* ST-A/PI	01S_731	0002 - Servicetools for SAP Basis 731

# ABAP Usage  
There is several ways in which to call the python script, the easiest way is to wrap it in a function module / class and then call it in your report program.
For the purposes of this project, it has been wrapped in a function module and in the main ABAP Program, the function module is called which passes the parameters to the python script.

*DATA IV_SRC_FOLDER TYPE CHAR0064.
*DATA IV_LOG_FOLDER TYPE CHAR0064.
*DATA IV_DST_FOLDER TYPE CHAR0064.

CALL FUNCTION 'Z_XLS_CONV'
  EXPORTING
    iv_src_folder       = '/usr/sap/SAPBW/intf/in/'
    iv_log_folder       = '/usr/sap/SAPBW/intf/in/'
    iv_dst_folder       = '/usr/sap/SAPBW/intf/out/'.