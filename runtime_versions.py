#############################################################################################################
#####  Script Name: python_env_ver.py
#####  Author: Daniel Jacinto
#####  email: daniel.jacinto@sap.com
#####  Purpose : This script will return the libraries versions
#####  Version: 0.1
##############################################################################################################


import sys, os, time, shutil, csv, logging, openpyxl, setuptools

for name, module in sorted(sys.modules.items()): 
  if hasattr(module, '__version__'): 
    print name, module.__version__