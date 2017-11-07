#############################################################################################################
#####  Script Name: xlsx_convert.py
#####  Author: Daniel Jacinto
#####  email: daniel.jacinto@sap.com
#####  Purpose : This script will be responsbile for the convertion of Microsoft XLSX (2010, etc)  to CSV
#####  Requirements : Required Library openpyxl
#####  {Installation : pip install openpyxl}
#####  https://openpyxl.readthedocs.io/en/default/
#####  Version: 1.1
##############################################################################################################

try:
	# Import the Standard Libraries - In theory, there should be no issues, but it could be that the python runtime is stripped down and the standard libraries
	# are not loaded and could cause issues in the subsequent steps
	import sys, os, time, shutil, csv, logging, logging.handlers, getopt, argparse
	
	try:
		# We Try to import the OpenpyXl library
		# Import our named libraries linke OpenPyXl
		import openpyxl as xlr
	except:
		print sys.exc_info()
		sys.exit(2)
		
except:
	print sys.exc_info()
	sys.exit(2)

def main():
	
	# We set our command line arguments
	parser = argparse.ArgumentParser(description='Python program to convert Microsoft Excel Files to CSV for processing')
	parser.add_argument('-v','--Version', action='version', version='%(prog)s 1.1')
	parser.add_argument('-l','--Logging Folder', dest='log_folder',help='Log file folder', required=True)
	parser.add_argument('-s','--Source File Folder', dest='src_folder', help='Source folder of XLSX files', required=True)
	parser.add_argument('-d','--Destination File folder', dest='dst_folder', help='Archive / Destination folder for original XLSX file', required=True)
	# This command line parameter will be enabled in a later revision of the script
	#parser.add_argument('-f','--File to Convert', dest='file_name', help='Convert single file to CSV')
	args = parser.parse_args()
	
	# Call the xls function 
	xls_convert( args.log_folder, args.src_folder, args.dst_folder )

# we define the xlsconvert function 
def xls_convert( log_folder, src_folder, dst_folder ):
	# We need the string variables, as when the os.listdir method runs, it then causes an issue with the shutil method call
	src = src_folder
	dst = dst_folder

	# Logging Variables and Parameters
	log_filename = log_folder +"xls_conv.log"
	logger = logging.getLogger("xls_convert")
	logger.setLevel(logging.DEBUG)
	# We set the rotating file to a size limit of 10 MB 
	handler = logging.handlers.RotatingFileHandler(filename=log_filename,maxBytes=10485760,backupCount=1)
	logger.addHandler(handler)
	log_format = "%(asctime)s %(levelname)s %(message)s"
	logging.basicConfig(format=log_format)
	# Set our csv file handling parameters 
	files = os.listdir(src_folder)
	files.sort()
	
	# First we check to see if the folder exists 
	if not os.path.exists(src) or not os.path.exists(dst):
		logger.error("*** " + time.strftime("%c"))
		logger.error("Unable to validate if folder exists:")
		if not os.path.exists(src):
			logger.error("Validate folder :" +src)
		if not os.path.exists(dst):
			logger.error("Validate folder :" +dst)
	else:
		
		print("Script Start Time       : " + time.ctime() )
		for f in files:
			src_file = src_folder +f
			
			if f != 'archive' and f.endswith(".xlsx") == True:
				
				xl_conv_file = src_file.replace(".xlsx","")
				
				with open( xl_conv_file+".csv", 'wb') as csvf:
					c = csv.writer(csvf)
					
					xlf = xlr.load_workbook(filename = src_file)
					xl_sheet_range = xlf.get_sheet_names()
					
					try:
						for sheet in xl_sheet_range:
							active_sheet = xlf.get_sheet_by_name(sheet)
							
							for row in active_sheet.rows:
								values = []
								
								for cell in row:
									value = cell.value
									
									if not isinstance(value, unicode):
										value=unicode(value) 
										value=value.encode('utf8') 
										values.append(value)
								# write the row to the csv file
								c.writerow( [cell.value for cell in row] )
			
						# Close the csv file
						csvf.close()
					
					except Exception, e:
						logger.error("*** " + time.strftime("%c"))			
						logger.error("There was an error in processing the file for conversion : " + src_file)
						logger.exception(e)
						
					# Log an entry in the log file
					logger.info("*** " + time.strftime("%c"))			
					logger.info("Succesfully converted : " + src_file)
					
					try:
						shutil.move( os.path.join(src,f), os.path.join(dst,f) )
					except Exception, e:
						print("There was an error archiving the file:", os.path.join(src,f))
						logger.error("*** " + time.strftime("%c"))			
						logger.error("There was an error archiving the file : " + src_file)
						logger.exception(e)
			else:
				# Log an entry in the log file
				logger.info("*** " + time.strftime("%c"))
				logger.info("Inspected File : " + src_file )
				logger.info("Nothing to do ?, as file doesnt meet the criteria to trigger the file convertion process - (Needs to be 'xlsx' File)")
			
if __name__ == "__main__":
	main()