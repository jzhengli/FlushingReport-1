##To generate a daily report from database for Sewer Flushing and send email to Sewer Team Manager
##Created by Joe Zheng Li, PU GIS, joe.li@raleighnc.gov
##
import arcpy, os, sys, datetime, email, smtplib
from arcpy import env
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email.Utils import formatdate
from email import Encoders
import xlwt, xlrd
from xlutils.filter import process, XLRDReader, XLWTWriter
from shutil import copy2


##export report of previous day to excel
def ExportReport(table, delta_date):
	env.overwriteOutput = True
	env.workspace = "Database Connections/RPUD_TESTDB - MOBILE_EDIT_VERSION.sde"
	#env.workspace = os.path.join(os.path.dirname(sys.argv[0]), "RPUD_TESTDB - MOBILE_EDIT_VERSION.sde") #the name of database connection may need to be changed when in production

	# if table == "RPUD.SewerMainFlushing":
	# 	reportName = "Gravity Main Flushing Report_"
	# elif table == "RPUD.SewerMHFlushing":
	# 	reportName = "Manhole Flushing Report_"
	today = datetime.date.today() + datetime.timedelta(hours=4)
	yesterday = today - datetime.timedelta(days=delta_date)
	outputExcel = os.path.join("//corfile/Public_Utilities_NS/5215_Capital_Improvement_Projects/636_Geographic_Info_System/Joe/Collector App/Flushing app/Daily Report/", table + "_" + yesterday.strftime("%Y%m%d") + ".xls")
	print "Input table is: " + table
	print "Output Excel file is: " + os.path.basename(outputExcel)
	print "Exporting table to Excel..."

	#query report table for records in previous day
	whereClause = '"CREATED_DATE" < date \'{0}\' AND "CREATED_DATE" > date \'{1}\' AND "CREW" NOT LIKE \'_GIS\' AND "CREW" NOT LIKE \'_test\' ORDER BY REPORT_DATE'.format(str(today), str(yesterday))
	arcpy.MakeQueryTable_management(table, 'queryTable', "NO_KEY_FIELD", "", "", whereClause) 
	print str(arcpy.GetCount_management('queryTable')) + " " + table + " reports for " + (yesterday).strftime("%b %d, %Y")

	#for test, print out fiels in queryTable
	# fields = arcpy.ListFields('queryTable')
	# for field in fields:
	# 	print("{0} is a type of {1}".format(field.aliasName, field.type))

	#export queried table to excel, ALIAS option does not work here so far, need a solution
	arcpy.TableToExcel_conversion('queryTable', outputExcel, 'ALIAS')
	print os.path.basename(outputExcel) + " has been exported."
	#return yesterday date for naming
	return yesterday
	

#send email
def SendEmail(fPaths, isAttachmt, body, toList, ccList, subject):
	
	HOST = "cormailgw.raleighnc.gov"
	#FROM = "joe.li@raleighnc.gov"
	FROM = "PubUtilGIS@raleighnc.gov"
	TO = toList
	CC = ccList
	msg = MIMEMultipart()
	msg['FROM'] = FROM 
	msg['TO'] = TO
	msg['CC'] = CC
	msg['Date'] = formatdate(localtime = True)
	msg['Subject'] = subject
	msg.attach(MIMEText(body))
	if isAttachmt:
		for fPath in fPaths:
			part = MIMEBase('text/plain', 'octet-stream')
			part.set_payload(open(fPath, 'rb').read())
			Encoders.encode_base64(part)
			part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(fPath))
			msg.attach(part)
			print "message attached"
	server = smtplib.SMTP(HOST)
	print "Connected to server"
	server.sendmail(FROM, TO.split(",") + CC.split(","), msg.as_string())
	print "Sending Email..."
	server.close()
	for fPath in fPaths:
		os.remove(fPath)	
	print "Email sent"

#combined main and manhole into one spreadsheet
#
#copy spread sheet with cell style
def copy(wb):
	w = XLWTWriter()
	process(
		XLRDReader(wb, 'unknown.xls'),
		w
		)
	return w.output[0][1], w.style_list
#
#copy and paste two report sheets into one workbook
def CombineReport(filelist):
	wkbk = xlwt.Workbook()
	outrow_idx = 0
	file_idx = 0
	#iterate through workbook to copy and past sheets
	for f in filelist:
		outrow_idx = 0
		inBook = xlrd.open_workbook(f, formatting_info=True, on_demand=True)
		insheet = inBook.sheets()[0]
		if file_idx == 0:
				wkbk, outStyle = copy(inBook)
				wkbk.get_sheet(0).name = "GravityMainFlushing"
				outsheet1 = wkbk.get_sheet(0)
				outsheet2 = wkbk.add_sheet("ManholeFlushing")
		else:
			wkbk_mh, outStyle = copy(inBook)
		for row_idx in xrange(insheet.nrows):
			for col_idx in xrange(insheet.ncols):
				xf_idx = insheet.cell_xf_index(row_idx,col_idx)
				#print "xf_idx of row {0}, col {1} is {2}".format(row_idx, col_idx, xf_idx) 
				saved_style = outStyle[xf_idx]
				if file_idx == 0:
					outsheet1.write(outrow_idx, col_idx, insheet.cell_value(row_idx, col_idx), saved_style)
				else:					
					outsheet2.write(outrow_idx, col_idx, insheet.cell_value(row_idx, col_idx), saved_style)
			outrow_idx += 1
		file_idx += 1
	combinedReport = os.path.dirname(filelist[0]) + "/FlushingReport_" + yesterday.strftime("%Y%m%d") + ".xls"
	wkbk.save(combinedReport)
	return combinedReport

# pw = "2369" #raw_input("password to run script:")
# print pw
# if pw == "2369":
# print "Verified..."
outputDir = "//corfile/Public_Utilities_NS/5215_Capital_Improvement_Projects/636_Geographic_Info_System/Joe/Collector App/Flushing app/Daily Report/"
deltaDate = 1 #can be changed if need multiple days report
yesterday = ExportReport("RPUD.SewerMainFlushing", deltaDate)
ExportReport("RPUD.SewerMHFlushing", deltaDate)
reportList = ["{0}/RPUD.SewerMainFlushing_{1}.xls".format(outputDir, yesterday.strftime("%Y%m%d")), "{0}/RPUD.SewerMHFlushing_{1}.xls".format(outputDir, yesterday.strftime("%Y%m%d"))]
filepaths = [CombineReport(reportList)]
copy2(filepaths[0], "{0}/FR_copy_{1}.xls".format(os.path.dirname(filepaths[0]), yesterday.strftime("%Y%m%d")))
filepaths_copy = ["{0}/FR_copy_{1}.xls".format(os.path.dirname(filepaths[0]), yesterday.strftime("%Y%m%d"))]
print filepaths_copy
##send out email with attached excel report for previous day
emailSub = "Daily Flushing Report"
#message = "Hi, \n\n"
message = "Attached please find Flushing Report for " + yesterday.strftime("%b %d, %Y")
message += "\n\nThanks,\n"
message += "PUGIS"
#message += "Joe\n\n"
#message += "Zheng (Joe) Li\nGIS Programmer/Analyst\nCity of Raleigh Public Utilities\n3304 Terminal Dr Bldg 300\nRaleigh, NC 27604\n919.996.2369"

#email to send in test
#to = "joe.li@raleighnc.gov"
# cc = ""
#email to send in production
to = "jeffrey.bognar@raleighnc.gov"
cc = "david.jackson@raleighnc.gov, chris.mosley@raleighnc.gov, tom.johnson@raleighnc.gov" ## email copy list

isAttach = True
print 'Formatting email...'
print '-' * 80
print 'Subject: %s' % emailSub
print 'Message: %s' % message
if isAttach:
	print 'Attachment: %s' % filepaths
print '-' * 80


#send daily report to Sewer Team Manager
SendEmail(filepaths, isAttach, message, to, cc, emailSub)
#Notify GIS team member
SendEmail(filepaths_copy, isAttach, "Flushing Report has been sent to Sewer team.", "zheng.li@raleighnc.gov", "", "Flushing Report Sent")

# else:
# 	print "Fail to verify"




