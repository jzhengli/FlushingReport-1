##To generate a daily report for Sewer Flushing and send email to Sewer Team Manager
##Created by Joe Zheng Li, COR PU GIS, joe.li@raleighnc.gov
####################################################################################
import arcpy, os, sys, datetime, email, smtplib, logging
from arcpy import env
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email.Utils import formatdate
from email import Encoders
import xlwt, xlrd
from xlutils.filter import process, XLRDReader, XLWTWriter
from shutil import copy2


logging.basicConfig(filename=os.path.join(os.path.dirname(sys.argv[0]),'updates.log'),level=logging.INFO, format='%(asctime)s %(message)s')
# log message to keep track
def logMessage(msg):
    print time.strftime("%Y-%m-%d %H:%M:%S ", time.localtime()) + msg
    logging.warning(msg)
    return

##export report to excel
def ExportReport(table, delta_date):
	env.overwriteOutput = True
	env.workspace = "Database Connections/RPUD_TESTDB - MOBILE_EDIT_VERSION.sde"
	#env.workspace = os.path.join(os.path.dirname(sys.argv[0]), "RPUD_TESTDB - MOBILE_EDIT_VERSION.sde") #the name of database connection may need to be changed when in production

	#convert local time to UTC for query
	now = datetime.date.today().strftime('%Y%m%d')
	today = datetime.datetime(int(now[:4]), int(now[4:6]), int(now[6:]), 00, 00, 00) + datetime.timedelta(hours=4)
	yesterday = today - datetime.timedelta(days=delta_date)
	outputExcel = os.path.join("//corfile/Public_Utilities_NS/5215_Capital_Improvement_Projects/636_Geographic_Info_System/Joe/Collector App/Flushing app/Daily Report/", table + "_" + yesterday.strftime("%Y%m%d") + ".xls")
	logMessage("Input table is: " + table)
	logMessage("Output Excel file is: " + os.path.basename(outputExcel))
	print ("Exporting table to Excel...")

	#query report table for records in previous day
	whereClause = '"CREATED_DATE" < timestamp \'{0}\' AND "CREATED_DATE" > timestamp \'{1}\' AND "CREW" NOT LIKE \'_GIS\' AND "CREW" NOT LIKE \'_test\' ORDER BY REPORT_DATE'.format(str(today), str(yesterday))
	arcpy.MakeQueryTable_management(table, 'queryTable', "NO_KEY_FIELD", "", "", whereClause)
	recordNum = arcpy.GetCount_management('queryTable')
	logMessage(str(recordNum) + " " + table + " reports for " + (yesterday).strftime("%b %d, %Y"))

	#for test, print out fiels in queryTable
	# fields = arcpy.ListFields('queryTable')
	# for field in fields:
	# 	print("{0} is a type of {1}".format(field.aliasName, field.type))

	#export queried table to excel, ALIAS option does not work here so far, need a solution
	arcpy.TableToExcel_conversion('queryTable', outputExcel, 'ALIAS')
	logMessage(os.path.basename(outputExcel) + " has been exported.")
	#return yesterday date for naming
	return yesterday, recordNum
	

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
			print ("message attached")
	server = smtplib.SMTP(HOST)
	print ("Connected to server")
	server.sendmail(FROM, TO.split(",") + CC.split(","), msg.as_string())
	print ("Sending Email...")
	server.close()
	for fPath in fPaths:
		os.remove(fPath)	
	print ("Email sent")

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
	print ("Combining reports...")
	combinedReport = os.path.dirname(filelist[0]) + "/FlushingReport_" + yesterday.strftime("%Y%m%d") + ".xls"
	wkbk.save(combinedReport)
	logMessage("Report Merged: " + combinedReport)
	return combinedReport


##main

logMessage("******************Export Flushing Report******************")
outputDir = "//corfile/Public_Utilities_NS/5215_Capital_Improvement_Projects/636_Geographic_Info_System/Joe/Collector App/Flushing app/Daily Report/"
deltaDate = 1 #can be changed if need multiple days report
yesterday, mainNum = ExportReport("RPUD.SewerMainFlushing", deltaDate)
yesterday_1, mhNum = ExportReport("RPUD.SewerMHFlushing", deltaDate)
print "Main: {0}, MH: {1}".format(mainNum, mhNum)
reportList = ["{0}/RPUD.SewerMainFlushing_{1}.xls".format(outputDir, yesterday.strftime("%Y%m%d")), "{0}/RPUD.SewerMHFlushing_{1}.xls".format(outputDir, yesterday.strftime("%Y%m%d"))]
filepaths = [CombineReport(reportList)]
copy2(filepaths[0], "{0}/FR_copy_{1}.xls".format(os.path.dirname(filepaths[0]), yesterday.strftime("%Y%m%d")))
filepaths_copy = ["{0}/FR_copy_{1}.xls".format(os.path.dirname(filepaths[0]), yesterday.strftime("%Y%m%d"))]

##send out email with attached excel report for previous day
emailSub = "Daily Flushing Report"
#message = "Hi, \n\n"
message = "Attached please find Flushing Report for " + yesterday.strftime("%b %d, %Y")
message += "\n\nGravity Main Flushing reports: {1}\nManhole Flushing reports: {2}".format(yesterday.strftime("%b %d, %Y"), mainNum, mhNum)
message += "\n\nThanks,\n"
message += "PUGIS"

#email to send in test
#to = "joe.li@raleighnc.gov"
# cc = ""
#email to send in production
to = "jeffrey.bognar@raleighnc.gov"
cc = "david.jackson@raleighnc.gov, chris.mosley@raleighnc.gov, dustin.tripp@raleighnc.gov" #tom.johnson@raleighnc.gov" ## email copy list

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
logMessage("Email sent to sewer")
#Notify GIS team member

SendEmail(filepaths_copy, isAttach, message, "zheng.li@raleighnc.gov", "", "Flushing Report Sent")
logMessage("Email notification to GIS team")
# else:
# 	print "Fail to verify"




