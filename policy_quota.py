import pymssql
import pymysql.cursors
import win32com.client as win32
import psutil
import os
import sys
import subprocess


class PolicyDB:
	# Connect(ip, port, username, password)
	
	#def __init__(self, host, user, password, database):
	def __init__(self):
		self.host = "" #pass your own host name
		self.user = "" #pass your own user name
		self.password = "" #pass your own password
		self.database='RequestCenter'
	def fetch_data(self):
		connection = pymssql.connect(host=self.host,
                                             user=self.user,
                                             password=self.password,
                                             database=self.database)

		try:
			with connection.cursor() as cursor:
				sql = r"select Name from AcAccount where primaryid = 643" # your own query you can pass here
				cursor.execute(sql)
				row = cursor.fetchall()
				print(row)
		finally:
			connection.close()

db=PolicyDB()
db.fetch_data()

# We can ignore below code if integration of result is done with Splunk.Otherwise below part can be used and enhanced accordingly
def data_receive():
                outlook = win32.Dispatch("outlook.application")
                outlookMailItem = 0x0
                outlookFormatPlain = 1
                outlookFormatRichText = 3
                outlookFormatHTML = 2
                outlookFormatUnspecified = 0

                newMail = outlook.CreateItem(outlookMailItem)
                newMail.Subject = "Policies and Governance DB Notification"
                newMail.BodyFormat = outlook.FormatRichText
                newMail.HTMLBody = "Hi This is Demonstration.!!"
                newMail.To = "ankush.a.aggarwal@accenture.com"
                newMail.Send()


result = openpyxl.load_workbook(filename = r"C:\Users\ankush.a.aggarwal\Documents\Quota_db_Results\results.xlsx") #pass your own file save location
finale_result = result.worksheets[0]
result.save(r"C:\Users\ankush.a.aggarwal\Documents\Quota_db_Results\results.xlsx") #pass your own file save location
        
# Lastchange
