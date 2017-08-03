###	Program: Google Account Generator
###	Programmer: Josue Sanchez
###	Company: Computerwisekids Inc.
###	Date: 07/17/2017
###	Purpose:To generate emails and passwords and save them in
###			a file so they may be uploaded to Google Admin.

import csv
import pyexcel as pe
import os

for file in os.listdir("files"):
	print(file)
	
	#open file
	iFileReader = pe.get_sheet(start_row = 1, file_name= str(os.path.join("files", file))) 
				
	# Lists that will hold student info
	emails = []		
	firstNames = []
	lastNames = []
	passwords = []

	#Ask user for school and year
	school = input("Enter school abbreviation: ")
	year = input("Enter Culminating Year: ")
	year = str(year)

	#Retrive, generate, and store data for each student
	for row in iFileReader:
		tempFirst = row[1]
		tempLast = row[3]
		tempEmail = year[2:] + school + str(tempFirst) + "@compclass.org" #Generate Email
		tempPassword = row[7]
		password = tempPassword + year	#Generate Password
		i = 1
		while tempEmail in emails:	#Change email if it already exist 
			tempEmail = year[2:] + school + str(tempFirst) + str(i) + "@compclass.org"
			i +=1
		#Add data to respective list
		firstNames.append(tempFirst)
		lastNames.append(tempLast)
		emails.append(tempEmail)
		passwords.append(password)

	#Generate filename
	oFileName = school + year +".csv"

	#Write data to csv file
	with open(oFileName, "w") as oFile:
		oFileWriter = csv.writer(oFile, lineterminator = "\n") #Set file variable and line terminator
		oFileWriter.writerow(["First Name","Last Name","Email","Password"])	#Write headers
		idx = 0
		#Write data for each student
		for email in emails:
			oFileWriter.writerow([firstNames[idx], lastNames[idx], emails[idx], passwords[idx]])
			idx+=1

	#Close the file
	oFile.close()
	print ()

print("Accounts Created")