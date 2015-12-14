
import imaplib
import email
import gspread
import json
import sys
import base64
import csv
from oauth2client.client import SignedJwtAssertionCredentials
from apscheduler.schedulers.blocking import BlockingScheduler



# this function is for parsing emails
# it connects to GMail to read License activation emails, and connects to
# Google Sheets and writes the parsed information to a spreadsheet
def parse():
    print "***************************************"
    print "Starting parser..."


    # first we are going to connect to Google Sheets
    print "Connecting to Google Sheets..."
    json_key = json.load(open('License-9fda17e1b871.json'))
    scope = ['https://spreadsheets.google.com/feeds']
    credentials = SignedJwtAssertionCredentials(json_key['client_email'], json_key['private_key'].encode(), scope)
    gsheet = gspread.authorize(credentials);
    licensesheet = gsheet.open("License Testing").sheet1 # this is the spreadsheet we will be writing to


    # now we are going to connect to the gmail account where the emails are being sent
    print "Connecting to GMail..."
    file_cred=open("cred.txt",'r')
    username=file_cred.readline()
    password=file_cred.readline()
    try:
        mail=imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(base64.b64decode(username),base64.b64decode(password))
    except:
        print "ERROR: Unable to connect to GMail."
        sys.exit()

    # now we are going to get access to the emails of interest (those labelled by 'License')
    try:
        mail.list() # Lists all labels in GMail
        mail.select('License') # Select emails that have the License label
        result, data = mail.uid('search', None, "UNSEEN") # select emails that have NOT been read
    except:
        print "ERROR: Unable to get emails labelled with 'License'"
        sys.exit()


    # now let's read the emails
   
    i = len(data[0].split()) # i will be equal to the number of unread emails

    if (i==0):
        print "No new emails to parse."
        print "Closing program"
        sys.exit()
        
    emailNumber = 0
    print "Reading emails..."
    for x in range(i):
        emailNumber=emailNumber+1
        print "\nReading email ",emailNumber," of",i
        
        latest_email_uid = data[0].split()[x] 
        result, email_data = mail.uid('fetch', latest_email_uid, '(RFC822)')
     
        # fetch the email body (RFC822) for the given ID
        raw_email = email_data[0][1]
         
        #continue inside the same for loop as above
        raw_email_string = raw_email.decode('utf-8')

        # find if registration was declined
        registration_declined=raw_email_string.find("REGISTRATION DECLINED", 0,len(raw_email_string))

        
     
        if (registration_declined!=-1):
            # this means that registration was declined, so we should skip parsing this email
            continue
        else:

            #find if 'REGITRATION INFO' is in the email (this string will not appear
            registration_info_present=raw_email_string.find("Registration info :", 0,len(raw_email_string))
            if (registration_info_present!=-1):
                # find computer name
                FIRST_DELIMITER = "Computer ["
                SECOND_DELIMITER = "]"
                location1=raw_email_string.find(FIRST_DELIMITER, 0,len(raw_email_string))
                location2=raw_email_string.find(SECOND_DELIMITER,location1,len(raw_email_string))
                computer_name=raw_email_string[location1+len(FIRST_DELIMITER):location2]

                
                #find email address of client
                FIRST_DELIMITER = "Email : ["
                SECOND_DELIMITER = "]"
                location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
                location2=raw_email_string.find(SECOND_DELIMITER,location1,len(raw_email_string))
                email=raw_email_string[location1+len(FIRST_DELIMITER):location2]

               
                #find the UI version
                FIRST_DELIMITER = "UI version : ["
                SECOND_DELIMITER = "]"
                location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
                location2=raw_email_string.find(SECOND_DELIMITER,location1,len(raw_email_string))
                ui_version=raw_email_string[location1+len(FIRST_DELIMITER):location2]
         
                
                #find the DB version
                FIRST_DELIMITER = "Database version : ["
                SECOND_DELIMITER = "]"
                location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
                location2=raw_email_string.find(SECOND_DELIMITER,location1,len(raw_email_string))
                db_version=raw_email_string[location1+len(FIRST_DELIMITER):location2]
         
                
                #find the license
                FIRST_DELIMITER = "License : ["
                SECOND_DELIMITED = "]"
                location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
                location2=raw_email_string.find(SECOND_DELIMITER,location1,len(raw_email_string))
                lic=raw_email_string[location1+len(FIRST_DELIMITER):location2]


                #find the license id
                FIRST_DELIMITER = "\" --license"
                location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
                if (location1==-1):
                    FIRST_DELIMITER = "--license"
                    location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
                    lic_id=raw_email_string[location1-32:location1]
                else:
                    lic_id=raw_email_string[location1-31:location1]

                
                
                #find the customer name
                FIRST_DELIMITER = "license\";"
                SECOND_DELIMITER = ";"
                location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
                location2=raw_email_string.find(SECOND_DELIMITER,location1+len(FIRST_DELIMITER),len(raw_email_string))
                customer_name=raw_email_string[location1+len(FIRST_DELIMITER):location2]

                
                #find the company name
                FIRST_DELIMITER = ";"
                SECOND_DELIMITER = ","
                location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
                location2=raw_email_string.find(SECOND_DELIMITER,location1,len(raw_email_string))
                company=raw_email_string[location1+len(FIRST_DELIMITER):location2]

                
                #find the street address
                FIRST_DELIMITER =","
                SECOND_DELIMITER = ","
                location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
                location2=raw_email_string.find(SECOND_DELIMITER,location1+len(FIRST_DELIMITER),len(raw_email_string))
                street=raw_email_string[location1+len(FIRST_DELIMITER):location2]

                
                #find the city
                location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
                location2=raw_email_string.find(SECOND_DELIMITER,location1+len(FIRST_DELIMITER),len(raw_email_string))
                city=raw_email_string[location1+len(FIRST_DELIMITER):location2]

                
                #find the prov
                location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
                location2=raw_email_string.find(SECOND_DELIMITER,location1+len(FIRST_DELIMITER),len(raw_email_string))
                prov=raw_email_string[location1+1:location2]


                #find the country
                location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
                location2=raw_email_string.find(SECOND_DELIMITER,location1+len(FIRST_DELIMITER),len(raw_email_string))
                country=raw_email_string[location1+1:location2]


                #find the postal code
                SECOND_DELIMITER = ";"
                location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
                location2=raw_email_string.find(SECOND_DELIMITER,location1,len(raw_email_string))
                postal=raw_email_string[location1+len(FIRST_DELIMITER):location2]

                 
                #find the date
                SECOND_DELIMITER = " "
                location1=location2
                location2=raw_email_string.find(SECOND_DELIMITER,location1,len(raw_email_string))
                date=raw_email_string[location1+len(FIRST_DELIMITER):location2]


                #find the time
                SECOND_DELIMITER = "\""
                location1=location2
                location2=raw_email_string.find(SECOND_DELIMITER,location1,len(raw_email_string))
                time=raw_email_string[location1+len(FIRST_DELIMITER):location2]


                # now that we have all the parsed data, we will write it to the spreadsheet
                array_of_registration_data=[computer_name,email,ui_version,db_version,lic,lic_id,customer_name,company,street,city,prov,country,postal,date,time]
                licensesheet.append_row(array_of_registration_data)

                
        
    #now we're going to export the google spreadsheet data to a file
    print "Exporting spreadsheet to file..."
    try:
        spreadsheet_data = licensesheet.export(format='csv')
        csv_file = open("licenseinfo.csv","w")
        csv_file.write(spreadsheet_data)
        csv_file.close()
    except:
        print "Unable to write to CSV file"
    print "Finished reading emails..."
    print "Closing parser..."
    mail.logout()


# this is the 'main' program, which contains a scheduled job that runs
# the parse() function every minute
print "Starting scheduler..."
sched=BlockingScheduler()
sched.add_job(parse,'interval',minutes=1)

try:
    sched.start()
except:
    print "Could not run script"
    sys.exit()
