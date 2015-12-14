import imaplib
import email
import gpread-master/gspread



mail=imaplib.IMAP4_SSL('imap.gmail.com')

mail.login('cgriffiths@cognisens.com','Cathryn88192cat')
mail.list() # Lists all labels in GMail

mail.select('License') # Select emails that have the License label

result, data = mail.uid('search', None, "UNSEEN") # select emails that have NOT been read

i = len(data[0].split())


for x in range(i):
    latest_email_uid = data[0].split()[x] # unique ids wrt label selected
    result, email_data = mail.uid('fetch', latest_email_uid, '(RFC822)')
 
    # fetch the email body (RFC822) for the given ID
    raw_email = email_data[0][1]
     
    #continue inside the same for loop as above
    raw_email_string = raw_email.decode('utf-8')

    # find if registration was declined
    registration_declined=raw_email_string.find("REGISTRATION DECLINED", 0,len(raw_email_string))
 
    if registration_declined!=-1:
        continue;
    else:
        # find computer name
        FIRST_DELIMITER = "Computer ["
        SECOND_DELIMITER = "]"
        location1=raw_email_string.find(FIRST_DELIMITER, 0,len(raw_email_string))
        location2=raw_email_string.find(SECOND_DELIMITER,location1,len(raw_email_string))
        computer_name=raw_email_string[location1+len(FIRST_DELIMITER):location2]
        print computer_name
        
        #find email address of client
        FIRST_DELIMITER = "Email : ["
        SECOND_DELIMITER = "]"
        location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
        location2=raw_email_string.find(SECOND_DELIMITER,location1,len(raw_email_string))
        email=raw_email_string[location1+len(FIRST_DELIMITER):location2]
        print email;
       
        #find the UI version
        FIRST_DELIMITER = "UI version : ["
        SECOND_DELIMITER = "]"
        location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
        location2=raw_email_string.find(SECOND_DELIMITER,location1,len(raw_email_string))
        ui_version=raw_email_string[location1+len(FIRST_DELIMITER):location2]
        print ui_version;
        
        #find the DB version
        FIRST_DELIMITER = "Database version : ["
        SECOND_DELIMITER = "]"
        location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
        location2=raw_email_string.find(SECOND_DELIMITER,location1,len(raw_email_string))
        db_version=raw_email_string[location1+len(FIRST_DELIMITER):location2]
        print db_version;
        
        #find the license
        FIRST_DELIMITER = "License : ["
        SECOND_DELIMITED = "]"
        location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
        location2=raw_email_string.find(SECOND_DELIMITER,location1,len(raw_email_string))
        lic=raw_email_string[location1+len(FIRST_DELIMITER):location2]
        print lic;

        #find the license id
        FIRST_DELIMITER = "Registration info :"
        location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
        if db_version is "2.5.5b":
            SECOND_DELIMITER = "\""
            location1=raw_email_string.find(SECOND_DELIMITER,location1,len(raw_email_string))
            location2=raw_email_string.find(SECOND_DELIMITER,location1+1,len(raw_email_string))
            lic_id=raw_email_string[location1+len(SECOND_DELIMITER):location2]
        else:
            SECOND_DELIMITER = " "
            location2=raw_email_string.find(SECOND_DELIMITER,location1+len(FIRST_DELIMITER),len(raw_email_string))
            lic_id=raw_email_string[location1+len(FIRST_DELIMITER):location2]
        print lic_id;
        
        #find the customer name
        FIRST_DELIMITER = "license\";"
        SECOND_DELIMITER = ";"
        location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
        location2=raw_email_string.find(SECOND_DELIMITER,location1+len(FIRST_DELIMITER),len(raw_email_string))
        customer_name=raw_email_string[location1+len(FIRST_DELIMITER):location2]
        print customer_name;
        
        #find the company name
        FIRST_DELIMITER = ";"
        SECOND_DELIMITER = ","
        location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
        location2=raw_email_string.find(SECOND_DELIMITER,location1,len(raw_email_string))
        company=raw_email_string[location1+len(FIRST_DELIMITER):location2]
        print company;
        
        #find the street address
        FIRST_DELIMITER =","
        SECOND_DELIMITER = ","
        location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
        location2=raw_email_string.find(SECOND_DELIMITER,location1+len(FIRST_DELIMITER),len(raw_email_string))
        street=raw_email_string[location1+len(FIRST_DELIMITER):location2]
        print street;
        
        #find the city
        location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
        location2=raw_email_string.find(SECOND_DELIMITER,location1+len(FIRST_DELIMITER),len(raw_email_string))
        city=raw_email_string[location1+len(FIRST_DELIMITER):location2]
        print city;
        
        #find the prov
        location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
        location2=raw_email_string.find(SECOND_DELIMITER,location1+len(FIRST_DELIMITER),len(raw_email_string))
        prov=raw_email_string[location1+1:location2]
        print prov;

        #find the country
        location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
        location2=raw_email_string.find(SECOND_DELIMITER,location1+len(FIRST_DELIMITER),len(raw_email_string))
        country=raw_email_string[location1+1:location2]
        print country;

        #find the postal code
        SECOND_DELIMITER = ";"
        location1=raw_email_string.find(FIRST_DELIMITER, location2,len(raw_email_string))
        location2=raw_email_string.find(SECOND_DELIMITER,location1,len(raw_email_string))
        postal=raw_email_string[location1+len(FIRST_DELIMITER):location2]
        print postal;
         
        #find the date
        SECOND_DELIMITER = " "
        location1=location2
        location2=raw_email_string.find(SECOND_DELIMITER,location1,len(raw_email_string))
        date=raw_email_string[location1+len(FIRST_DELIMITER):location2]
        print date;

        #find the time
        SECOND_DELIMITER = "\""
        location1=location2
        location2=raw_email_string.find(SECOND_DELIMITER,location1,len(raw_email_string))
        time=raw_email_string[location1+len(FIRST_DELIMITER):location2]
        print time;

        array_of_registration_data=[computer_name,email,ui_version,db_version,lic,lic_id,customer_name,company,street,city,prov,country,postal,date,time]
    
    

     

 


 
mail.logout()

