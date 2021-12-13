# this code sends an email from UnileverIran SMTP server to a list of recipients
from smtplib import SMTP
from mimetypes import guess_type
from os.path import basename
from email.message import EmailMessage as emsg
import csv
###=================Sender========================###
user = 'example@foo.com'                             # sender email
passw = '[password]'                                        # sender password
smtpServ = '[IP address]'                                    # smtp ip address

smtpPort = 587                                                 # smtp port

###==============Server setting==================###

server = SMTP(smtpServ, smtpPort)
server.login("[usename]", "[password]")

###================ CSV ========================###

csvpath = 'Emails.csv'
with open(csvpath, newline='') as csvf:
    sheet = csv.reader(csvf)
    sheet = list(sheet)
NumRec = range(1, len(sheet))

###=================Message=====================###

for ii in NumRec:                                               # Loop over all the recipients
    msg            = emsg()                                     # Create a Message Object
    msg['From']    = user                                     # Single Item
    msg['To']      = sheet[ii][1]                              # The recipient address as list
    msg['Subject'] = "Hello {} !".format(sheet[ii][3])          # Message's Title

    msgBody        = "Dear {},\r\n\n\
This message is sent to you to test my code. \
your token is : {} and please make reservation as soon as you could. \r\n\n\
Department {} \r\n\n\
Regards,".format(sheet[ii][3], sheet[ii][4], sheet[ii][5])      # Message's Body
    msg.set_content(msgBody)                                        # Add Body to msg object

###==============Attachement=====================###
    fpath = '[example.pdf]'
    fname = basename(fpath)                                         # get the basename of the file
    atch = open(fpath, 'r+b')                                       # open the file for reading in binary
    ctype, encoding = guess_type(fpath)                             # guess main and sub type of the file

    msg.add_attachment(atch.read(), \
                       maintype=ctype[0], subtype=ctype[1], filename=fname) # Add attachment to msg object

    atch.close()                                                    # Close the file

###================Send message===================###

    server.send_message(msg)                                         # send_message instead of sendmail
    print('Message #{} Was Sent Successfully...'.format(ii))
server.close()
