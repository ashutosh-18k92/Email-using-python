# -*- coding: utf-8 -*-
"""
Created on Sun Dec  8 11:47:20 2019

@author: ashutosh
"""
import smtplib,ssl
import csv
from socket import gaierror
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

class EmailReporter:
    _context = ssl.create_default_context()
    def __init__(self,login,password,smtp_server,port=587):
        self._login = login
        self._password=password
        self._smtp_server = smtp_server
        self._port = port
        self._sender = login
    
    def _connect_mail_server(self):
        try:
            server = smtplib.SMTP(self._smtp_server, self._port)
            server.starttls(context=self._context)
            server.login(self._login, self._password)
        except (gaierror, ConnectionRefusedError):
            print('Failed to connect to the server. Bad connection settings?')
        except smtplib.SMTPServerDisconnected:
            print('Failed to connect to the server. Wrong user/password?')
        except smtplib.SMTPException as e:
            print('SMTP error occurred: ' + str(e))
        else:
            return server
        return None
    
    def _get_HTML_table_from_list(self,mylist):
        
        html_table="""<table align="center" width="700" border="1" cellspacing="1" cellpadding="0" class="em_main_table" style="width:700px;"><tr>{_table_header}</tr>{_table_body}</table>"""
        table_row="""<tr>{_table_row}</tr>"""
        table_header="""<th>{_header_item}</th>"""
        table_data="""<td align="center">{_item}</td>"""
        header_str=""
        for header_item in mylist[0]:
            header_str += table_header.format(_header_item=header_item)
        
        #create table data 
        body_str=""
        for i in range(1,len(mylist)):
            row_str=""
            for item in mylist[i]:
                row_str += table_data.format(_item=item)
            body_str += table_row.format(_table_row=row_str)
                
        html_table = html_table.format(_table_header=header_str,_table_body=body_str)
        return html_table
        


    def send_report(self,controlFile_validation,process_validation):
        server = self._connect_mail_server()
        if server==None:
            print("Error occured while login!")
            return
        message = MIMEMultipart("alternative")
        message["Subject"] = "Validation Report"
        message["From"] = self._sender
        
        altText = """Hi, this is an alternate text. Email could not be rendered."""
        mail_template = open("./email_template.txt")
        line = mail_template.readline();
        html = ""
        while line:
            html += line
            line = mail_template.readline()
        
        #create table for control validation and process validation
        controlFile_validation_table = self._get_HTML_table_from_list(controlFile_validation)
        process_validation_table = self._get_HTML_table_from_list(process_validation)
#        print(html)
        message.attach( MIMEText(altText,"plain"))
        message.attach(MIMEText(html,"html"))
        
        file = open('./contacts.csv')
        reader = csv.reader(file)
        next(reader)
        for name,email in reader:
            message["To"]=email
            server.sendmail(self._sender, email, message.as_string().format(_name=name,\
                        _controlFile_validation_table=controlFile_validation_table,\
                        _process_validation_table=process_validation_table))
            print("Sent to {0}".format(name))
        
if __name__=="__main__":
    #dummy data
    control_file_validation = [["entity","validation"],
                               ["abc.xls","passed"],
                               ["def.csv","failed"],
                               ["ghi.txt","error"],
                               ["jkl.json","passed"]]
    
    process_validation= [["entity","cdc","load","comments"],
                         ["abc.json","	unsusseful","	unsuccessful","	check the log details"],
                         ["abc.csv","unsusseful","unsuccessful","check the log details"]]
    
    emailer = EmailReporter("email","password","smtp.gmail.com")
    emailer.send_report(control_file_validation,process_validation)