'''cleanmail.py - test program for cleaning up email text dump for text processing'''
import sys

try:
    with open('mailsample.txt') as text_file:
        mail_data = text_file.read().replace('\n', '')
except FileNotFoundError:
    sys.exit('Error: Expecting mailsmaple.txt in current folder')

delete_list = ['\\r\\n', 'From:', 'To:', 'Subject:', 'Sent:', 'https://', '@microsoft.com', '<', '>']
for crap in delete_list:
    mail_data = mail_data.replace(crap, '')

clean_data = mail_data.replace('_', ' ')
print(clean_data)