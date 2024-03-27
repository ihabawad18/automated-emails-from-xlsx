# first lets read the excel file
# second lets generate a unique qr code for each student 
# third lets start sending the email procedure 
# fourth attach the Iftar-Photo along with the QR code along with a personalized 
# message for each entry in the excel file

from qr_code_generator import generate_qr_code
from openpyxl import load_workbook
from send_email import send_email_via_outlook

workbook = load_workbook(filename='Ramadan Iftar 2024 - Byblos Campus.xlsx', keep_vba=True)

# Access a specific sheet by name
sheet = workbook['Males']

i=2
while(sheet[f'A{i}'].value):
    # send email for not paid
    student_name = sheet[f'A{i}'].value+" "+sheet[f'B{i}'].value    
    if(sheet[f'E{i}'].value=="No"):
        print(student_name)
        print(sheet[f'D{i}'].value)
        print()
        subject = "Byblos Ramadan Iftar 2024 Reminder"
        message_to_send = f'''
            <html>
            <body>
            <p>Dear {student_name},</p>

            <p>We would like to thank you for registering for our Iftar gathering this Tuesday, the 26th of March, 2024.</p>

            <p>Our ticketing system shows that you still have not completed the payment for the Iftar. However, as you have already registered, you can pay on entry on Tuesday if you wish to join us. </p>

            <p>Please try to have exact change (15$) if possible, to make the transaction easier.</p>

            <p>We look forward to having you!</p>

            <p style="font-size: smaller;"><i>LAU Byblos Iftar Organizing Committee</i></p>
            </body>
            </html>
        '''

        send_email_via_outlook("email","password",sheet[f'D{i}'].value,subject,message_to_send,[])        
    i+=1
