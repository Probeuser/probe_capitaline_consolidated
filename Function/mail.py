import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import date

# getting current date
today = date.today()


# function to send an mail
def send_email(api_name , exception_message):
    port = 587  # For starttls
    smtp_server = "smtp.gmail.com"
    sender_email = "probepoc2023@gmail.com"
    receiver_email = "probepoc2023@gmail.com"
    password = "rovqljwppgraopla"
    subject = f"Manual intervention required for {api_name}  | " + str(today)

    msg = "\n Hi , \n\n Please look in to the below issue \n\n Error Message: \n\n "
    msg = msg + str(exception_message)
    msg += "\n\n\n\n Thanks! "
        
    # Create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject

    # Add body to email
    message.attach(MIMEText(msg, "plain"))

    context = ssl.create_default_context()

    with smtplib.SMTP(smtp_server, port) as server:
        server.starttls(context=context)
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message.as_string())
        print("Mail sent successfully!")
        server.quit()