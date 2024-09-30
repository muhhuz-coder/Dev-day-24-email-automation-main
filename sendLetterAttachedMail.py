import datetime
from email.mime.text import MIMEText
import smtplib
from email.message import EmailMessage

# TODO Implement correct way of fetching the sender's password
# Temporary method of getting the sender's password
# passFile = open("pass.txt", "r")
# password = passFile.readlines(  )[0].lstrip().rstrip()
password=""

def sendPdfAttachmentMail(emailAddress, attachment,htmlContent):

    # TODO Set global variable for sender
    sender = "acm.khi@nu.edu.pk"

    senderPassword = password
    recieverMail = emailAddress
    msg = EmailMessage()
    msg["Subject"] = "Congratulations on Your Selection to the DevDay'24 Team"
    msg["From"] = sender
    msg["To"] = recieverMail
    msg["Cc"] = "k200240@nu.edu.pk"

    # Attach the pdf
    with open(attachment, "rb") as content_file:
        content = content_file.read()
        msg.add_attachment(
            content, maintype="application", subtype="pdf", filename=attachment
        )

    # Attaching the email design
    emailBody = MIMEText(htmlContent, "html")

    msg.attach(emailBody)

    # Sending Messages...
    print(f"[+] Sending attachment mail to {recieverMail}")

    with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
        try:

            # Raised a test exception to test exception handling
            # raise smtplib.SMTPResponseException(12, "nahi hora bhai")

            smtp.starttls()
            smtp.login(sender, senderPassword)
            smtp.send_message(msg)

            print(f"[+] Successfully sent mail to {recieverMail}")

            # A last minute check to see if the sent mail is correct
            if recieverMail == sender:
                check = input("All set? (y/n):")

                if check.lower() != "y":
                    return 0

            return True                
        except smtplib.SMTPRecipientsRefused as e:
            print(f"\n[-] Failed to send mail to {recieverMail}: Refused email.")

            # loging refused emails
            unsuccessfulMailLog = open("notSentToEmails.log", "a")
            unsuccessfulMailLog.write(
                f"{recieverMail} | Error: {e} | {datetime.datetime.now()}\n"
            )
            unsuccessfulMailLog.close()

        except smtplib.SMTPAuthenticationError:
            #invalid sender email or password
            print("\n[-] Authentication error: Failed to authenticate with the SMTP server. Please check your email and password.")

        except smtplib.SMTPResponseException as e:
            print(f"\n[-] Could not send mail to {recieverMail}: Error: {e.smtp_error}")

            # Logging the mail as not sent with error message and timestamp
            unsuccessfulMailLog = open("notSentToEmails.log", "a")
            unsuccessfulMailLog.write(
                f"{recieverMail} | Error: {e.smtp_error} | {datetime.datetime.now()}\n"
            )
            unsuccessfulMailLog.close()

        return False

# sendPdfAttachmentMail("k230703@nu.edu.pk", "test.pdf")
