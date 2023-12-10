from datetime import datetime
import base64
from email.mime.text import MIMEText
from email.utils import make_msgid
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from requests import HTTPError


def get_professor_email():
    """
    Asks for professors email
    Checks they end in @babson.edu
    Splits name @babson.edu
    Confirms they are the correct recipient
    Returns recipients email address and professors name with first letter capitalized
    """
    prof_confirmation = "0"
    babson_code = "@babson.edu" # 11 chars long
    while prof_confirmation != "1": 
        recipient = input("Professor's email: ") 
        while recipient[-11:] != babson_code: # makes sure its going to babson prof
            recipient = input("This is not going to a Babson Professor. Try Again: ")
        prof_confirmation = input(f"\nEnter 1 to confirm you want to send this email to {recipient} \nEnter anything else to retype: " )
        prof_last_name, _ = recipient[1:].split("@") # first letter is first initial so start at index 1 and stop at @ to get last name
    return recipient, prof_last_name.title() # capitalize first letter of last name

def get_date():
    """
    Gets date of the exam
    Emails are sent the day the student takes the exam so "now" can be used
    Formats it in mm/dd/yyyy
    Returns todays date
    """
    today = datetime.now()
    today_formatted = today.strftime("%m/%d/%Y")
    return today_formatted

def get_class_info():
    """
    Gets course code ex: OIM3640
    Gets section number
    Adds a zero in front of any single digit section for formatting
    Returns course code and section number
    """
    confirmation = "0"
    while confirmation != "1":
        course_code = input("\nCourse code: ").upper()
        section = input("Section #: ")
        if len(section) == 1: # changes section from 1 to 01 according to template
            section = "0" + section
        confirmation = input(f"\nEnter 1 to confirm \nThis is section {section} for class {course_code}\nEnter anything else to retype: " )
    return course_code, section 

def assessment_type():
    """
    Asks whether it's a quiz or exam
    Returns type of assessment 
    """
    assessment = ""
    while assessment not in ["e", "q"]:
        assessment = input("\nPlease enter either Q for Quiz or E for Exam: ").lower()
    if assessment == "q":
        as_type = "QUIZ"
    if assessment == "e":
        as_type = "EXAM"
    return as_type

def students(section):
    """
    Asks user for all student names and status iteratively
    Returns list of student dictionaries
    """
    student_data = [] 
    while True:
        student_name = input("\nEnter Students First & Last Name\nType done when finished: ").title() # Capitalizes name
        if student_name.lower() == "done": # when all names are entered
            break
        
        completion_status = "-"
        while completion_status == "-": # iterates until status is assigned
            status = input("Enter C if completed. X if not completed. Y for other: ").lower()
            if status == "c":
                completion_status = "Completed"
            elif status == "x":
                completion_status = "Did Not Attend"
            elif status == "y":
                completion_status = input("Enter Custom Response: ").title() 
            else:
                print("Invalid Input. C for Completed. X for Did Not Complete. Y for Other")

        student_data.append({"Student Name": student_name, "Course Section": section, "Status": completion_status}) # adds student dicts to list
    return student_data


def create_table(student_data, course, section):
    """
    Uses student_data, course, and section to create table
    Color codes based on completion status
    Uses HTML because allows  to embed the table in the email & looks appealing
    """
    # Create HTML content
    table_html = """
    <html>
    <body>
    <h2>{}.{}</h2>
    <table border="1" cellspacing="0" cellpadding="5">
        <tr>
            <th>Student Name</th>
            <th>Course Section</th>
            <th>Status</th>
        </tr>
    """.format(course, section)

    for student in student_data:
        status_color = 'green' if student["Status"] == "Completed" else ('red' if student["Status"] == "Did Not Attend" else 'grey') # color codes status
        table_html += "<tr style = 'background-color:{}'>".format(status_color) # colors entire row
        table_html += "<td>{}</td>".format(student["Student Name"])
        table_html += "<td>{}</td>".format(student["Course Section"])
        table_html += "<td>{}</td>".format(student["Status"])
        table_html += "</tr>"
        
    table_html += """
    </table>
    </body>
    </html>
    """
    return table_html

def send_email(prof_last_name, course, section, assessment, date, table_html, recipient_email):
    """
    Uses gmail API from https://developers.google.com/gmail/api/guides to access associated gmail account & send templated message
    Uses all other functions to build proper email
    WILL DIRECT YOU TO GOOGLE ACCOUNTS
    SIGN IN AS LILYDASEXAMPROCTOR@GMAIL.COM
    PASSWORD: oim3640ismyfavoriteclass
    Google has not verified the app yet, so click advanced in the bottom left
    Hit "Go to DAS Accessibility Email Sender (unsafe)"
    Hit Continue
    """
    SCOPES = [
            "https://www.googleapis.com/auth/gmail.send" # specifies permissions app has when accessing user gmail data
        ]
    flow = InstalledAppFlow.from_client_secrets_file("gmail.json", SCOPES) # sets OAuth 2.0 flow using gmail.json
    creds = flow.run_local_server(port=0) # runs local server to verify user & get credentials

    service = build("gmail", "v1", credentials=creds) # builds gmail api using credentials
    
    boundary = make_msgid() # creates boundary identifier for email
    message = MIMEText(f"""
Dear Professor {prof_last_name},<br><br>

As you know, the following students with academic accommodations in your <b>{course}.{section}</b> completed their Department of Accessibility Services (DAS) proctored <b>{assessment}</b> at the DAS testing center on <b>{date}</b>, and their exams may be picked up at your convenience during normal business hours this week or next week (Monday-Friday, 8:30am-4:30pm) from the DAS front desk, in Hollister 220:<br>

{table_html} <br>

Students with registered academic accommodations in your section(s) who were scheduled/notified to complete the final exam proctored by DAS, but did not arrive to complete the exam are marked above in red.<br> Please do not hesitate to reach out to us with any questions or concerns.<br><br>

Best,<br>
The DAS Team<br>
Lily Steinwold
""",
"html",
"utf-8"
)
    message["to"] = recipient_email
    # message["cc"] = "accessibilityservices@babson.edu" FOR PRACTICAL USE ONLY- DO NOT UNHIGHLIGHT AS IT WILL CC THEM
    message["subject"] = f"{course}.{section} Assessments Completed"
    message["Content-Type"] = f"text/html; charset=utf-8; boundary={boundary}"
    
    raw_message_bytes = base64.urlsafe_b64encode(message.as_bytes()) # encodes email msg bytes
    raw_message_str = raw_message_bytes.decode("utf-8") # decodes into utf-8 string
    create_message = {"raw": raw_message_str} # creates dictionary with raw email message

    try:
        message = (service.users().messages().send(userId="me", body=create_message).execute()) # Sends email
        print(f"sent message to {message} Message Id: {message['id']}")
    except HTTPError as error:
        print(f"An error occurs: {error}")
        message = None
   
def main():
    # Gathering data to create email & table
    recipient_email, prof_last_name = get_professor_email()
    date = get_date()
    course, section = get_class_info()
    assessment = assessment_type()
    student_data = students(section)
    table_html = create_table(student_data, course, section)

    # Sending email
    send_email(prof_last_name, course, section, assessment, date, table_html, recipient_email)

if __name__ == "__main__":
    main()







