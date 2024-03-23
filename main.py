from urllib.parse import urlparse
from appscript import app, k
from mactypes import Alias
from pathlib import Path
import csv


def generate_emails(name_string, url):
    """
    Generates 5 email addresses in different formats based on a given name and URL.

    Args:
        name_string (str): A string in the format "<firstname> <lastname>".
        url (str): The domain part of the email address.

    Returns:
        list: A list of 5 generated email addresses.
    """

    firstname, lastname = name_string.lower().split(" ")
    first_initial = firstname[0]
    last_initial = lastname[0]

    emails = [
        f"{firstname}@{url}",
        f"{first_initial}{lastname}@{url}",
        f"{firstname}{lastname}@{url}",
        f"{firstname}{last_initial}@{url}",
        f"{lastname}@{url}",
        f"{firstname}.{lastname}@{url}",
    ]

    return emails


class Outlook(object):
    def __init__(self):
        self.client = app("Microsoft Outlook")


class Message(object):
    def __init__(
        self, parent=None, subject="", body="", to_recip=[], cc_recip=[], bcc_recip=[], show_=True
    ):

        if parent is None:
            parent = Outlook()
        client = parent.client

        self.msg = client.make(
            new=k.outgoing_message,
            with_properties={k.subject: subject, k.content: body},
        )

        self.add_recipients(emails=to_recip, type_="to")
        self.add_recipients(emails=cc_recip, type_="cc")
        self.add_recipients(emails=bcc_recip, type_="bcc")
        

        if show_:
            self.show()

    def show(self):
        self.msg.open()
        self.msg.activate()

    def add_attachment(self, p):
        # p is a Path() obj, could also pass string

        p = Alias(str(p))  # convert string/path obj to POSIX/mactypes path

        attach = self.msg.make(new=k.attachment, with_properties={k.file: p})

    def add_recipients(self, emails, type_="to"):
        if not isinstance(emails, list):
            emails = [emails]
        for email in emails:
            self.add_recipient(email=email, type_=type_)

    def add_recipient(self, email, type_="to"):
        msg = self.msg

        if type_ == "to":
            recipient = k.to_recipient
        elif type_ == "cc":
            recipient = k.cc_recipient
        elif type_ == "bcc":
            recipient = k.bcc_recipient

        msg.make(new=recipient, with_properties={k.email_address: {k.address: email}})


def createEmail(name_input, company, url):
    # name_input = input("Enter Names (separated by comma): ")

    names = name_input
    names = [name.strip() for name in names.split(",")]
    first_names = [name.split()[0] for name in names]
    # print(' or '.join(list1[:-1]) + ' or ' + list1[-1])
    email_names = first_names[0]
    if len(first_names) > 1:
        email_names = ", ".join(first_names[:-1]) + " and " + first_names[-1]

    # company = input("Enter Company: ")

    body = f"""
    <p>
    Hi {email_names},
    </p>

    <p>
    My name is Nikhil and I am a CS sophomore at the University of Washington who is very interested in an internship at {company}. I know your time is valuable so I’ll keep this brief.
    </p>

    <p>
    I have a lot of experience as a Software Developer for a company named CodeDay. I’ve been with CodeDay for 4 years now and have worked extensively with technologies like NodeJS, React, and SQL. I’ve also worked with other technologies like C++, Java and Rust outside of work over the past ~9 years. You can check out some of my projects at <a href="https://github.com/Nexite">https://github.com/Nexite</a>. One project that I’m particularly proud of is gql-server which is CodeDay’s main CMS system that I rewrote in TypeScript from the ground up. I’m very eager and willing to learn new technologies and would love to work at {company}.
    </p>

    <p>
    I’ve attached my resume and my LinkedIn can be found at <a href="https://www.linkedin.com/in/nikhilkgarg/">https://www.linkedin.com/in/nikhilkgarg/.</a>
    </p>

    <p>Thank you for your time!<br/>
    Nikhil Garg<br/>
    425-236-2215</p>
    """

    subject = "Internship Candidate with Strong Development Background"
    to_recip = []
    bcc_recip = []
    for name in names:
        emails = generate_emails(name, urlparse(url).hostname.replace("www.", ""))
        to_recip.append(emails[0])
        if len(emails) > 1:
            bcc_recip.extend(emails[1:])
    msg = Message(subject=subject, body=body, to_recip=to_recip, bcc_recip=bcc_recip, show_=False)
    p = Path("Nikhil Garg Resume.pdf")
    msg.add_attachment(p)

    msg.show()


filename = "geekwire200.csv"


with open(filename, "r") as csvfile:
    datareader = csv.reader(csvfile)
    for row in datareader:
        rank = int(row[0])
        # if you were in the first 40 companies, you were contacted manually :)
        if rank <= 40:
            continue
        company_name = row[1]
        url = row[2]
        info = row[3]
        print(info)
        people_names = input(f"Enter Names for {company_name} (separated by comma): ")
        while people_names != "":
            createEmail(people_names, company_name, url)
            people_names = input(
                f"Enter Names for {company_name} (separated by comma): "
            )
