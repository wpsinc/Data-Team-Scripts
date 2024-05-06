"""
This Python script is designed to automate the process of drafting and opening emails in the default email client. It uses the webbrowser and os modules to open the email client, and the urllib.parse module to encode the email message into a URL-friendly format.

The script defines a class EmailDraft with four attributes: to, cc, subject, and message. The to and cc attributes are expected to be strings containing one or more email addresses separated by semicolons. The subject and message attributes are expected to be strings containing the subject line and body of the email, respectively.

The EmailDraft class has a method open_in_client which constructs a mailto: URL with the to, cc, subject, and message attributes. This URL is then opened using the os.startfile function, which opens the URL in the default program for handling mailto: URLs (typically the default email client). If os.startfile is not available (which can be the case on some non-Windows platforms), the webbrowser.open function is used as a fallback.

The script also defines a function generate_email_drafts which takes a list of dictionaries, where each dictionary contains the details for one email (the to, cc, subject, and message). This function creates an EmailDraft object for each dictionary and returns a list of these objects.

Finally, the script contains a list of dictionaries email_details with the details for several emails, and a loop that generates an EmailDraft for each dictionary and opens it in the default email client. This is where the actual work of the script is done.
"""

import webbrowser
import os
import urllib.parse


class EmailDraft:
    def __init__(self, to, cc, subject, message):
        self.to = to
        self.cc = cc
        self.subject = subject
        self.message = message

    def open_in_client(self):
        message = urllib.parse.quote(self.message)
        to_emails = ";".join(email.strip() for email in self.to.split(";"))
        cc_emails = ";".join(email.strip() for email in self.cc.split(";"))
        url = f"mailto:{to_emails}?cc={cc_emails}&subject={self.subject}&body={message}"
        try:
            os.startfile(url)
        except AttributeError:
            webbrowser.open(url)


def generate_email_drafts(email_details):
    email_drafts = []
    for details in email_details:
        email_drafts.append(
            EmailDraft(
                details["to"], details["cc"], details["subject"], details["message"]
            )
        )
    return email_drafts


email_details = [
    {
        "to": "phil.marmet@abus.com",
        "cc": "jeremy.anderson@wps-inc.com",
        "subject": "WPS Monthly Report for Abus",
        "message": "Hello Phil, \r\n\r\nAttached is the monthly report for Abus. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "jrowlands@burlybrand.com",
        "cc": "jeremy.anderson@wps-inc.com",
        "subject": "WPS Monthly Report for Burly Brand",
        "message": "Hello Jody, \r\n\r\nAttached is the monthly report for Burly Brand. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "jcastro@dunlopmotorcycletires.com; sdewey@dunlopmotorcycletires.com; mcotter@srnatire.com",
        "cc": "",
        "subject": "WPS Monthly Report for Dunlop",
        "message": "Hello All, \r\n\r\nAttached is the monthly report for Dunlop. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "Paul@giviusa.com",
        "cc": "jlehan@wps-inc.com",
        "subject": "WPS Monthly Report for Givi",
        "message": "Hello Paul, \r\n\r\nAttached is the monthly report for Givi. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "ddines@wps-inc.com",
        "cc": "",
        "subject": "WPS Monthly Report for KFI",
        "message": "Hello Darren, \r\n\r\nAttached is the monthly report for KFI. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "paulg@nationalcycle.com",
        "cc": "jlehan@wps-inc.com",
        "subject": "WPS Monthly Report for National Cycle",
        "message": "Hello Paul, \r\n\r\nAttached is the monthly report for National Cycle. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "thomas.bagnaschi@rizoma.com",
        "cc": "jlehan@wps-inc.com",
        "subject": "WPS Monthly Report for Rizoma",
        "message": "Hello Thomas, \r\n\r\nAttached is the monthly report for Rizoma. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "dt@sawickispeedshop.com; chris@sawickispeedshop.com",
        "cc": "jeremy.anderson@wps-inc.com",
        "subject": "WPS Monthly Report for Sawicki Speed Shop",
        "message": "Hello All, \r\n\r\nAttached is the monthly report for Sawicki Speed Shop. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "Helge.Koenig@sp-united.com; brian.valverde@sp-united.com; Gerald.Samer@sp-united.com; Ryan.Lewis@sp-united.com",
        "cc": "jlehan@wps-inc.com",
        "subject": "WPS Monthly Report for SP Connect",
        "message": "Hello All, \r\n\r\nAttached is the monthly report for SP Connect. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "twelch@yoshimura-rd.com",
        "cc": "",
        "subject": "WPS Monthly Report for Yoshimura",
        "message": "Hello Tim, \r\n\r\nAttached is the monthly report for Yoshimura. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "curtis@hofmann-designs.com",
        "cc": "jeremy.anderson@wps-inc.com",
        "subject": "WPS Monthly Report for Hofmann Designs",
        "message": "Hello Curtis, \r\n\r\nAttached is the monthly report for Hofmann Designs. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "oriol@puigusa.com; carles@puigusa.com",
        "cc": "jlehan@wps-inc.com",
        "subject": "WPS Monthly Report for PUIG",
        "message": "Hello All, \r\n\r\nAttached is the monthly report for PUIG. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "Nicholas.Lowe-Hale@rammount.com",
        "cc": "jlehan@wps-inc.com",
        "subject": "WPS Monthly Report for RAM Mounts",
        "message": "Hello Nicholas, \r\n\r\nAttached is the monthly report for RAM Mounts. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "debd@nelsonrigg.com",
        "cc": "jlehan@wps-inc.com",
        "subject": "WPS Monthly Report for Nelson-Rigg",
        "message": "Hello, \r\n\r\nAttached is the monthly report for Nelson-Rigg. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "export@shinkotire.co.kr",
        "cc": "",
        "subject": "WPS Monthly Report for Shinko",
        "message": "Hello, \r\n\r\nAttached is the monthly report for Shinko. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "carlos@kandstech.com",
        "cc": "jlehan@wps-inc.com",
        "subject": "WPS Monthly Report for K&S",
        "message": "Hello, \r\n\r\nAttached is the monthly report for K&S. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "jason.mccune@uswe.com",
        "cc": "jlehan@wps-inc.com",
        "subject": "WPS Monthly Report for Giant Loop",
        "message": "Hello, \r\n\r\nAttached is the monthly report for Giant Loop. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    }
]

drafts = generate_email_drafts(email_details)

for draft in drafts:
    draft.open_in_client()
