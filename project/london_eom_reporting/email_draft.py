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
        "subject": "WPS Monthly Report",
        "message": "Hello Phil, \r\n\r\nAttached is the monthly report for Abus. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "jrowlands@burlybrand.com",
        "cc": "kbrown@wps-inc.com",
        "subject": "WPS Monthly Report",
        "message": "Hello Jody, \r\n\r\nAttached is the monthly report for Burly Brand. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "jcastro@dunlopmotorcycletires.com; sdewey@dunlopmotorcycletires.com; mcotter@srnatire.com",
        "cc": "aabreu@wps-inc.com",
        "subject": "WPS Monthly Report",
        "message": "Hello All, \r\n\r\nAttached is the monthly report for Dunlop. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "Paul@giviusa.com",
        "cc": "jlehan@wps-inc.com",
        "subject": "WPS Monthly Report",
        "message": "Hello Paul, \r\n\r\nAttached is the monthly report for Givi. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "ddines@wps-inc.com",
        "cc": "",
        "subject": "WPS Monthly Report",
        "message": "Hello Darren, \r\n\r\nAttached is the monthly report for KFI. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "paulg@nationalcycle.com",
        "cc": "kbrown@wps-inc.com",
        "subject": "WPS Monthly Report",
        "message": "Hello Paul, \r\n\r\nAttached is the monthly report for National Cycle. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "thomas.bagnaschi@rizoma.com",
        "cc": "jlehan@wps-inc.com",
        "subject": "WPS Monthly Report",
        "message": "Hello Thomas, \r\n\r\nAttached is the monthly report for Rizoma. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "dt@sawickispeedshop.com; chris@sawickispeedshop.com",
        "cc": "kbrown@wps-inc.com; jeremy.anderson@wps-inc.com",
        "subject": "WPS Monthly Report",
        "message": "Hello All, \r\n\r\nAttached is the monthly report for Sawicki Speed Shop. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "Helge.Koenig@sp-united.com; brian.valverde@sp-united.com; Gerald.Samer@sp-united.com",
        "cc": "jlehan@wps-inc.com",
        "subject": "WPS Monthly Report",
        "message": "Hello All, \r\n\r\nAttached is the monthly report for SP Connect. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "twelch@yoshimura-rd.com",
        "cc": "aabreu@wps-inc.com",
        "subject": "WPS Monthly Report",
        "message": "Hello Tim, \r\n\r\nAttached is the monthly report for Yoshimura. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    },
    {
        "to": "curtis@hofmann-designs.com",
        "cc": "jeremy.anderson@wps-inc.com",
        "subject": "WPS Monthly Report",
        "message": "Hello Curtis, \r\n\r\nAttached is the monthly report for Hofmann Designs. Please feel free to reach out anytime with any questions or concerns regarding the content of this file. \r\nLondon Perry, \r\nProduct Content Specialist | Western Power Sports Inc.",
    }
]

drafts = generate_email_drafts(email_details)

for draft in drafts:
    draft.open_in_client()
