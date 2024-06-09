#!/usr/bin/env python
# coding: utf-8
"""
SMTPSender.py
Steven J Leathard (https://907sjl.github.io/)

This module provides a utility class for sending emails using SMTP.

Classes:
    SmtpSender - Represents an SMTP email message with methods to add content and send.\n

Top-Level Functions:
    send_email_message - Creates and sends an SMTP email message in one function call.\n
"""


from typing import List, Union

import time

import os.path as pth

import smtplib
from email import encoders
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase


class SmtpSender:
    """
    Represents an SMTP message with methods to add content and send.\n
    Methods:
        add_to_line_address - Adds an email address to the message To line.\n
        add_cc_line_address - Adds an email address to the message Cc line.\n
        add_bcc_line_address - Adds an email address to the message as a BCC.\n
        add_plain_text - Adds plain text content to the message body.\n
        add_html_text - Adds HTML text content to the message body.\n
        attach_file - Encodes and attaches a file to the message.\n
        send_email - Sends the message.\n

    Properties:
        sendFrom - The email address to send the message from.\n
        subject - The subject line of the email.\n
        smtp_server - The SMTP host address.\n
        to_addresses - A list of email addresses for the To line.\n
        cc_addresses - A list of email addresses for the CC line.\n
        bcc_addresses - A list of email addresses to receive the message without being listed in the header.\n
    """

    # Static template stylesheet for html email content
    _template_html_header: str = (
        """ 
        <style type="text/css">

        .ExternalClass {
            width: 100%;
        }
        .ExternalClass, .ExternalClass p, .ExternalClass span, .ExternalClass font, .ExternalClass td, .ExternalClass 
        div { line-height: 100%; }

        body {
            -webkit-text-size-adjust: none;
            -ms-text-size-adjust: none;
        }
        /* Prevents Webkit and Windows Mobile platforms from changing default font sizes. */

        * {
            font-family: Arial, Helvetica, sans-serif;
            font-size: 14px;
        }

        html,
        body {
            margin: 0;
            padding: 0;
            border: 0;
            outline: 0;
        }
        /* Resets all body margins and padding to 0 for good measure */

        table td {
            border-collapse: collapse;
            border-spacing: 0px;
            border: 1px solid;
            vertical-align: top;
            padding: 2px;
        }

        body,
        #body_style {
            background: #fff;
            min-height: 1000px;
            color: #000;
            font-family: Calibri, Arial, Helvetica, sans-serif;
            font-size: 12px;
        }

        html,
        body {
            line-height: 1.2;
        }
        h1,
        h2,
        h3,
        h4,
        h5,
        h6,
        p,
        i,
        b,
        a,
        ul,
        li,
        blockquote,
        hr,
        img,
        div,
        span,
        strong {
            font-family: Calibri, Arial, Helvetica, sans-serif;
            line-height: 1.2;
            margin:0;
            padding:0;
            margin-top:1.5em;
        }
        h1,
        h2,
        h3,
        h4,
        h5,
        h6 {
            font-family: Calibri, Arial, Helvetica, sans-serif;
            font-size: 18px;
            color: black;
        }
        h1 {
            font-family: Calibri, Arial, Helvetica, sans-serif;
            font-size: 24px;
            color: black;
        }
        /* A more sensible default for H1s */

        a,
        a:link {
            color: #2A5DB0;
            text-decoration: underline;
        }

        img {
            display: block;
            border: 0 none;
            outline: none;
            line-height: 100%;
            outline: none;
            text-decoration: none;
            vertical-align: bottom;
        }
        a img {
            border: 0 none;
        }

        small {
            font-size: 11px;
            line-height: 1.4;
        }
        small a {
            color: inherit;
            text-decoration: underline;
        }
        span.yshortcuts {
            color: #000;
            background-color: none;
            border: none;
        }
        span.yshortcuts:hover,
        span.yshortcuts:active,
        span.yshortcuts:focus {
            color: #000;
            background-color: none;
            border: none;
        }

        /*Optional:*/

        a:visited {
            color: #3c96e2;
            text-decoration: none
        }
        a:focus {
            color: #3c96e2;
            text-decoration: underline
        }
        a:hover {
            color: #3c96e2;
            text-decoration: underline
        }
        </style>
        """
    )

    def __init__(self, *args) -> None:
        """
        Basic constructor with positional parameters for email message field values. The parameters are
        for backwards compatibility. Create new messages with no parameters and use the "add_" methods
        to build the message.\n

        Args:
            txt: (1) Message content as text
            html: (2) Message content as HTML text
            subj: (3) Subject line content as text
            sendTo: (4) A list of To header recipients as comma delimited text
            bcc: (5) A list of BCC recipients as comma delimited text
            sendFrom: (6) The email address for the From header
            smtp_server: (7, optional) The SMTP host address
        """

        # Init properties
        self.sendFrom: str = 'do-not-reply@somewhere.org'
        self.subject: str = ''
        self.smtp_server: str = 'smtp.host.address'

        self._to_addresses: List[str] = []
        self._cc_addresses: List[str] = []
        self._bcc_addresses: List[str] = []
        self._recipients: List[str] = []
        self._smtp_message: MIMEMultipart = MIMEMultipart("alternative")
        self._html_header: str = ''
        self._text_parts: List[str] = []
        self._html_parts: List[str] = []

        # Process the optional parameters, if given
        to_list = None
        bcc_list = None

        for i in range(0, len(args)):
            if i == 0:
                self.add_plain_text(args[0])
            elif i == 1:
                self.add_html_text(args[1])
            elif i == 2:
                self.subject = args[2]
            elif i == 3:
                to_list = [a for a in args[3].split(",")]
            elif i == 4:
                bcc_list = [a for a in args[4].split(",")]
            elif i == 5:
                self.sendFrom = args[5]
            elif i == 6:
                self.smtp_server = args[6]

        if not ((to_list is None) or (len(to_list) == 0)):
            self._recipients = to_list.copy()
            self._to_addresses = to_list.copy()

        if not ((bcc_list is None) or (len(bcc_list) == 0)):
            if self._recipients is None:
                self._recipients = bcc_list.copy()
            else:
                self._recipients.extend(bcc_list)
            self._bcc_addresses = bcc_list.copy()
        # END __init__

    def add_to_line_address(self, address: str) -> None:
        """Adds To line recipient(s) to the message. Use a comma delimited list."""
        self._to_addresses.extend(address.split(','))
        self._recipients.extend(address.split(','))

    def add_cc_line_address(self, address: str) -> None:
        """Adds Cc line recipient(s) to the message. Use a comma delimited list."""
        self._cc_addresses.extend(address.split(','))
        self._recipients.extend(address.split(','))

    def add_bcc_line_address(self, address: str) -> None:
        """Adds BCC recipient(s) to the message. Use a comma delimited list."""
        self._bcc_addresses.extend(address.split(','))
        self._recipients.extend(address.split(','))

    def add_plain_text(self, plain_text: str) -> None:
        """Adds plain text content to the message body."""
        self._text_parts.append(plain_text)

    def add_html_text(self, html_text: str, style_text: str = None) -> None:
        """Adds HTML content to the message body using the optional stylesheet text."""
        if len(self._html_header) == 0:
            if style_text is None:
                self._html_header = self._template_html_header
            else:
                self._html_header = style_text
        self._html_parts.append(html_text)

    def attach_file(self, file_path: str) -> None:
        """Attaches a file to the message."""
        file_name = pth.basename(file_path)
        with open(file_path, "rb") as attachment:
            file_part = MIMEBase("application", "octet-stream")
            file_part.set_payload(attachment.read())
        encoders.encode_base64(file_part)
        file_part.add_header("Content-Disposition", f"attachment; filename= {file_name}")
        self._smtp_message.attach(file_part)

    @property
    def to_addresses(self) -> List[str]:
        return self._to_addresses

    @to_addresses.setter
    def to_addresses(self, value: List[str]) -> None:
        self._to_addresses = value.copy()
        self._recipients.extend(self._to_addresses)

    @property
    def cc_addresses(self) -> List[str]:
        return self._cc_addresses

    @cc_addresses.setter
    def cc_addresses(self, value: List[str]) -> None:
        self._cc_addresses = value.copy()
        self._recipients.extend(self._cc_addresses)

    @property
    def bcc_addresses(self) -> List[str]:
        return self._bcc_addresses

    @bcc_addresses.setter
    def bcc_addresses(self, value: List[str]) -> None:
        self._bcc_addresses = value.copy()
        self._recipients.extend(self._bcc_addresses)

    def send_email(self, enable_subj_timestamp: bool = True) -> None:
        """
        Sends this message via SMTP. Requires the following properties to be set at a minimum:\n
             to_addresses - A list of email addresses for the message To line.\n
             sendFrom - The email address to send the message from.\n
             smtp_server - The SMTP host address.\n

        Args:
             enable_subj_timestamp - Adds the current timestamp to the subject line if True.\n
        """

        # Assertions
        if (self._to_addresses is None) or (len(self._to_addresses) == 0):
            raise Exception("An SMTP message must have a To: line address to be sent.")
        if (self.sendFrom is None) or (len(self.sendFrom) == 0):
            raise Exception("An SMTP message must have a From: line address to be sent.")
        if (self.smtp_server is None) or (len(self.smtp_server) == 0):
            raise Exception("An SMTP message must have an SMTP host address to be sent.")

        # Optional subject line timestamp
        timestamp_str = time.strftime('%Y-%m-%d %H:%M')
        if enable_subj_timestamp:
            self.subject = f"{self.subject} {timestamp_str}"

        # Set SMTP header fields (no BCC header by design)
        self._smtp_message['Subject'] = self.subject
        self._smtp_message['From'] = self.sendFrom
        self._smtp_message['To'] = ','.join(self._to_addresses)
        if len(self._cc_addresses) > 0:
            self._smtp_message['Cc'] = ','.join(self._cc_addresses)

        # Add plain text message content, if any
        if len(self._text_parts) > 0:
            text_part = MIMEText('\n'.join(self._text_parts), "plain")
            self._smtp_message.attach(text_part)

        # Add html text message content, if any
        if len(self._html_parts) > 0:
            html_part = MIMEText(self._html_header + '<BR>' + (''.join(self._html_parts)), "html")
            self._smtp_message.attach(html_part)

        # Send the message to all recipients including BCC addresses
        with smtplib.SMTP(self.smtp_server) as s:
            s.sendmail(self.sendFrom, self._recipients, self._smtp_message.as_string())
        # END send_email()
    # END CLASS SmtpSender


def send_email_message(to_addresses: Union[str, List[str]],
                       subject: str,
                       html_text: str,
                       style: str = None,
                       from_address: str = None,
                       cc_addresses: Union[str, List[str]] = None,
                       bcc_addresses: Union[str, List[str]] = None) -> None:
    """
    Creates and sends an SMTP email message in one function call.\n

    Args:
         to_addresses - Either a list of email addresses or a comma delimited string of email addresses for To line.\n
         subject - The subject line for the email message.\n
         html_text - Either HTML or plain text content for the email message.\n
         style - (optional) Stylesheet text for the HTML message content.\n
         from_address - (optional) Email address to send the message from, defaults to the "do not reply" address.\n
         cc_addresses - (optional) Either a list of email addresses or comma delimited string of Cc email addresses.\n
         bcc_addresses - (optional) Either a list of email addresses or comma delimited string of BCC email addresses.\n
    """

    msg = SmtpSender()

    # To line addresses are required
    if type(to_addresses) is str:
        to_list = to_addresses.split(',')
    elif type(to_addresses) is list:
        to_list = to_addresses
    else:
        raise Exception('Address list for To line not recognized')
    msg.to_addresses = to_list

    # Cc line addresses are optional
    if not (cc_addresses is None):
        if type(cc_addresses) is str:
            cc_list = cc_addresses.split(',')
        elif type(cc_addresses) is list:
            cc_list = cc_addresses
        else:
            raise Exception('Address list for Cc line not recognized')
        msg.cc_addresses = cc_list

    # BCC addresses are optional
    if not (bcc_addresses is None):
        if type(bcc_addresses) is str:
            bcc_list = bcc_addresses.split(',')
        elif type(bcc_addresses) is list:
            bcc_list = bcc_addresses
        else:
            raise Exception('BCC address list not recognized')
        msg.bcc_addresses = bcc_list

    # From address is optional
    if not (from_address is None):
        msg.sendFrom = from_address

    msg.subject = subject
    msg.add_html_text(html_text, style)

    msg.send_email()
    # END send_email_message()


# MAIN


if __name__ == "__main__":
    """Test harness. This should only really be invoked during development by running this module alone."""

    test_txt = "This is a sample message"
    test_html = "<p>This is a <strong>sample message</strong>."
    test_subj = "Sample Message from the SmtpSender module."
    test_sendTo = "do-not-reply@somewhere.org"
    test_cc = "do-not-reply@somewhere.org"
    test_bcc = "someone@somewhere.org"
    test_sendFrom = "do-not-reply@somewhere.org"

    # Test for backwards compatibility
    test_msg = SmtpSender(test_txt,
                          test_html,
                          test_subj,
                          test_sendTo,
                          test_bcc,
                          test_sendFrom)
    test_msg.send_email(enable_subj_timestamp=False)

    # Be nice to the spam monitor
    time.sleep(60)

    # Test top-level convenience function
    send_email_message(test_sendTo,
                       test_subj,
                       test_html,
                       from_address=test_sendFrom,
                       cc_addresses=test_cc,
                       bcc_addresses=test_bcc)

    # Be nice to the spam monitor
    time.sleep(60)

    # Test dynamic email construction and attachments
    test_msg = SmtpSender()
    test_msg.add_to_line_address(test_sendTo)
    test_msg.add_bcc_line_address(test_bcc)
    test_msg.add_cc_line_address(test_cc)
    test_msg.sendFrom = test_sendFrom
    test_msg.subject = test_subj
    test_msg.add_plain_text(test_txt)
    test_msg.add_plain_text(test_txt)
    test_msg.add_plain_text(test_txt)
    test_msg.attach_file('test_attachment.jpg')
    test_msg.send_email(enable_subj_timestamp=False)

    print("Test complete")
