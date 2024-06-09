# smtp-sender    
This is a simple utility class in Python that I use to send reports via SMTP.    

## Usage    
This example can be copied into another Python project and included in other Python source files. Copy the SJLTools 
package directory into your top-level project directory. Then include the SmtpSender module from SJLTools.    

```
from SJLTools.SMTPSender import SmtpSender as ss 
```    
Alternatively, import just the class if you don't want the top-level convenience function to send an email.    
```
from SJLTools.SMTPSender.SmtpSender import SmtpSender
```    

After which there are three possible patterns to follow in order to compile and send an email message: Compatible, 
Convenience, or Custom.    

### Compatible Use 
For backwards compatibility with prior code releases, the SmtpSender object can be instantiated with the message 
components passed to the class constructor:
```
test_msg = SmtpSender(test_txt,
                      test_html,
                      test_subj,
                      test_sendTo,
                      test_bcc,
                      test_sendFrom)
```    
Then send the message with a call to the **send_email** method, overriding the *enable_subj_timestamp* 
parameter default if you want to disable the timestamp in the subject line:    
```
test_msg.send_email(enable_subj_timestamp=False)
```    

### Convenience Function 
The SmtpSender module provides a top-level convenience function, **send_email_message**, that sends a 
message in one call:
```
send_email_message(test_sendTo,
                   test_subj,
                   test_html,
                   from_address=test_sendFrom,
                   cc_addresses=test_cc,
                   bcc_addresses=test_bcc)
```    
The only required parameters are the address to send To, the email subject line, and some text content that can be 
HTML or plain text but will be sent as HTML. The From email address can be overridden but defaults to the 
"do not reply" email address.    

### Custom Message 
Email messages can be compiled dynamically by incrementally adding content and headers. To create a message this way, 
start by instantiating an empty SmtpSender object:
```
test_msg = SmtpSender()
```    

Recipient addresses can be added one at a time, or in groups, by passing comma delimited lists of addresses to 
setter methods on the object:    
```
test_msg.add_to_line_address(test_sendTo)
test_msg.add_bcc_line_address(test_bcc)
test_msg.add_cc_line_address(test_cc)
```    

The From and Subject line message fields are directly set as properties on the object:    
```
test_msg.sendFrom = test_sendFrom
test_msg.subject = test_subj
```    

Text or HTML parts are added to the message body using methods on the object, either **add_plain_text** or 
**add_html_text**. Email clients probably won't show text content if any HTML content is sent.    
```
test_msg.add_plain_text(test_txt)
test_msg.add_plain_text(test_txt)
test_msg.add_plain_text(test_txt)
```    

Attachments are loaded and encoded using the **attach_file** method on the object. Pass the full path or the 
path relative to the current working directory as a parameter.    
```
test_msg.attach_file('test_attachment.jpg')
```    

Then send the message with a call to the **send_email** method, overriding the *enable_subj_timestamp* 
parameter default if you want to disable the timestamp in the subject line:    
```
test_msg.send_email(enable_subj_timestamp=False)
```    
