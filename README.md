# Nephthis-BP-VBO---Internal-MAPI-v2
Replacement for OOB Outlook VBO provided by BP. Condenses and simplifies complicated &amp; redundant actions, while providing additional functionality. Contains support for multiple profiles and shared inboxes on the same profile.

## Get Mail
Grabs and stores mail into a collection.

### Inputs:
* Profile - Text
  * Optional - Outlook profile on system. Defaults to system default, IE Outlook
* Subfolder - Text
  *  Optional - Inbox subfolder to start in, up to two levels deep
* Include Unread - Flag
  * Optional  - Include Unread. Defaults to True.
* Include Read - Flag
  * Optional - Include Read. Defaults to False.
* Sender Name - Text
  * Optional - Sender name to search for. * wildcards accepted.
* Sender Email - Text
  * Optional - Sender E-Mail to search for. * wildcards accepted.
* Subject - Text
  * Optional - Subject to search for. * wildcards accepted.
* Message - Text
  * Optional - Message body text to search for. * wildcards accepted.
* Earliest Date - DateTime
  * Optional - Earliest date to search for.
* Latest Date - DateTime
  * Optional - Latest date to search for.
* HTML - Flag
  * Optional - Get mail as HTML
* Filter Expression - Text
  * Optional - Overrides all above inputs other than Profile. Filter for message. See action description/documentation for more info.
* Shared E-mail - Text
  * Optional - Address to search within if multiple inboxes are on the same profile.
* Outlook Folder ID - Number
  * Optional - Defaults to 6 for inbox. Google OlDefaultFolders for list of values.
* Mark Read - Flag
  * Optional - Defaults to False. Will mark all results as read if True.
* MoveTo - Text
  * Optional - Inbox subfolder name to move the message to after reading
* Delete - Flag
  * Optional - Deletes the message upon reading
* MaxCount - Number
  * Optional - Specify a maximum number of e-mails to pull/return
* Oldest First? Flag
  * Optional - Defaults to True
* Account Display Name - Text
  * Optional. DisplayName of the email account. Default account is used if omitted.

### Outputs:
* Items - Collection
* Item Count - Number

## Send Mail
Uses code to automate Outlook for the purpose of sending mail.

### Inputs:
* Profile - Text
  * Optional - Defines profile to send mail through. If not set, uses system default (IE Outlook)
* SendFrom - Text
  *  Optional - Specifies which address to send as
* To - Text
  * Requred - Semicolon delimited list of e-mail addresses.
* CC - Text
  * Optional - Semilcolon delimited list of CC addresses
* BCC - Text
  * Optional - Semicolon delimited list of BCC addresses
* Subject - Text
  * Required - Text for Subject line
* MailBody - Text
  * Required - Text for mail body
* Attachment - Text
  * Optional - Comma delimited list of file paths to attach
* High - Flag
  * Required - High Importance?
* HTML - Flag
  * Optional - Send message as HTML. Defaults to False

## Forward Mail
Uses code to automate Outlook for the purpose of forewarding mail.

### Inputs:
* Profile - Text
  * Optional - Defines profile to send mail through. If not set, uses system default (IE Outlook)
* SendFrom - Text
  *  Optional - Specifies which address to send as
* To - Text
  * Requred - Semicolon delimited list of e-mail addresses.
* CC - Text
  * Optional - Semilcolon delimited list of CC addresses
* BCC - Text
  * Optional - Semicolon delimited list of BCC addresses
* New Subject - Text
  * Optional - Text for new Subject line in case you don't send with the forward subject
* MailBody - Text
  * Required - Text for mail body
* Attachment - Text
  * Optional - Comma delimited list of file paths to additionally attach
* NoForwardAttachments - Flag
  * Optional - Don't include the attachments in the forwarded e-mail
* High - Flag
  * Required - High Importance?
* HTML - Flag
  * Optional - Send message as HTML. Defaults to False
* Mail ID - Text
  * Required - ID of the mailitem you wish to forward

## Reply to Email
Uses code to automate Outlook for the purpose of replying to mail. Only does one recipient.

### Inputs:
* Profile - Text
* SendFrom - Text
* New Subject - Text
* MailBody - Text
* Attachment - Text
* High - Flag
* HTML - Flag
* Mail ID - Text

## Reply All to Email
Uses code to automate Outlook for the purpose of replying to mail. Sends to all recipients.

### Inputs:
* Profile - Text
* SendFrom - Text
* New Subject - Text
* MailBody - Text
* Attachment - Text
* High - Flag
* HTML - Flag
* Mail ID - Text

## Save Email As File
Stores a specifried email as a ".msg" file on any file location (file path).

### Inputs:
* File Path - Text
* File Name - Text
* Entry ID - Text
* Profile - Text

## Force Sync
Forces a sync (send/receive) action within Outlook. Primarily used for outbound, but can be done for either direction if necessary.

### Inputs:
* Profile - Text
* SharedAddress - Text
  * Optional - Mail address of shared inbox on same profile
* Outlook Folder ID
  *  Mandatory. Google OlDefaultFolders Enumeration, 6 = Inbox, 5 = Sent

## Resolve_Name
Attempts to find the e-mail address of a supplied display name. For exmample, "Joe Smith" could resolve to "joe.smith@myorg.com". Relies on organizational and address book lookup.

### Inputs:
* Profile - Text
* displayName - Text

### Outputs:
* smtpName

## Delete Mail
This clears out the 'Deleted Items' folder of a specified account. Defaults to all on profile.

### Inputs:
* Profile - Text
* sharedAddress - Text

## Delete Single Mail
This deletes one mail by Entry ID from any mailboxes on the designated profile.

### Inputs:
* Profile - Text
* ID - Text

## Mark single item read/unread
This marks one mail read/unread by Entry ID. Wioll work on all mailboxes on the specified profile.

### Inputs:
* ID - Text
* Profile - Text
* MarkRead - Flag
  * Set to False if the item should be marked as unread

## Move Single Mail
This moves one mail from the specified folder to another specified folder by Entry ID. If the target is a nested subfolder, use Move Single Mail with Folder Path instead.

### Inputs:
* ID - Text
* Profile - Text
* RootfolderID - Number
* MoveTo - Text
* sharedAddress - Text

### Outputs:
* New Entry ID - Text
  * New MailID for the mailitem after it has been moved. This should never be the same as the original ID.

## Move Single Mail with Folder Path
This moves one mail from the specified folder to another specified folder by Entry ID.

### Inputs:
* ID - Text
* Profile - Text
* RootfolderID - Number
* MoveTo - Text
  * Expects a mailfolder path hierarchy similar to "Inbox\Subfolder\Another Subfolder\Yet Another Subfolder"
* sharedAddress - Text

### Outputs:
* New Entry ID - Text
  * New MailID for the mailitem after it has been moved. This should never be the same as the original ID.

## Mark Tasked/Complete
This marks an e-mail as tasked/complete or removes the status by Entry ID. Will work on any inbox attached to the same profile.

### Inputs:
* ID - Text
* Profile - Text
* Subfolder - Text
* markAs - Text
  * Optional - Set as "flag" or "done". Blank or any other value will clear the status. Not case-sensitive.

## Save Attachment
Downloads a specific attachment to a folder. Eg. use cases are "*.txt" or "*.docx" etc.

### Inputs:
* Entry ID - Text
* Folder Path - Text
* File Name - Text
* Profile - Text

## Save Attachments
Downloads all attachments to a folder.

### Inputs:
* Entry ID - Text
* Folder Path - Text
* Profile - Text
