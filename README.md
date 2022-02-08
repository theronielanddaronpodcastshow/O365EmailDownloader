# O365 Email Downloader
The RDPS O365 Email Downloader is a tool largely designed for to aid digital forensics professionals (although it can be used by users, system administrators, and others, for backing up emails, mass email downloads, etc.).  Simply put, the python script will log into Office 365 and download all emails (and their attachments if desired) based on the _received_ date-time.  To do this, it prompts the user for the credentials, then it prints out how many emails are in the user's inbox.

Next the script asks for the file containing the timestamps -- one timestamp per line.  It will use these timestamps to retrieve the emails -- if an email with the specific timestamp isn't found, it will try one second before and one second after -- and, based on the user's response to the prompt, it will download attachments, as well.  Each email is stored in an EML, with the EMLs being in the same folder structure as the timestamp file.

This tool is largely designed to allow forensics professionals to, in a targetted manner, download all emails (and their attachments, if desired) that someone had access to and/or had themselves downloaded.

For example, let's say that an Office 365 account had been compromised and the logs show unauthorised email, calendar, or address book (we get them all) access.  If the digital forensics professional extracts the received timestamps from the logs (e.g., from the ActiveSync logs via the RDPS ASP), and places said timestamps into a single file, with each timestamp sitting on a line by itself, they can run this tool and have it download all emails, calendar data, and address book data accessed by the individual.  These emails (all EML format) can then be included in the report or, itself, parsed and/or added to the case file (e.g., FTK or autopsy).

## Requirements ##
>Python 3.4
O365 (`pip install O365`)

## Future Work ##
1. Write an installer... I guess... it kind of doesn't need one...
2. Write some examples in this README (although it seems fairly straightforward to the author... who is clearly unbiased)
