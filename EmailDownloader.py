from O365 import Account
from datetime import timedelta
from dateutil.tz import tzutc
import dateutil.parser
from pathlib import Path


def boolean_prompter(prompt):
    while True:
        try:
            return {"true": True, "false": False}[input(prompt).lower()]
        except KeyError:
            print("Invalid input -- please enter 'true' or 'false'")


# Get the credentials and authenticate
clientId = input('Please provide the Office365 client ID: ')
clientSecret = input('Please provide the Office365 client secret: ')
credentials = (clientId, clientSecret)
account = Account(credentials)
if account.authenticate(scopes=['basic', 'mailbox', 'address_book', 'calendar', 'calendar_shared']):
    print('Authenticated!')
else:
    raise Exception('We failed to authenticate :-(')

mailbox = account.mailbox()
inbox = mailbox.inbox_folder()
print(inbox.total_items_count)

# Prompt the user for the file containing the email timestamps (it's best to pull this data from the ActiveSync logs
# which you can do en masse and targetted using the RDPS ActiveSync Parser available at:
# (https://github.com/theronielanddaronpodcastshow/ActiveSyncParser)
dateTimesFilePath = input('Please provide the location of the dates file: ')
dateTimesFile = open(dateTimesFilePath, 'r')
dateTimes = dateTimesFile.readlines()

# Ask the user if we are to grab the attachments
getAttachments = boolean_prompter('Download attachments (true/false): ')

# For each timestamp in the timestamp file, download the email
for dateTime in dateTimes:
    try:
        dateTimeObj = dateutil.parser.parse(dateTime.strip())
        oneSecBefore = (dateTimeObj - timedelta(seconds=1)).strftime('%Y-%m-%dT%H:%M:%S')
        oneSecAfter = (dateTimeObj + timedelta(seconds=1)).strftime('%Y-%m-%dT%H:%M:%S')
        msgs = mailbox.get_messages(limit=999, download_attachments=getAttachments,
                                    query='ReceivedDateTime gt ' +
                                          oneSecBefore +
                                          'Z and ReceivedDateTime lt ' +
                                          oneSecAfter
                                          + 'Z')
        for msg in msgs:
            if msg.subject is None:
                subject = ''
            else:
                subject = msg.subject.replace('/', '\\')
            emlFilePath = Path(dateTimesFilePath).parent / (msg.received.strftime(
                '%Y-%m-%dT%H:%M:%S.%fZ') + ' -- ' + subject + '.eml')
            print(msg.received.astimezone(tz=tzutc()).strftime('%Y-%m-%dT%H:%M:%S') + 'Z - ' + subject)
            msg.save_as_eml(to_path=emlFilePath)
    except Exception as err:
        print('That didn\'t work :-S: {}'.format(err))
