# Things to install via pip:
#
# pip3 install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
# (For Gmail API)
#
# pip3 install discord.py
# (For Discord API)
#
# pip3 install pyyaml
# (For yaml config loading)

# Discord dependencies
import discord

# Gmail dependencies
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Async dependencies
import asyncio

# General dependencies
import time
from datetime import date
import base64
from os import listdir
from os.path import isfile, join
from os import remove
import yaml

# Gmail token setup
# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']

def gmailInitialize(loadedConfig):
    """Gets the Gmail service

    Args:
      loadedConfig: configuration from yaml file
    Returns:
      Gmail service, based on credentials in the token
    """

    creds = None

    # Loading credentials from pickle
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If no valid creds, let user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                loadedConfig['gmail']['credFile'], SCOPES)
            creds = flow.run_local_server(port=0)
        # Save credentials for next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)
    return service

def getUnreadEmails(labelId, gmailService):
    """Gets all unread emails from Gmail with a certain label

    Args:
      labelId: Label id to get unread emails from
      gmailService: Gmail service
    Returns:
      Array of all unlread emails under the given label id
    """

    emails = []

    try:
        # Gmail's after tag breaks for most emails if you include a .0, so we need to remove that with rstrip
        query = 'after:'+(str(time.mktime(date.today().timetuple()))).rstrip('0').rstrip('.')
        emailResults = gmailService.users().messages().list(userId='me', labelIds=[labelId, 'UNREAD'], q=query).execute()

        if 'messages' in emailResults:
            emails.extend(emailResults['messages'])

        while 'nextPageToken' in emailResults:
            page_token = emailResults['nextPageToken']
            response = gmailService.users().messages().list(userId='me', labelIds=[labelId, 'UNREAD'], q=query, pageToken=page_token).execute()
            emails.extend(response['messages'])

        return emails
    except Exception as e:
        print(e)
        print('exception getUnreadEmails')
        return emails

def downloadPhotosFromEmail(gmailService, emailData):
    """Downloads all attached photos from a passed-in email

    Args:
      gmailService: Gmail service
      emailData: Email data to get photos from
    Returns:
      Void, downloads attached photos to subdirectory downloadedGmailPhotos
    """

    msgId = emailData['id']

    counter = 0
    for part in emailData['payload']['parts']:
        if '.jpg' in part['filename']:
            attachInfo = gmailService.users().messages().attachments().get(userId='me', messageId=msgId, id=part['body']['attachmentId']).execute()

            attachData = attachInfo['data']
            file_data = base64.urlsafe_b64decode(attachData.encode('UTF-8'))

            path = './downloadedGmailPhotos/' + str(counter) + '.jpg'
            with open(path, 'wb') as f:
                f.write(file_data)
            counter = counter + 1

async def uploadPhotosFromDownloaded(discordClient, loadedConfig):
    """Uploads downloaded files from downloadPhotosFromEmail to Discord

    Args:
      discordClient: Discord client
      loadedConfig: configuration from yaml file

    Returns:
      Void, uploads images to Discord, and then deletes the file it uploaded from downloadedGmailPhotos
    """
    discordChannel = discordClient.get_channel(loadedConfig['discord']['msgChannelId'])
    # Files to upload
    filePath = './downloadedGmailPhotos'
    imagesToUpload = [f for f in listdir(filePath) if isfile(join(filePath, f))]
    for fileVal in imagesToUpload:
        await discordChannel.send(file=discord.File(filePath + '/' + fileVal))
        remove(filePath + '/' + fileVal)

async def sendTextFromEmail(emailData, discordClient, loadedConfig):
    """Uploads the text contained in the email (versus the attachments)

    Args:
      emailData: Email data to get text from
      discordClient: Discord client
      loadedConfig: configuration from yaml file

    Returns:
      Void, uploads text from email to Discord
    """

    for part in emailData['payload']['parts']:
        if part['mimeType'] == 'text/plain':
            base64text = part['body']['data']
            # https://gist.github.com/perrygeo/ee7c65bb1541ff6ac770
            # apparently there's no padding in certain cases, so we need to add it in
            # we can add arbitrary padding length in, it'll ignore the rest of it.
            toDecode = base64text + '======================='

            messageText = ''
            try:
                # Parse body text
                messageText = base64.b64decode(toDecode).decode('utf-8')
            except:
                e = sys.exc_info()[1]
                print(e)

            discordChannel = discordClient.get_channel(loadedConfig['discord']['msgChannelId'])
            await discordChannel.send(messageText)

def getGmailVideoLabel(gmailService, loadedConfig):
    """Gets a Gmail label id, given the gmail service and string value defined in the config

    Args:
      gmailService: Gmail service
      loadedConfig: configuration from yaml file
    Returns:
      Numerical string for the Gmail label id
    """

    labelId = ''
    results = gmailService.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])
    if labels:
        labelId = ''
        for label in labels:
            if label['name'] == loadedConfig['gmail']['videoLabelName']:
                labelId = label['id']
                break
    return labelId

async def sendGmailAsDiscord(labelId, discordClient, gmailService, loadedConfig):
    """Main function to convert all unread Gmail emails under a certain label (defined in configs) to a series of Discord messages
    Args:
      labelId: Gmail label id
      discordClient: Discord client
      gmailService: Gmail service
      loadedConfig: configuration from yaml file
    Returns:
      Void, Sens email photos and message content as Discord messages
    """

    emails = getUnreadEmails(labelId, gmailService)

    for email in emails:
        try:
            emailData = gmailService.users().messages().get(userId='me', id=email['id']).execute()

            downloadPhotosFromEmail(gmailService, emailData)
            await uploadPhotosFromDownloaded(discordClient, loadedConfig)
            await sendTextFromEmail(emailData, discordClient, loadedConfig)

            labelToRemove = {
                "removeLabelIds": [
                    'UNREAD'
                ],
                "addLabelIds": []
            }
            try:
                # Remove unread label
                gmailService.users().messages().modify(userId='me', id=emailData['id'], body=labelToRemove).execute()
            except:
                pass

        except Exception as e:
            print(e)
            print('exception sendGmailAsDiscord')
            continue

# Discord client class
class MyClient(discord.Client):
    videoLabelId = None
    gmailService = None
    loadedConfig = None

    async def on_ready(self):
        """On_Ready function, when the discord bot is up and running

        Args:
          self: self
        Returns:
          Void, logs a success message
        """

        print('Logged on as {0}!'.format(self.user))

    async def on_message(self, message):
        """On_Message function, fired whenever a discord message is noticed in the server

        Args:
          self: self
          message: Message text
        Returns:
          Void, logs a message and handles message text appropriately (if applicable)
        """

        print('Message from {0.author}: {0.content}'.format(message))

        if message.content == '!checkEmail':
            await sendGmailAsDiscord(videoLabelId, self, gmailService, loadedConfig)
            print('Check Email Manually triggered')

# Generic helper functions
async def do_stuff_every_x_seconds(timeout, stuff, *args):
    """Helper function for calling an async function every timeout seconds

    Args:
      timeout: Seconds to call "stuff"
      stuff: Name of async function to call
      *args: Arguments to pass into the "stuff" function
    Returns:
      Void, calls stuff with args every "timeout" seconds as part of the asyncio event loop
    """

    while True:
        await asyncio.sleep(timeout)
        await stuff(*args)

# Pulling in config parameters
loadedConfig = yaml.safe_load(open("./piDiscordConfig.yaml"))

# Gmail initialization
gmailService = gmailInitialize(loadedConfig)
videoLabelId = getGmailVideoLabel(gmailService, loadedConfig)

# Discord client initialization
client = MyClient()
if len(videoLabelId):
    client.videoLabelId = videoLabelId
    client.gmailService = gmailService
    client.loadedConfig = loadedConfig

    # Start checking every 5 seconds to send Gmail camera emails as discord messages
    # Added into the discord client event loop
    client.loop.create_task(do_stuff_every_x_seconds(300, sendGmailAsDiscord, videoLabelId, client, gmailService, loadedConfig))
# Starting the discord bot
client.run(loadedConfig['discord']['clientToken'])