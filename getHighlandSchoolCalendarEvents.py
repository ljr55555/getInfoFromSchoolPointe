import requests
from urllib.parse import urlparse
import requests.adapters

import re
from datetime import datetime, timedelta
import dateutil.parser as parser
import pytz

# Used to decrypt password from config file
from cryptography.fernet import Fernet
from base64 import b64encode, b64decode

# Requirements for ExchangeLib
from exchangelib import DELEGATE, IMPERSONATION, Account, Credentials, FaultTolerance, \
    Configuration, NTLM, GSSAPI, SSPI, Build, Version, CalendarItem, EWSDateTime

from exchangelib.folders import Calendar
from exchangelib.items import MeetingRequest, MeetingCancellation, SEND_TO_ALL_AND_SAVE_COPY
from exchangelib.protocol import BaseProtocol

# User-specific config
from config import strUsernameCrypted, strPasswordCrypted, strEWSHost, strPrimarySMTP, iMaxExchangeResults, strKey, listWantedCategories

# Set our time zone
local_tz = pytz.timezone('US/Eastern')

# This needs to be the key used to stash the username and password values stored in config.py
f = Fernet(strKey)
strUsername = f.decrypt(strUsernameCrypted)
strUsername = strUsername.decode("utf-8")

strPassword = f.decrypt(strPasswordCrypted)
strPassword = strPassword.decode("utf-8")



def createExchangeItem(objExchangeAccount, strTitle, strLocn, strStartDate, strEndDate, strInviteeSMTP=None):
    print("Creating item {} which starts on {} and ends at {}".format(strTitle, strStartDate,strEndDate))
    objStartDate = parser.parse(strStartDate)
    objEndDate = parser.parse(strEndDate)

    if strInviteeSMTP is None:
        item = CalendarItem(
            account=objExchangeAccount,
            folder=objExchangeAccount.calendar,
            start=objExchangeAccount.default_timezone.localize(EWSDateTime(objStartDate.year, objStartDate.month, objStartDate.day, objStartDate.hour, objStartDate.minute)),
            end=objExchangeAccount.default_timezone.localize(EWSDateTime(objEndDate.year,objEndDate.month,objEndDate.day,objEndDate.hour,objEndDate.minute)),
            subject=strTitle,
            reminder_minutes_before_start=1440,
            reminder_is_set=True,
            location=strLocn,
            body=""
        )
    else:
        item = CalendarItem(
            account=objExchangeAccount,
            folder=objExchangeAccount.calendar,
            start=objExchangeAccount.default_timezone.localize(EWSDateTime(objStartDate.year, objStartDate.month, objStartDate.day, objStartDate.hour, objStartDate.minute)),
            end=objExchangeAccount.default_timezone.localize(EWSDateTime(objEndDate.year,objEndDate.month,objEndDate.day,objEndDate.hour,objEndDate.minute)),
            subject=strTitle,
            reminder_minutes_before_start=1440,
            reminder_is_set=True,
            location=strLocn,
            body="",
            required_attendees=[strInviteeSMTP]
        )
    item.save(send_meeting_invitations=SEND_TO_ALL_AND_SAVE_COPY )


def utc_to_local(utc_dt):
    local_dt = utc_dt.replace(tzinfo=pytz.utc).astimezone(local_tz)
    return local_tz.normalize(local_dt)



def main():
    class RootCAAdapter(requests.adapters.HTTPAdapter):
        # An HTTP adapter that uses a custom root CA certificate at a hard coded location
        def cert_verify(self, conn, url, verify, cert):
            cert_file = {
                'exchange01.rushworth.us': './ca.crt'
            }[urlparse(url).hostname]
            super(RootCAAdapter, self).cert_verify(conn=conn, url=url, verify=cert_file, cert=cert)

    #Use this SSL adapter class instead of the default
    BaseProtocol.HTTP_ADAPTER_CLS = RootCAAdapter

    # Get Exchange calendar events and save to dictEvents
    dictEvents = {}

    credentials = Credentials(username=strUsername, password=strPassword)
    config = Configuration(server=strEWSHost, credentials=credentials)
    account = Account(primary_smtp_address=strPrimarySMTP, config=config,
                    autodiscover=False, access_type=DELEGATE)

    for item in account.calendar.all().order_by('-start')[:iMaxExchangeResults]:
        if item.start:
            objEventStartTime = parser.parse(str(item.start))
            objEventStartTime = utc_to_local(objEventStartTime)

            strEventKey = "{}{:02d}{:02d}{:02d}".format(str(item.subject), int(objEventStartTime.year), int(objEventStartTime.month), int(objEventStartTime.day))
            strEventKey = strEventKey.replace(" ","")
            dictEvents[strEventKey]=1
            print(f"I added {strEventKey} to the dictionary")

    # Get next 30 days of data from the ICAL feed
    strFirstDay = datetime.today().strftime('%m/%d/%Y')

    dateLastDay = datetime.today() + timedelta(days=31)
    strLastDay = dateLastDay.strftime('%m/%d/%Y')

    strBaseURL = f'http://www.highlandschools.org/ical.ashx?e=&s=0&t=&sd={strFirstDay}&ed={strLastDay}&wk=&n=Events&v=&l=&dee=true'
    iTimeout = 600

    strHeader={'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36'}

    page = requests.get(strBaseURL, timeout=iTimeout, headers=strHeader)

    strContent = page.content
    strContent = strContent.decode("utf8")
    
    strRSSItem = re.findall('BEGIN:VEVENT(.*?)END:VEVENT', strContent, re.DOTALL)
    for strRSSRecord in strRSSItem:
        strStartTime = re.search('DTSTART:(.*?)\r\n', strRSSRecord, re.DOTALL)
        strEndTime = re.search('DTEND:(.*?)\r\n', strRSSRecord, re.DOTALL)
        strEventName =  re.search('SUMMARY:(.*?)\r\n', strRSSRecord, re.DOTALL)
        strEventCategory = re.search('TYPENAME:(.*?)\r\n', strRSSRecord, re.DOTALL)

        if strEventCategory[1] in listWantedCategories or strEventName[1] == 'BOE Meeting':
            try:
                # Test if start date exists
                strStartTime[1]
                dateStart = strStartTime[1]
                dateEnd = strEndTime[1]

            except:
                # Record is all-day event
                strStartTime = re.search('DTSTART;VALUE=DATE:(.*?)\r\n', strRSSRecord, re.DOTALL)
                strEndTime = re.search('DTEND;VALUE=DATE:(.*?)\r\n', strRSSRecord, re.DOTALL)
                dateStart = f"{strStartTime[1]}T000000"
                dateEnd = f"{strEndTime[1]}T000000"

            # Changed meeting title prefix for CalDav filtering
            #if strEventCategory[1] is None:
            #    strSummary = "Highland Schools: {}".format(strEventName[1])
            #else:
            #    strSummary = "{}: {}".format(strEventCategory[1], strEventName[1])
            strSummary = f"Highland Schools: {strEventName[1]}"
            strSummary = strSummary.replace("HE - ","Hinckley Elementary ")
            strSummary = strSummary.replace("BOE Meeting","Board of Education Meeting")

            strThisEventKey = strSummary + (str(dateStart).split('T'))[0]
            strThisEventKey = strThisEventKey.replace(" ","")

            print(f"Checking for {strThisEventKey} in dictEvents")
            if strThisEventKey not in dictEvents:
                #print("The event {} on {} does not exist in the calendar and would be created.".format(strThisEventKey, str(dateStart)))
                createExchangeItem(account, strSummary, strEventCategory[1], dateStart, dateEnd)
            else:
                print("The event {} on {} already exists in the calendar.".format(strThisEventKey, str(dateStart)))

if __name__ == '__main__':
    main()

