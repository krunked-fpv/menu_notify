"""Alex Goresch 04/02/2019 rewritten 4/18

Notify subscribers if their are items being served today
that they have selected that they are interested in

using google forms & sheets for poor man front-end and database

sending emails via MAILGUN SMTP """

import requests
import os
import json
import gspread
import smtplib
import logging
import time
from oauth2client.service_account import ServiceAccountCredentials
from email.message import EmailMessage


def build_service():
    # build the a Google service

    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        'google_api_key.json',
        scope
    )
    service = gspread.authorize(credentials)

    return service


def get_menu_items(meal):
    """get the menu for all halls for that meal for that day return dict of halls \
    and items also check to see if the item has been historically recorded or not"""

    def add_to_history(meal, hall, item):
        # add to history_sheet if it wasn't found_in_history
        historic_sheet.append_row([meal, hall, item])
        log.info("\tRecord Added to historic_items(" +
                 meal + "," + hall + "," + item + ")")

    def found_in_history():
        # search history_items sheet for the unique meal/hall/item combo
        found = False
        for sublist in historic_items:
            if g_meals[meal] in sublist[0] \
                    and hall["Name"].strip() in sublist[1] \
                    and entry['FoodName'].strip() in sublist[2]:
                found = True
                break
        return found

    # get_menu_items starts here
    # remove the preceeding/trailing whitespace from items for comparison
    items = {}
    main_url = 'https://www.sais.wmich.edu/SA_Services/MenuPublic.svc/'

    log.info("Starting get_menu_items(" + g_meals[meal] + ")")

    for hall in json.loads(requests.get(main_url + 'Building').text):
        for entry in json.loads(requests.get(main_url + 'Menu/' + str(meal) + "/" + str(hall["ID"]) + "/" + time.strftime("%m-%d-%Y")).text):
            try:
                items[hall["Name"]].append(entry["FoodName"].strip())

                if not found_in_history():
                    add_to_history(
                        g_meals[meal], hall["Name"], entry["FoodName"].strip())

            except KeyError:
                items[hall["Name"]] = []  # set header
                items[hall["Name"]].append(entry['FoodName'].strip())

                if not found_in_history():
                    add_to_history(
                        g_meals[meal], hall["Name"], entry["FoodName"].strip())

    log.info("End get_menu_items(" + g_meals[meal] + ")")

    return items


def notify_subscriber(message):
    """email the subscriber the html formatted message"""

    log.info("Sending Email to " + subscriber.email)

    mailgun_username = os.environ.get('MAILGUN_USER')
    mailgun_password = os.environ.get('MAILGUN_PASSWORD')

    to = [subscriber.email]

    msg = EmailMessage()  # use this class
    msg['Subject'] = "Menu Hits! for " + time.strftime("%m-%d-%Y")
    msg['From'] = mailgun_username
    msg['To'] = ', '.join(to)
    msg.set_content(message)
    msg.add_alternative(message, subtype='html')  # sending as HTML

    try:
        with smtplib.SMTP('smtp.mailgun.org', 587) as server:
            server.login(mailgun_username, mailgun_password)
            server.send_message(msg)
            log.info("Email sent to " + subscriber.email + " successfully!")

    except Exception as error:
        log.error('Something went wrong...', error)


def check_menu(todays_menu, subscriber_items, subscriber_halls, meal):
    """check the subscribers items against todays items"""

    message = ''
    found = False

    for hall, item in todays_menu:
        for y in range(len(subscriber_items)):
            if subscriber_items[y] in item and hall in subscriber_halls:
                if not found:
                    message += "<dt><b>" + meal + "</b></dt>\n"
                message += "<dd>" + \
                    subscriber_items[y] + " @ " + hall + "</dd>\n"

                found = True
                subscriber.send = True

                log.info("\t" + meal + " @ " + hall +
                         " for " + subscriber_items[y])
    return message


class Person:
    """subscriber object"""

    def __init__(self, email, halls, breakfast, lunch, dinner, frequency):
        self.email = email
        self.halls = halls.split(', ')
        self.breakfast = breakfast.split(', ')
        self.lunch = lunch.split(', ')
        self.dinner = dinner.split(', ')
        self.frequency = frequency.split(', ')
        self.send = False

    def whoami(self):
        whoami = "Subscriber: " + self.email
        return whoami


def start_log():
    """automatically set the logname to scriptname.log"""
    log_format = "%(levelname)s %(asctime)s - %(message)s"
    logging.basicConfig(filename=os.path.basename(__file__)[:-3] + ".log",
                        level=logging.DEBUG,
                        format=log_format)
    log = logging.getLogger(__name__)

    return log


# MAIN STARTS HERE

log = start_log()
log.info(' PROGRAM STARTED '.center(36, '*'))

g_meals = {1: "Breakfast", 2: "Lunch", 3: "Dinner"}

gsheet = build_service()
wks = gsheet.open('menu_subscription')  # open this Google sheet

# gather the sheets into objects
historic_sheet = wks.worksheet("historic_items")
historic_items = historic_sheet.get_all_values()[1:]
subscriptions = wks.worksheet("Form Responses 5")

# get todays menu items for all halls breakfast,lunch,dinner
todays_breakfast = get_menu_items(1).items()
todays_lunch = get_menu_items(2).items()
todays_dinner = get_menu_items(3).items()

log.info("** Subscriber Processing Starting **")
sub_count = 0

# skip the header line of the response sheet
for sub in subscriptions.get_all_values()[1:]:

    # set the subscriber information
    subscriber = Person(sub[1], sub[2], sub[3], sub[4], sub[5], sub[6])
    sub_count += 1

    if time.strftime('%A') in subscriber.frequency:
        log.info(subscriber.whoami())

        # start the message html
        message = """\
<!DOCTYPE html>
<html>
<head></head>
    <body>
        <dl>\n"""

        # check subscription against todays menus
        message += check_menu(todays_breakfast,
                              subscriber.breakfast, subscriber.halls, g_meals[1])
        message += check_menu(todays_lunch, subscriber.lunch,
                              subscriber.halls, g_meals[2])
        message += check_menu(todays_dinner, subscriber.dinner,
                              subscriber.halls, g_meals[3])

        # close the message HTML
        message += """\
        <br><br>
        <footer><a href="https://alexgoresch.com">Click Here to get to the form and edit your responses</a></footer>
    </body>
</html>"""

        # notify the subscriber if they have any hits
        if subscriber.send:
            notify_subscriber(message)
        else:
            log.info("\tNo Hits for " + subscriber.whoami())

    else:
        log.info(subscriber.whoami() +
                 "\tIndicated to not be notified on " + time.strftime("%A"))

log.info("** Subscriber Processing Finished **")
log.info("Subscribers Processed: " + str(sub_count))
log.info(" PROGRAM ENDED ".center(36, '*') + "\n")
