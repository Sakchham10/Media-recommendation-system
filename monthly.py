import gspread
import random
import smtplib,ssl 
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from time import strftime

BOOK_SHEET = 0
GAME_SHEET = 1

def access_sheets(name,sheet_num):
    """
    Take in  name and sheet_num and return a editable data sheet
    """
    gc = gspread.service_account()
    sh = gc.open(name)
    table = sh.get_worksheet(sheet_num)
    return table

def table_generator(table):
    """
    Take in an editable data sheet and return a dict like: 
    {Column name: Column values,...}
    """
    keys=[]
    actual_table = {}
    headers = table.row_values(1)
    for x in range (len(headers)):
        keys.append(headers[x])
        value_list = table.col_values(x+1)
        actual_table[keys[x]] = value_list[1:]
    return actual_table

def email_generator(table):
    """
    Take in the dict of the data sheet and return a dict like:
    {User:Email}
    Note: Used only for user info sheet
    """
    user_info_table= {}
    key_list = table['User']
    flag = 1
    for key in key_list:
        user_info_table[key] = table['Email'][flag-1:flag]
        flag+=1
    return user_info_table

def media_generator(actual_table,media_type):
    """
    Take in the dict from table_generator and select random game for every user.
    Return a dict like {User: Game,...} 
    """
    rec = {}
    for key in actual_table.keys():
        media_list = actual_table[key]
        # Filter out empty values
        media_list = [media for media in media_list if media != '']
        #Remove duplicates
        media_list = list(set(media_list))
        rec[key] = random.choice(media_list)
    rec["type"] = media_type
    return rec

def send_email(user_info_table,media_table):
    """
    Take in the dict from email_generator and media_generator.
    Send out game recommendations to each user.
    """
    smtp_server = "smtp.gmail.com"
    port = 465
    with open("credentials.txt") as inp:
        credentials = inp.readlines()
        sender_email = credentials[0]
        password = credentials[1]
    month = strftime("%B")
    text = "Your reccomendations for {} are: ".format(month)
    for user in user_info_table.keys():
        message = MIMEMultipart("alternative")
        message["Subject"] = "Media of the Month"
        message["From"] = sender_email
        receiver_email = user_info_table[user]
        message["To"] = receiver_email[0]
        print("media_table: ",media_table, "user: ", user)
        user_rec = media_table[user]
        final_message = " "
        for media in user_rec.keys():
            rec = user_rec[media]
            final_message = final_message + " " + rec
        final_message = text + final_message
        part1 = MIMEText(final_message, "plain")
        message.attach(part1)
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(smtp_server,port,context=context) as server:
            server.login(sender_email,password)
            server.sendmail(sender_email,receiver_email,message.as_string())


def generate_rec(sheet,media_type):
    media_sheet = access_sheets("Media Club",sheet)
    media_table = table_generator(media_sheet)
    media = media_generator(media_table,media_type)
    return media


# {Nikhil: {Books: Illiad, Games: Metro}} 

def unify_rec(*args):
    media={}
    users = [user for user in args[0].keys() if user != "type"]
    for user in users:
        media[user]={}
    for arg in args:
        media_type = arg["type"]
        for user in users:
            media[user][media_type] = arg[user]
    return media
    
books = generate_rec(BOOK_SHEET,"Books")
games = generate_rec(GAME_SHEET,"Games")
media = unify_rec(books,games) 
user_sheet = access_sheets("Media Club",2)
user_table = table_generator(user_sheet)
emails= email_generator(user_table)
send_email(emails,media)
