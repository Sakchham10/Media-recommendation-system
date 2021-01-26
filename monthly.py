import gspread
import random
import smtplib,ssl 
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from time import strftime

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
    info_table= {}
    key_list = table['User']
    flag = 1
    for key in key_list:
        info_table[key] = table['Email'][flag-1:flag]
        flag+=1
    return info_table

def media_generator(actual_table):
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
    return rec

def send_email(info_table,rec):
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
    text = "Your game for {} is: ".format(month)
    for key in info_table.keys():
        message = MIMEMultipart("alternative")
        message["Subject"] = "Game of the Month"
        message["From"] = sender_email
        receiver_email = info_table[key]
        message["To"] = receiver_email[0]
        final_message = text + rec[key]
        part1 = MIMEText(final_message, "plain")
        message.attach(part1)
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(smtp_server,port,context=context) as server:
            server.login(sender_email,password)
            server.sendmail(sender_email,receiver_email,message.as_string())

game_sheet = access_sheets("Media Club",1)
game_table = table_generator(game_sheet)
games = media_generator(game_table)
user_sheet = access_sheets("Media Club",2)
user_table = table_generator(user_sheet)
emails= email_generator(user_table)
send_email(emails,games)
