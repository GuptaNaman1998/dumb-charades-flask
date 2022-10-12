"""
"""

from flask import Flask, render_template
import datetime as dt
import win32com.client as win32
import pythoncom
import random

app = Flask(__name__)

file = open("movies.txt","r")
Qs = file.read().split('\n')

email_mapping = {"Naman":"abcd@domain.com","Palak":"abcd@domain.com","Raja":"abcd@domain.com","Bhashkar":"abcd@domain.com","Pronoy":"abcd@domain.com","Bharat":"abcd@domain.com","Nimma":"abcd@domain.com","Fahima":"abcd@domain.com","Amrita":"abcd@domain.com","Vikash":"abcd@domain.com","Renu":"abcd@domain.com","Namrata":"abcd@domain.com"}

def sending_email(Email_address,User_name,report):
    pythoncom.CoInitialize()
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.Subject = 'You Just Got Picked to Enact Movie Enclosed Below... Happy Charading!'
    # mail.Subject = 'Test email... Happy Weekend!'
    User_name += ",<br><br>"
    mail.To = Email_address 
    mail.HTMLBody = f"""
    Dear {User_name}
    <br><br><center><h1>{report}</h1></center><br><br>
    The one who is acting needs to be silent, as the word ‘dumb’ in the name of the game suggests.<br>
    The player will have to use facial expressions, gestures, and even body language.<br>
    Things like lip-reading, humming tunes, etc. are banned.<br>
    There is a set time limit for each team to guess.<br><br>
    Best regards,<br>
    Naman Gupta.
    """
    mail.Send()
    return True

@app.route("/boss2")
def question():# df):
    player_name = "boss or manager or team lead"
    sending_email("abcd@domain.com",player_name,report="<some difficult movie>")
    return '<br><br><br><br><br><br><br><br><br><br><br><center><h1>{}</h1></center>'.format(player_name)

@app.route("/player")
def player():# df):
    list_names = list(email_mapping.keys())
    player_name = random.choice(list_names)
    choose = random.choice(Qs)
    Qs.remove(choose)
    del email_mapping[player_name]
    sending_email(email_mapping[player_name],player_name,report=choose)
    return '<br><br><br><br><br><br><br><br><br><br><br><center><h1>{}</h1></center>'.format(player_name)

@app.route("/boss1")
def app_name():
    player_name = "boss or manager or team lead"
    sending_email("abcd@domain.com",player_name,report="<some difficult movie>")
    return '<br><br><br><br><br><br><br><br><br><br><br><center><h1>{}</h1></center>'.format(player_name)
    
@app.route("/test")
def special():
    choose = random.choice(Qs)
    Qs.remove(choose)
    return '<br><br><br><br><br><br><br><br><br><br><br><center><h1>{}</h1></center>'.format(choose)
    
if __name__ == "__main__":
    app.run(debug=True)
