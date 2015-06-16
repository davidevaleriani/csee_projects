import cherrypy
import sqlite3
import hashlib
import os
import string
import time
import random
import smtplib
from jinja2 import Environment, FileSystemLoader
from copy import deepcopy as copy
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import datetime
from score import get_score

db_name = "users.db"
base_site = "http://cseepgdp2.essex.ac.uk/ceec-poker/"
#base_site = "http://127.0.0.1:8081/ceec-poker/"
site_port = 8081

env = Environment(loader=FileSystemLoader('templates'))
menu = [{"link": base_site+"home", "caption": "Home", "active": False},
        {"link": base_site+"data", "caption": "Get the data", "active": False},
        {"link": base_site+"rank", "caption": "Leaderboard", "active": False},
        {"link": base_site+"about", "caption": "About", "active": False},
        {"link": base_site+"login", "caption": "Login", "active": False},
        {"link": base_site+"signup", "caption": "Signup", "active": False},
        ]
menu_logged = [{"link": base_site+"home", "caption": "Home", "active": False},
        {"link": base_site+"data", "caption": "Get the data", "active": False},
        {"link": base_site+"rank", "caption": "Leaderboard", "active": False},
        {"link": base_site+"about", "caption": "About", "active": False},
        {"link": base_site+"submit", "caption": "Submit", "active": False},
        {"link": base_site+"logout", "caption": "Logout", "active": False},
        ]


def datetimeformat(value):
    return datetime.datetime.strptime(value, '%Y-%m-%d %H:%M:%S').strftime('%a, %d %b %Y %H:%M:%S')
env.filters['datetimeformat'] = datetimeformat


def is_logged():
    if "id" in cherrypy.session:
        return True
    return False


class Index(object):
    @cherrypy.expose
    def index(self):
        print("REDIRECT")
        raise cherrypy.HTTPRedirect("home/")


class HomePage(object):
    @cherrypy.expose
    def index(self):
        template = env.get_template("template.html")
        if is_logged():
            current_menu = copy(menu_logged)
        else:
            current_menu = copy(menu)
        current_menu[0]["active"] = True
        content = [{"title": "Description", "text": (''.join(open("index.html").readlines())).decode('utf-8')},
                   {"title": "Getting started", "text": "<h3><a href='../data/'>Get the data</a> &#8594; "
                                                        "<a href='../signup/'>Signup</a> &#8594; "
                                                        "<a href='../submit/'>Submit</a></h3>"}]
        return template.render(navigation=current_menu, content=content)


class About(object):
    @cherrypy.expose
    def index(self):
        template = env.get_template("template.html")
        if is_logged():
            current_menu = copy(menu_logged)
        else:
            current_menu = copy(menu)
        current_menu[3]["active"] = True
        content = [{"title": "Organizers", "text": (''.join(open("about.html").readlines())).decode('utf-8')}]
        return template.render(navigation=current_menu, content=content)


class Signup(object):
    @cherrypy.expose
    def index(self):
        template = env.get_template("template.html")
        if is_logged():
            current_menu = copy(menu_logged)
        else:
            current_menu = copy(menu)
            current_menu[5]["active"] = True
        content = [{"title": "Signup", "text": (''.join(open("signup.html").readlines())).decode('utf-8')}]
        return template.render(navigation=current_menu, content=content)

    @cherrypy.expose
    def register(self, username, password2, country, affiliation, name2, name1, password, email):
        if password != password2:
            template = env.get_template("template.html")
            current_menu = copy(menu)
            content = [{"title": "Error", "text": "<p style='color:red'>The password are not the same</p>"}]
            return template.render(navigation=current_menu, content=content)
        password = hashlib.sha1(password).hexdigest()
        conn = sqlite3.connect(db_name)
        c = conn.cursor()
        c.execute("SELECT * FROM users WHERE username = ?", [username])
        user = c.fetchone()
        if user is not None:
            conn.close()
            template = env.get_template("template.html")
            current_menu = copy(menu)
            content = [{"title": "Error", "text": "<p style='color:red'>Your username has already been taken or you.</p>"}]
            return template.render(navigation=current_menu, content=content)
        c.execute("SELECT * FROM users WHERE email = ?", [email])
        user = c.fetchone()
        if user is not None:
            conn.close()
            template = env.get_template("template.html")
            current_menu = copy(menu)
            content = [{"title": "Error", "text": "<p style='color:red'>Your email address is already registered. Please login.</p>"}]
            return template.render(navigation=current_menu, content=content)
        # Send the confirmation email
        from_address = "competition@ceec.uk"
        link = base_site+"/activate?user="+username
        msg = MIMEMultipart('alternative')
        msg['Subject'] = "CEEC Competition Registration"
        msg['From'] = "CEEC Competition<"+from_address+">"
        msg['To'] = email
        # Create the body of the message (a plain-text and an HTML version).
        text = "Hi "+name1+" "+name2+"!\nThank you for registering to the CEEC 2015 Poker Expected Hand Strength Generalization Competition.\n" \
                                     "Please click on this link to complete your registration: "+link+"\n" \
                                     "Best regards\n" \
                                     "CEEC 2015 Programme Committee"

        html = "<html><body><p>Hi "+name1+" "+name2+"<br>" \
            "Thank you for registering to the CEEC 2015 Poker Expected Hand Strength Generalization Competition.</p>" \
            "<p>Please click on this link to complete your registration <a href='"+link+"'>"+link+"</a></p>" \
            "<p>Best regards<br>CEEC 2015 Programme Committee</p></body></html>"
        # Record the MIME types of both parts - text/plain and text/html.
        msg.attach(MIMEText(text, 'plain'))
        msg.attach(MIMEText(html, 'html'))
        # Send the message
        server = smtplib.SMTP("smtp.123-reg.co.uk:587")
        server.login("competition@ceec.uk", "cherryPy123__")
        server.sendmail(from_address, [email], msg.as_string())
        server.quit()
        # Add the user to the database
        c.execute("INSERT INTO users(name1,name2,affiliation,country,email,username,password,active) VALUES(?, ?, ?, ?, ?, ?, ?, 0)",
                  [name1, name2, affiliation, country, email, username, password])
        conn.commit()
        conn.close()
        # Build the page to show
        template = env.get_template("template.html")
        current_menu = copy(menu)
        current_menu[5]["active"] = True
        content = [{"title": "Account created!",
                    "text": "Check your email to activate your account and start competing!"}]
        return template.render(navigation=current_menu, content=content)


class Activate(object):
    @cherrypy.expose
    def index(self, user):
        conn = sqlite3.connect(db_name)
        c = conn.cursor()
        c.execute("SELECT * FROM users WHERE username = ?", [user])
        res = c.fetchone()
        if res is None:
            template = env.get_template("template.html")
            current_menu = copy(menu)
            content = [{"title": "Error", "text": "<p style='color:red'>Link not valid</p>"}]
            return template.render(navigation=current_menu, content=content)
        c.execute("SELECT * FROM users WHERE username = ? AND active=1", [user])
        res = c.fetchone()
        if res is not None:
            template = env.get_template("template.html")
            current_menu = copy(menu)
            content = [{"title": "Error", "text": "<p style='color:red'>Your account has already been activated. Please <a href='../login'>login</a>.</p>"}]
            return template.render(navigation=current_menu, content=content)
        c.execute("UPDATE users SET active=1 WHERE username = ?", [user])
        conn.commit()
        conn.close()
        template = env.get_template("template.html")
        current_menu = copy(menu)
        content = [{"title": "Account activated!", "text": "You can now <a href='../login'>login</a> and start submitting.</p>"}]
        return template.render(navigation=current_menu, content=content)


class Login(object):
    @cherrypy.expose
    def index(self):
        template = env.get_template("template.html")
        if is_logged():
            current_menu = copy(menu_logged)
        else:
            current_menu = copy(menu)
            current_menu[4]["active"] = True
        content = [{"title": "Login", "text": (''.join(open("login.html").readlines())).decode('utf-8')}]
        return template.render(navigation=current_menu, content=content)

    @cherrypy.expose
    def authenticate(self, username, password):
        conn = sqlite3.connect(db_name)
        c = conn.cursor()
        c.execute("SELECT * FROM users WHERE username=? AND password=?",
                  [username, hashlib.sha1(password).hexdigest()])
        user = c.fetchone()
        if user is None:
            template = env.get_template("template.html")
            current_menu = copy(menu)
            content = [{"title": "Error", "text": (''.join(open("error.html").readlines())).decode('utf-8')}]
            return template.render(navigation=current_menu, content=content)
        if "id" not in cherrypy.session:
            cherrypy.session["id"] = user[0]
            cherrypy.session["name"] = user[1]+" "+user[2]
            cherrypy.session["username"] = user[5]
        raise cherrypy.HTTPRedirect("../submit")


class GetNewPassword(object):
    @cherrypy.expose
    def index(self):
        template = env.get_template("template.html")
        if is_logged():
            current_menu = copy(menu_logged)
        else:
            current_menu = copy(menu)
        content = [{"title": "Password recovery", "text": (''.join(open("get_new_password.html").readlines())).decode('utf-8')}]
        return template.render(navigation=current_menu, content=content)

    @cherrypy.expose
    def get_passwd(self, email):
        conn = sqlite3.connect(db_name)
        c = conn.cursor()
        c.execute("SELECT * FROM users WHERE email = ?", [email])
        res = c.fetchone()
        if res is None:
            template = env.get_template("template.html")
            current_menu = copy(menu)
            content = [{"title": "Email not valid!", "text": "<p style='color:red'>The email you have provided does not exist in our databases.</p>"}]
            return template.render(navigation=current_menu, content=content)
        new_pass = ''.join(random.SystemRandom().choice(string.ascii_uppercase + string.ascii_lowercase + string.digits) for _ in range(10))
        c.execute("UPDATE users SET password = ? WHERE email = ?", [hashlib.sha1(new_pass).hexdigest(), email])
        conn.commit()
        conn.close()
        name = res[1] + " " + res[2]
        username = res[6]
        # Send the confirmation email
        from_address = "competition@ceec.uk"
        msg = MIMEMultipart('alternative')
        msg['Subject'] = "CEEC Password Recovery"
        msg['From'] = "CEEC Competition<"+from_address+">"
        msg['To'] = email
        # Create the body of the message (a plain-text and an HTML version).
        text = "Hi "+name+"!\nYour new login details are:\n" \
                                     "Username: "+username+"\n" \
                                     "Password: "+new_pass+"\n" \
                                     "Best regards\n" \
                                     "CEEC 2015 Programme Committee"

        html = "<html><body><p>Hi "+name+"<br>" \
            "Your new login details are:</p>" \
            "<p>Username: <strong>"+username+"</strong><br>" \
            "Password: <strong>"+new_pass+"</strong></p>" \
            "<p>Best regards<br>CEEC 2015 Programme Committee</p></body></html>"
        # Record the MIME types of both parts - text/plain and text/html.
        msg.attach(MIMEText(text, 'plain'))
        msg.attach(MIMEText(html, 'html'))
        # Send the message
        server = smtplib.SMTP("smtp.123-reg.co.uk:587")
        server.login("competition@ceec.uk", "cherryPy123__")
        server.sendmail(from_address, [email], msg.as_string())
        server.quit()
        # Success page
        template = env.get_template("template.html")
        current_menu = copy(menu)
        content = [{"title": "Password reset!", "text": "<p>Your password has been reset and your new login details have been emailed to you.</p>"}]
        return template.render(navigation=current_menu, content=content)


class Data(object):
    @cherrypy.expose
    def index(self):
        template = env.get_template("template.html")
        if is_logged():
            current_menu = copy(menu_logged)
        else:
            current_menu = copy(menu)
        current_menu[1]["active"] = True
        content = [{"title": "Data", "text": (''.join(open("data.html").readlines())).decode('utf-8')}]
        return template.render(navigation=current_menu, content=content)


class Logout(object):
    @cherrypy.expose
    def index(self):
        cherrypy.session.regenerate()
        raise cherrypy.HTTPRedirect("/")


class Submit(object):
    @cherrypy.expose
    def index(self):
        if "id" in cherrypy.session:
            template = env.get_template("userdata.html")
            current_menu = copy(menu_logged)
            current_menu[4]["active"] = True
            content = [{"title": "Submit an entry", "text": (''.join(open("submit.html").readlines())).decode('utf-8')}]
            user = {"name": cherrypy.session["name"]}
            return template.render(navigation=current_menu, content=content, user=user)
        else:
            raise cherrypy.HTTPRedirect("../login")

    @cherrypy.expose
    def make_submission(self, entry1000, entry10000, entry20000):
        if entry1000.filename == "" or entry10000.filename == "" or entry20000.filename == "":
            template = env.get_template("template.html")
            current_menu = copy(menu_logged)
            current_menu[4]["active"] = True
            content = [{"title": "File missing!", "text": "<p style='color:red'>Please upload the three requested CSV files.</p>"}]
            return template.render(navigation=current_menu, content=content)
        if entry1000.filename[-3:].upper() != "CSV" or entry10000.filename[-3:].upper() != "CSV" or entry20000.filename[-3:].upper() != "CSV":
            template = env.get_template("template.html")
            current_menu = copy(menu_logged)
            current_menu[4]["active"] = True
            content = [{"title": "File not valid!", "text": "<p style='color:red'>One of the files you are trying to upload is not a CSV file.</p>"}]
            return template.render(navigation=current_menu, content=content)
        # Save the files in the submission directory
        date = time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())
        submission_dir = "submissions/"+date.replace(" ", "_")+"_"+str(cherrypy.session["id"])+"/"
        if not os.path.exists(submission_dir):
            os.makedirs(submission_dir)
        filename = submission_dir+"sample_1000.csv"
        output = open(filename, "w")
        output.write(entry1000.file.read())
        output.close()
        filename = submission_dir+"sample_10000.csv"
        output = open(filename, "w")
        output.write(entry10000.file.read())
        output.close()
        filename = submission_dir+"sample_20000.csv"
        output = open(filename, "w")
        output.write(entry20000.file.read())
        output.close()

        # Get the score
        score = get_score(submission_dir)

        # Store the score in the database
        conn = sqlite3.connect(db_name)
        c = conn.cursor()
        c.execute("INSERT INTO submissions(author, date, score) VALUES(?, ?, ?)",
                  [cherrypy.session["id"], date, score])
        conn.commit()
        conn.close()
        score_text = "This submission has achieved a score of <strong>"+str(round(score, 5))+"</strong> on the validation set." \
                                                                                   "<br><a class='btn btn-success' href='../rank'>Check the rank</a>"
        # Build the page to show
        template = env.get_template("userdata.html")
        current_menu = copy(menu_logged)
        current_menu[4]["active"] = True
        content = [{"title": "Submitted!",
                    "text": (''.join(open("submission.html").readlines())).decode('utf-8')},
                   {"title": "Score",
                    "text": score_text}]
        user = {"name": cherrypy.session["name"]}
        return template.render(navigation=current_menu, content=content, user=user)


class Rank(object):
    @cherrypy.expose
    def index(self):
        conn = sqlite3.connect(db_name)
        c = conn.cursor()
        c.execute("SELECT u.username, s.date, s.score FROM users as u, submissions as s WHERE u.id = s.author ORDER BY s.score DESC, s.date DESC")
        res = c.fetchall()
        template = env.get_template("template_rank.html")
        if is_logged():
            current_menu = copy(menu_logged)
        else:
            current_menu = copy(menu)
        current_menu[2]["active"] = True
        return template.render(navigation=current_menu, submissions=res)


def get_users():
    with sqlite3.connect(db_name) as c:
        cursor = c.cursor()
        cursor.execute("SELECT username, password FROM users")
        return dict(cursor.fetchall())

if __name__ == '__main__':
    cherrypy.server.socket_host = '0.0.0.0'
    cherrypy.server.socket_port = site_port
    users = {"admin": "secretPassword", "editor": "otherPassword"}
    conf = {'/': {'tools.sessions.on': True,
                  'tools.staticdir.on': True,
                  'tools.staticdir.dir': os.path.abspath(os.getcwd())
                  },
            }
    root = Index()
    root.home = HomePage()
    root.data = Data()
    root.about = About()
    root.login = Login()
    root.get_new_password = GetNewPassword()
    root.logout = Logout()
    root.signup = Signup()
    root.submit = Submit()
    root.rank = Rank()
    root.activate = Activate()
    cherrypy.quickstart(root, '/ceec-poker', config=conf)