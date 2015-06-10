import cherrypy
import sqlite3
import hashlib
import os
import time
import random
import smtplib
from jinja2 import Environment, FileSystemLoader
from copy import deepcopy as copy

db_name = "users.db"
base_site = "localhost:8080"

env = Environment(loader=FileSystemLoader('templates'))

menu = [{"link": "/", "caption": "Home", "active": False},
        {"link": "/data", "caption": "Get the data", "active": False},
        {"link": "/about", "caption": "About", "active": False},
        {"link": "/login", "caption": "Login", "active": False},
        {"link": "/signup", "caption": "Signup", "active": False},
        ]
menu_logged = [{"link": "/", "caption": "Home", "active": False},
        {"link": "/data", "caption": "Get the data", "active": False},
        {"link": "/about", "caption": "About", "active": False},
        {"link": "/submit", "caption": "Submit", "active": False},
        {"link": "/logout", "caption": "Logout", "active": False},
        ]


def is_logged():
    if "id" in cherrypy.session:
        return True
    return False


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
                   {"title": "Getting started", "text": "<h3><a href='data/'>Get the data</a> &#8594; "
                                                        "<a href='signup/'>Signup</a> &#8594; "
                                                        "<a href='submit/'>Submit</a></h3>"}]
        return template.render(navigation=current_menu, content=content)


class About(object):
    @cherrypy.expose
    def index(self):
        template = env.get_template("template.html")
        if is_logged():
            current_menu = copy(menu_logged)
        else:
            current_menu = copy(menu)
        current_menu[2]["active"] = True
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
            current_menu[4]["active"] = True
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
            template = env.get_template("template.html")
            current_menu = copy(menu)
            content = [{"title": "Error", "text": "<p style='color:red'>Your username has already been taken or you.</p>"}]
            return template.render(navigation=current_menu, content=content)
        c.execute("SELECT * FROM users WHERE email = ?", [email])
        user = c.fetchone()
        if user is not None:
            template = env.get_template("template.html")
            current_menu = copy(menu)
            content = [{"title": "Error", "text": "<p style='color:red'>Your email address is already registered. Please login.</p>"}]
            return template.render(navigation=current_menu, content=content)
        # Send the confirmation email
        # TODO Fix this
        enable = 1
        '''
        from_address = "competition@ceec.uk"
        server = smtplib.SMTP_SSL("smtp.123-reg.co.uk")
        server.login("competition@ceec.uk", "cherryPy123__")
        link = base_site+"/activate?token="+username
        message = "Hi "+name1+" "+name2+"<br>Thank you for registering to the CEEC 2015 Poker Expected Hand Strength Generalization Competition." \
                                        "Please click on this link to complete your registration <a href='"+link+"'>"+link+"</a>.<br><br>" \
                                        "CEEC15 Programme Committee"
        msg = "To: %s\r\nFrom: %s\r\nSubject: %s\r\nContent-type: text/html\r\n\r\n%s" % (email, from_address, "CEEC Competition Registration", message)
        server.sendmail(email, from_address, msg)
        server.quit()
        '''
        # Add the user to the database
        c.execute("INSERT INTO users(name1,name2,affiliation,country,email,username,password,active) VALUES(?, ?, ?, ?, ?, ?, ?, ?)",
                  [name1, name2, affiliation, country, email, username, password, enable])
        conn.commit()
        conn.close()
        # Build the page to show
        template = env.get_template("template.html")
        current_menu = copy(menu)
        current_menu[4]["active"] = True
        content = [{"title": "Account created!",
                    "text": "Check your email to activate your account and start competing!"}]
        return template.render(navigation=current_menu, content=content)


class Login(object):
    @cherrypy.expose
    def index(self):
        template = env.get_template("template.html")
        if is_logged():
            current_menu = copy(menu_logged)
        else:
            current_menu = copy(menu)
            current_menu[3]["active"] = True
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
        raise cherrypy.HTTPRedirect("/submit")


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
            current_menu[3]["active"] = True
            content = [{"title": "Submit an entry", "text": (''.join(open("submit.html").readlines())).decode('utf-8')}]
            user = {"name": cherrypy.session["name"]}
            return template.render(navigation=current_menu, content=content, user=user)
        else:
            raise cherrypy.HTTPRedirect("/login")

    @cherrypy.expose
    def make_submission(self, entry):
        # Save the file
        filename = str(cherrypy.session["id"])+"_"+time.strftime("%Y-%m-%d_%H:%M:%S", time.gmtime())+".txt"
        output = open("submissions/"+filename, "w")
        output.write(entry.file.read())
        output.close()
        # TODO Process the file to compute the score
        score = random.random()
        score_text = "This submission has achieved a score of <strong>"+str(score)+"</strong> on the validation set." \
                                                                                   "<br><a class='btn btn-success' href='/rank'>Check the rank</a>"
        # Build the page to show
        template = env.get_template("userdata.html")
        current_menu = copy(menu_logged)
        current_menu[3]["active"] = True
        content = [{"title": "Submitted!",
                    "text": (''.join(open("submission.html").readlines())).decode('utf-8')},
                   {"title": "Score",
                    "text": score_text}]
        user = {"name": cherrypy.session["name"]}
        return template.render(navigation=current_menu, content=content, user=user)


def get_users():
    with sqlite3.connect(db_name) as c:
        cursor = c.cursor()
        cursor.execute("SELECT username, password FROM users")
        return dict(cursor.fetchall())

if __name__ == '__main__':
    cherrypy.server.socket_host = '0.0.0.0'
    users = {"admin": "secretPassword", "editor": "otherPassword"}
    conf = {'/': {'tools.sessions.on': True,
                  'tools.staticdir.on': True,
                  'tools.staticdir.dir': os.path.abspath(os.getcwd())
                  },
            }
    root = HomePage()
    root.data = Data()
    root.about = About()
    root.login = Login()
    root.logout = Logout()
    root.signup = Signup()
    root.submit = Submit()
    cherrypy.quickstart(root, '/', config=conf)
