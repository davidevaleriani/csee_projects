import cherrypy
import sqlite3
import hashlib
import os

db_name = "users.db"

class HomePage(object):
    @cherrypy.expose
    def index(self):
        return open("index.html").readlines()

class About(object):
    @cherrypy.expose
    def index(self):
        return open("about.html").readlines()

class Signup(object):
    @cherrypy.expose
    def index(self):
        return open("signup.html").readlines()

    @cherrypy.expose
    def register(self, username, password2, country, affiliation, name2, name1, password):
        if password != password2:
            return "ERRORE"
        password = hashlib.sha1(password).hexdigest()
        conn = sqlite3.connect(db_name)
        c = conn.cursor()
        c.execute("INSERT INTO users(name1,name2,affiliation,country,username,password) VALUES(?, ?, ?, ?, ?, ?)",
                  [name1, name2, affiliation, country, username, password])
        conn.commit()
        conn.close()
        return "OK"

class Login(object):
    @cherrypy.expose
    def index(self):
        return open("login.html").readlines()

class Data(object):
    @cherrypy.expose
    def index(self):
        return open("data.html").readlines()

if __name__ == '__main__':
    cherrypy.server.socket_host = '0.0.0.0'
    users = {"admin": "secretPassword", "editor": "otherPassword"}
    conf = {'/': {'tools.sessions.on': True,
                  'tools.staticdir.on': True,
                  'tools.staticdir.dir': os.path.abspath(os.getcwd())
                  },
            '/login': {'tools.digest_auth.on': True,
                       'tools.digest_auth.realm': 'CEEC 2015 Poker Competition',
                       'tools.digest_auth.users': users},
            }
    root = HomePage()
    root.data = Data()
    root.about = About()
    root.login = Login()
    root.signup = Signup()
    cherrypy.quickstart(root, '/', config=conf)
