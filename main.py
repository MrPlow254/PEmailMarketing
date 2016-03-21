__version__ = version = "1.0.6"
__status__ = "Development"
__copyright__ = copy  = "Copyright 2016, Nicholas Smith (nicholas-smith.tk)"
__doc__ = info = '''This program was written by Nicholas Smith - nicholas-smith.tk\n
Version %s\n
Python 2.7.11\n
Modules used:
MySQLdb\n
Since version 1.0.5:
5 hours of research
~40 hours of coding & debuging.
Since version 1.0.6:
~1 hour of coding\n
This script can get emails from any mysql table structure. It will send out the html file
on the the Main (SMTP) tab. It is possible to import the html from any filetype. It is also
possible to save the html file to any filetype (default html) if needed. The html section is
not saved in the program and will need to be loaded or typed each time the program is opened.\n
To Do:\n
- Option to update table with the status of the email to prevent bad emails from staying
in the list.\n
%s
''' % (version, copy)

# Import smtplib for the actual sending function
import email.utils, MySQLdb as mysql, smtplib, sys, threading, tkFileDialog, ttk
import tkMessageBox as messagebox, Queue
from ConfigParser import SafeConfigParser
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr
from Tkinter import *
from ttk import *
BASE = RAISED
SELECTED = FLAT

class Application(Frame):
    def send(self):
        self.SEB.config(state=DISABLED)
        i=0
        t1_stop.set(False)
        mysqlError.set(False)
        self.popup = popup = Toplevel(self)
        Label(self.popup, text="Please wait emails send...").grid(row=0, columnspan=2)
        Label(self.popup, textvariable=v).grid(row=1, columnspan=2)
        v.set("Email: %d" % (i))
        self.progressbar = progressbar = ttk.Progressbar(self.popup, orient=HORIZONTAL, length=200, mode='indeterminate')
        progressbar.grid(row=2, columnspan=2)
        progressbar.start()
        t=threading.Thread(target=self.sendout)
        t.start()
        self.myS=myS=Button(self.popup, text="Stop", command=self.stopButton)
        myS.grid(row=3, column=0)
        self.myB=myB=Button(self.popup, text="Dismiss", command=self.popup.destroy)
        myB.grid(row=3, column=1)
        myB.config(state=DISABLED)

    def stopButton(self):
        t1_stop.set(True)
        try:
            self.progressbar.stop()
        except: pass
        try:
            self.myB.config(state=NORMAL)
        except: pass
        self.SEB.config(state=NORMAL)
        if mysqlError.get() or smtpError.get():
            self.popup.destroy()
        else:
            self.myS.config(state=DISABLED)

    def mysqlPopup(self):
        if mysqlError.get():
            try:
                e = self.callback_queue.get()
                error_message="MySQL Error %d: %s" % (e.args[0],e.args[1])
                popupMySQL = Toplevel(self)
                popupMySQL.title("MySQL Error %d" % e.args[0])
                popupMySQL.lift()
                popupMySQL.grab_set()
                msg = Message(popupMySQL, text=error_message, width=300)
                msg.pack()
                button = Button(popupMySQL, text="Dismiss", command=popupMySQL.destroy)
                button.pack()
                self.stopButton()
                mysqlError.set(False)
                self.master.after(200, self.mysqlPopup)
            except:
                pass
        else:
            self.master.after(200, self.mysqlPopup)

    def smtpPopup(self):
        if smtpError.get():
            try:
                e = self.callback_queue.get()
                error_message="SMTP Error %d: %s" % (e.args[0],e.args[1])
                popupSMTP = Toplevel(self)
                popupSMTP.title("SMTP Error %d" % e.args[0])
                popupSMTP.lift()
                popupSMTP.grab_set()
                msg = Message(popupSMTP, text=error_message, width=300)
                msg.pack()
                button = Button(popupSMTP, text="Dismiss", command=popupSMTP.destroy)
                button.pack()
                self.stopButton()
                smtpError.set(False)
                self.master.after(200, self.smtpPopup)
            except:
                pass
        else:
            self.master.after(200, self.smtpPopup)

    def sendout(self):
        if testemailstatus.get():
            TO = [config.get('email', 'to')]
            v.set("Email 1: %s" % TO[0])
            self.setup(TO)
            self.stopButton()
        else:
            # Connect to mysql
            if mysqlactive.get():
                try:
                    db = mysql.connect(host = config.get('mysql', 'server'), port = config.getint('mysql', 'port'), user = config.get('mysql', 'username'), passwd = config.get('mysql', 'password'), db = config.get('mysql', 'database'))
                    cur = db.cursor()
                    cur.execute(config.get('mysql', 'sql'))
                    rows = cur.fetchall()
                    for i, row in enumerate(rows):
                         TO = [row[0]]
                         v.set("Email %d" % (i))
                         self.setup(TO)
                         if t1_stop.get():
                             break
                    v.set("Email %d: %s" % (i, TO[0]))
                    self.stopButton()
                    db.close()
                except mysql.Error, e:
                    TO = config.get('email', 'to')
                    mysqlError.set(True)
                    self.callback_queue.put(e)
            else:
                TO = [config.get('email', 'to')]
                v.set("Email 1: %s" % TO[0])
                self.setup(TO)

    def setup(self, TO):
        SERVER = config.get('smtp', 'server')
        smtpport = config.getint('smtp', 'port')
        COMMASPACE = ', '
        FROM = formataddr((str(config.get('email', 'fromn')), config.get('email', 'from')))
        subject = config.get('email', 'subject')
        # Create the container (outer) email message.
        msg = MIMEMultipart()
        msg['Subject'] = '%s' % subject
        msg['From'] = FROM
        msg['To'] = COMMASPACE.join(TO)
        msg.preamble = 'PEmailMarketing v1.x'
        msg.date = email.utils.formatdate()
        html = MIMEText(self.htmlbox.get('1.0', END).strip().encode('utf-8'), 'html')
        msg.attach(html)
        # Send the email via our own SMTP server.
        try:
            s = smtplib.SMTP(SERVER, smtpport)
        except smtplib.socket.gaierror, e:
            error_message="SMTP Error %d: %s" % (e.args[0],e.args[1])
            smtpError.set(True)
            self.callback_queue.put(e)
            return False

        try:

            if config.get('smtp', 'username') != '':
                s.login(config.get('smtp', 'username'), config.get('smtp', 'password'))
        except smtplib.SMTPException, errormsg:
            print errormsg

        try:
            s.sendmail(FROM, TO, msg.as_string())
        except smtplib.SMTPException, errormsg:
            smtpError.set(True)
            e="Couldn't send message: %s" % (errormsg)
            self.callback_queue.put(e)
            return False
##        except smtplib.socket.timeout:
##            smtpError.set(True)
##            e="Socket error while sending message"
##            self.callback_queue.put(e)
##            return False
        finally:
            s.quit()

    def createWidgets(self):
        def openfile():
            filename = askopenfilename(parent=root)
            f = open(filename)
            f.read()
        def saveSettings():
            config.set('smtp', 'server', smtpserver.get())
            config.set('smtp', 'port', smtpport.get())
            config.set('smtp', 'username', smtpusername.get())
            config.set('smtp', 'password', smtppassword.get())
            config.set('email', 'From', fromemail.get())
            config.set('email', 'FromN', fromname.get())
            config.set('email', 'To', toemail.get())
            config.set('email', 'Bounce', toemail.get())
            config.set('email', 'subject', subject.get())
            with open('config.ini', 'w') as f:
                config.write(f)
        def select_all(event):
            self.htmlbox.tag_add(SEL, "1.0", END)
            self.htmlbox.mark_set(INSERT, "1.0")
            self.htmlbox.see(INSERT)
            return 'break'

        Label(self, text="SMTP Server:", justify=LEFT).grid(row=0, sticky=E)
        e=Entry(self,textvariable=smtpserver)
        e.insert(END, config.get('smtp', 'server'))
        e.grid(row=0, column=1, sticky=W)

        Label(self, text="SMTP Port:", justify=LEFT).grid(row=0, column=2, sticky=E)
        e=Entry(self,textvariable=smtpport)
        e.insert(END, config.get('smtp', 'port'))
        e.grid(row=0, column=3, sticky=W)

        Label(self, text="SMTP Username:", justify=LEFT).grid(row=1, column=2, sticky=E)
        e=Entry(self,textvariable=smtpusername)
        e.insert(END, config.get('smtp', 'username'))
        e.grid(row=1, column=3, sticky=W)

        Label(self, text="SMTP Password:", justify=LEFT).grid(row=2, column=2, sticky=E)
        e=Entry(self,textvariable=smtppassword,show="*")
        e.insert(END, config.get('smtp', 'password'))
        e.grid(row=2, column=3, sticky=W)

        Label(self, text="From Name:", justify=LEFT).grid(row=1, sticky=E)
        e=Entry(self, textvariable=fromname)
        e.insert(END, config.get('email', 'fromn'))
        e.grid(row=1, column=1, sticky=W)

        Label(self, text="From:", justify=LEFT).grid(row=2, sticky=E)
        e=Entry(self, textvariable=fromemail)
        e.insert(END, config.get('email', 'from'))
        e.grid(row=2, column=1, sticky=W)

        Label(self, text="To Test Email:", justify=LEFT).grid(row=3, sticky=E)
        e=Entry(self, textvariable=toemail)
        e.insert(END, config.get('email', 'to'))
        e.grid(row=3, column=1, sticky=W)
        myB=Checkbutton(self, text="Test Email:", variable=testemailstatus)
        myB.grid(row=3, column=2, sticky=E)

        Label(self, text="Subject:", justify=LEFT).grid(row=4, sticky=E)
        e=Entry(self, textvariable=subject)
        e.insert(END, config.get('email', 'subject'))
        e.grid(row=4, column=1, sticky=W)


        txt_frm = Frame(self, width=600, height=400)
        txt_frm.grid(row=6, columnspan=10)
        # ensure a consistent GUI size
        txt_frm.grid_propagate(False)
        # implement stretchability
        txt_frm.grid_rowconfigure(0, weight=1)
        txt_frm.grid_columnconfigure(0, weight=1)

        xscrollbar = Scrollbar(txt_frm, orient=HORIZONTAL)
        xscrollbar.grid(row=1, sticky='nsew')
        yscrollbar = Scrollbar(txt_frm)
        yscrollbar.grid(row=0, column=1, sticky='nsew')

        self.htmlbox = Text(txt_frm, wrap=NONE,
            xscrollcommand=xscrollbar.set,
            yscrollcommand=yscrollbar.set)
        self.htmlbox.grid(row=0)
        self.htmlbox.insert(END,
"""<html>
    <head></head>
    <body>
        <p>Hi!<br>
            How are you?<br>
            Here is the <a href="https://www.python.org">link</a> you wanted.
        </p>
    </body>
</html>""")
        self.htmlbox.focus_set()
        self.htmlbox.bind("<Control-Key-a>", select_all)
        xscrollbar.config(command=self.htmlbox.xview)
        yscrollbar.config(command=self.htmlbox.yview)

        Button(self, text="Save settings", command=saveSettings).grid(row=15, column=0)
        self.SEB = SEB = Button(self, text="Send Email", command=self.send)
        SEB.grid(row=15, column=1)
        menubar = Menu(self)
        fileMenu = Menu(menubar, tearoff=0)
        fileMenu.add_command(label="Open", command=self.openFile)
        fileMenu.add_command(label="Save As", command=self.saveFile)
        fileMenu.add_separator()
        fileMenu.add_command(label="Exit", command=root.quit)
        menubar.add_cascade(label="File", menu=fileMenu)
        root.config(menu=menubar)

    def openFile(self):
        ftypes = [('HTML files', '*.html'), ('All files', '*')]
        dlg = tkFileDialog.Open(self, filetypes = ftypes)
        fl = dlg.show()
        if fl != '':
            text = self.readFile(fl)
            self.htmlbox.delete('1.0', END)
            self.htmlbox.insert(END, text)

    def readFile(self, filename):
        f = open(filename, "r")
        text = f.read()
        return text

    def saveFile(self):
        ftypes = [('HTML files', '*.html'), ('All files', '*')]
        f = tkFileDialog.asksaveasfile(mode='w', defaultextension=".html", filetypes = ftypes)
        if f is None:
            return
        text2save = self.htmlbox.get('1.0', END).strip().encode('utf-8')
        f.write(text2save)
        f.close()

    def onError(self):
        box.showerror("Error", "Could not open file")

    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack()
        self.createWidgets()
        self.mysqlPopup()
        self.smtpPopup()
        self.callback_queue = Queue.Queue()

class MySQLTab(Frame):
    def createWidgets(self):
        def saveSettings():
            config.set('mysql', 'Username', username.get())
            config.set('mysql', 'Password', password.get())
            config.set('mysql', 'Database', database.get())
            config.set('mysql', 'Server', server.get())
            config.set('mysql', 'Port', port.get())
            config.set('mysql', 'Status', str(mysqlactive.get()))
            config.set('mysql', 'SQL', self.sqlbox.get('1.0', END).strip().encode('utf-8'))
            with open('config.ini', 'w') as f:
                config.write(f)

        def select_all(event):
            self.sqlbox.tag_add(SEL, "1.0", END)
            self.sqlbox.mark_set(INSERT, "1.0")
            self.sqlbox.see(INSERT)
            return 'break'

        if config.get('mysql', 'status') == 'True':
            mysqlactive.set(True)

        Label(self, text="Username:", justify=LEFT).grid(row=0, sticky=E)
        e=Entry(self,textvariable=username)
        e.insert(END, config.get('mysql', 'username'))
        e.grid(row=0, column=1, sticky=W)

        Label(self, text="Password:", justify=LEFT).grid(row=1, sticky=E)
        e=Entry(self, textvariable=password,show="*")
        e.insert(END, config.get('mysql', 'password'))
        e.grid(row=1, column=1, sticky=W)

        Label(self, text="Database:", justify=LEFT).grid(row=2, sticky=E)
        e=Entry(self, textvariable=database)
        e.insert(END, config.get('mysql', 'database'))
        e.grid(row=2, column=1, sticky=W)

        Label(self, text="Server:", justify=LEFT).grid(row=3, sticky=E)
        e=Entry(self, textvariable=server)
        e.insert(END, config.get('mysql', 'server'))
        e.grid(row=3, column=1, sticky=W)

        Label(self, text="Port:", justify=LEFT).grid(row=4, sticky=E)
        e=Entry(self, textvariable=port)
        e.insert(END, config.get('mysql', 'port'))
        e.grid(row=4, column=1, sticky=W)

        myB = Checkbutton(self, text="Turn MySQL On", variable=mysqlactive, onvalue = True, offvalue = False)
        myB.grid(row=5, columnspan=10, sticky=W)

        txt_frm = Frame(self, width=600, height=400)
        txt_frm.grid(row=6, columnspan=10)
        txt_frm.grid_propagate(False)
        txt_frm.grid_rowconfigure(0, weight=1)
        txt_frm.grid_columnconfigure(0, weight=1)

        xscrollbar = Scrollbar(txt_frm, orient=HORIZONTAL)
        xscrollbar.grid(row=1, sticky='nsew')
        yscrollbar = Scrollbar(txt_frm)
        yscrollbar.grid(row=0, column=1, sticky='nsew')

        self.sqlbox = Text(txt_frm, wrap=NONE,
            xscrollcommand=xscrollbar.set,
            yscrollcommand=yscrollbar.set)
        self.sqlbox.grid(row=0)
        self.sqlbox.focus_set()
        self.sqlbox.bind("<Control-Key-a>", select_all)
        self.sqlbox.insert(END, config.get('mysql', 'sql'))
        xscrollbar.config(command=self.sqlbox.xview)
        yscrollbar.config(command=self.sqlbox.yview)
        Button(self, text="Save settings", command=saveSettings).grid(row=15)

    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack()
        self.createWidgets()

def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        root.quit()
        #root.destroy()
def on_esc(event):
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        root.quit()
        root.destroy()

config = SafeConfigParser()
config.read('config.ini')
if not config.has_section('mysql'):
    config.add_section('mysql')
    config.set('mysql', 'Username', '')
    config.set('mysql', 'Password', '')
    config.set('mysql', 'Database', '')
    config.set('mysql', 'Server', '')
    config.set('mysql', 'Port', '3306')
    config.set('mysql', 'SQL', '')
    config.set('mysql', 'Status', 'True')

if not config.has_section('email'):
    config.add_section('email')
    config.set('email', 'from', '')
    config.set('email', 'fromn', '')
    config.set('email', 'to', '')
    config.set('email', 'subject', '')

if not config.has_section('smtp'):
    config.add_section('smtp')
    config.set('smtp', 'Server', '')
    config.set('smtp', 'Port', '25')
    config.set('smtp', 'username', '')
    config.set('smtp', 'password', '')

with open('config.ini', 'w') as f:
    config.write(f)

root = Tk()
root.title("PEmail Marketing")
mysqlactive = BooleanVar()
testemailstatus = BooleanVar()
mysqlError = BooleanVar()
smtpError = BooleanVar()
t1_stop = BooleanVar()
password = StringVar()
username = StringVar()
database = StringVar()
server = StringVar()
port = StringVar()
smtpserver = StringVar()
smtpport = StringVar()
smtpusername = StringVar()
smtppassword = StringVar()
fromemail= StringVar()
fromname=StringVar()
toemail=StringVar()
subject=StringVar()
v=StringVar()

note = Notebook(root)
tab1 = Frame(note)
tab2 = Frame(note)
tab3 = Frame(note)

Label(tab3, text="INFO:\n"+info, justify=LEFT, anchor=W, width=200).pack(side=TOP, expand=YES, fill=BOTH)
Application(master=tab1)
MySQLTab(master=tab2)
note.add(tab1, text = "Main")
note.add(tab2, text = "MySQL")
note.add(tab3, text = "Info")
note.pack(expand=YES, fill=BOTH)

root.protocol("WM_DELETE_WINDOW", on_closing)
root.bind('<Escape>', on_esc)
root.mainloop()
root.destroy()
