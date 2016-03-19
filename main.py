__doc__ = info = '''
This program was written by Nicholas Smith - nicholas-smith.tk

Version 1.0.1
This script has three main classes:
Application - Page layouts
Tab - Basic tab used by TabBar for main functionality
TabBar - The tab bar that is placed above tab bodies (Tabs)

It uses a pretty basic structure:
root
-->TabBar(root, init_name) (For switching tabs)
-->Tab    (Place holder for content)
	\t-->content (content of the tab; parent=Tab)
-->Tab    (Place holder for content)
	\t-->content (content of the tab; parent=Tab)
-->Tab    (Place holder for content)
	\t-->content (content of the tab; parent=Tab)
etc.
'''
# Import smtplib for the actual sending function
import smtplib, sys, tkFileDialog, MySQLdb as mysql, threading
from Tkinter import *
from ttk import *
from ConfigParser import SafeConfigParser
import ttk
# Here are the email package modules we'll need
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import email.utils

BASE = RAISED
SELECTED = FLAT
        
class Application(Frame):
    def send(self):
        self.SEB.config(state=DISABLED)
        i=0
        t1_stop.set(False)
        self.popup = popup = Toplevel(self)
        Label(popup, text="Please wait emails send...").grid(row=0)
        Label(self.popup, textvariable=v).grid(row=1)
        v.set("Email number: %d" % (i))
        self.progressbar = progressbar = ttk.Progressbar(popup, orient=HORIZONTAL, length=200, mode='indeterminate')
        progressbar.grid(row=2)
        progressbar.start()
        t=threading.Thread(target=self.sendout)
        t.start()
        Button(popup, text="Stop", command=self.stopButton).grid(row=3, column=0)
        Button(popup, text="Dismiss", command=self.popup.destroy).grid(row=3, column=1)
        
    def stopButton(self):
        t1_stop.set(True)
        self.progressbar.stop()
        self.SEB.config(state=NORMAL)
        #self.popup.destroy()

    def sendout(self):
        SERVER = config.get('email', 'smtpserver')
        COMMASPACE = ', '
        FROM = config.get('email', 'from')
        if testemailstatus.get():
            TO = [config.get('email', 'to')]
            print TO
            print "Testing Email."
        else:
            # Connect to mysql
            if mysqlactive.get():
                try:
                    db = mysql.connect(host = config.get('mysql', 'server'), port = int(config.get('mysql', 'port')), user = config.get('mysql', 'username'), passwd = config.get('mysql', 'password'), db = config.get('mysql', 'database'))
                    cur = db.cursor()
                    cur.execute(config.get('mysql', 'sql'))
                    
                    rows = cur.fetchall()
                    for i, row in enumerate(rows):
                         TO = [row[0]]
                         v.set("Email %d: %s" % (i, TO))
                         if t1_stop.get():
                             break
                    #self.progressbar.stop()
                    self.stopButton()
                    db.close()
                except mysql.Error, e:
                    top = Toplevel()
                    TO = config.get('email', 'to')
                    top.title("MySQL Error %d" % e.args[0])
                    error_message="MySQL Error %d: %s" % (e.args[0],e.args[1])
                    msg = Message(top, text=error_message, width=225)
                    msg.pack()

                    button = Button(top, text="Dismiss", command=top.destroy)
                    button.pack()
                    print error_message
                                    
            else:
                print "MySQL turned off."
                TO = [config.get('email', 'to')]

        subject = config.get('email', 'subject')
        # Create the container (outer) email message.
        msg = MIMEMultipart()
        msg['Subject'] = '%s' % subject
        # me == the sender's email address
        # family = the list of all recipients' email addresses
        msg['From'] = FROM
        msg['To'] = COMMASPACE.join(TO)
        msg.preamble = 'Test email preamble'
        msg.date = email.utils.formatdate()
        
        # Assume we know that the image files are all in PNG format
        try:
                for file in pngfiles:
                        # Open the files in binary mode.  Let the MIMEImage class automatically
                        # guess the specific image type.
                        fp = open(file, 'rb')
                        img = MIMEImage(fp.read())
                        fp.close()
                        msg.attach(img)
        except: pass

        html = MIMEText(self.htmlbox.get('1.0', END).strip().encode('utf-8'), 'html')
        msg.attach(html)

        # Send the email via our own SMTP server.
        try:
                #s = smtplib.SMTP(SERVER)
                #s.sendmail(FROM, TO, msg.as_string())
                #s.quit()
                #print('Emails sent!')
            s=1
        except smtplib.SMTPException, e:
            top = Toplevel()
            TO = config.get('email', 'to')
            top.title("MySQL Error %d" % e.args[0])
            error_message="MySQL Error %d: %s" % (e.args[0],e.args[1])
            msg = Message(top, text=error_message, width=225)
            msg.pack()

            button = Button(top, text="Dismiss", command=top.destroy)
            button.pack()
            print error_message

    def createWidgets(self):
        def openfile():
            filename = askopenfilename(parent=root)
            f = open(filename)
            f.read()
        def saveSettings():
            config.set('email', 'smtpserver', smtpserver.get())
            config.set('email', 'From', fromemail.get())
            config.set('email', 'To', toemail.get())
            config.set('email', 'Bounce', toemail.get())
            config.set('email', 'subject', subject.get())
            with open('config.ini', 'w') as f:
                config.write(f)

        Label(self, text="SMTP Server", justify=LEFT).grid(row=0, sticky=E)
        e=Entry(self,textvariable=smtpserver)
        e.insert(END, config.get('email', 'smtpserver'))
        e.grid(row=0, column=1, sticky=W)
        
        Label(self, text="From", justify=LEFT).grid(row=1, sticky=E)
        e=Entry(self, textvariable=fromemail)
        e.insert(END, config.get('email', 'from'))
        e.grid(row=1, column=1, sticky=W)

        Label(self, text="To Test Email", justify=LEFT).grid(row=2, sticky=E)
        e=Entry(self, textvariable=toemail)
        e.insert(END, config.get('email', 'to'))
        e.grid(row=2, column=1, sticky=W)
        myB=Checkbutton(self, text="Test Email", variable=testemailstatus)
        myB.grid(row=3, column=2, sticky=E)
        
        Label(self, text="Subject", justify=LEFT).grid(row=2, sticky=E)
        e=Entry(self, textvariable=subject)
        e.insert(END, config.get('email', 'subject'))
        e.grid(row=3, column=1, sticky=W)

        
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
        
        if f is None: # asksaveasfile return `None` if dialog closed with "cancel".
            return
        
        text2save = self.htmlbox.get('1.0', END).strip().encode('utf-8') # starts from `1.0`, not `0.0`
        f.write(text2save)
        f.close() # `()` was missing.

    def onError(self):
        box.showerror("Error", "Could not open file")

    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack()
        self.createWidgets()

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
        if config.get('mysql', 'status') == 'True':
            mysqlactive.set(True)
        
        Label(self, text="Username", justify=LEFT).grid(row=0, sticky=E)
        e=Entry(self,textvariable=username)
        e.insert(END, config.get('mysql', 'username'))
        e.grid(row=0, column=1, sticky=W)

        Label(self, text="Password", justify=LEFT).grid(row=1, sticky=E)
        e=Entry(self, textvariable=password,show="*")
        e.insert(END, config.get('mysql', 'password'))
        e.grid(row=1, column=1, sticky=W)

        Label(self, text="Database", justify=LEFT).grid(row=2, sticky=E)
        e=Entry(self, textvariable=database)
        e.insert(END, config.get('mysql', 'database'))
        e.grid(row=2, column=1, sticky=W)

        Label(self, text="Server", justify=LEFT).grid(row=3, sticky=E)
        e=Entry(self, textvariable=server)
        e.insert(END, config.get('mysql', 'server'))
        e.grid(row=3, column=1, sticky=W)

        Label(self, text="Port", justify=LEFT).grid(row=4, sticky=E)
        e=Entry(self, textvariable=port)
        e.insert(END, config.get('mysql', 'port'))
        e.grid(row=4, column=1, sticky=W)

        myB = Checkbutton(self, text="Turn MySQL On", variable=mysqlactive, onvalue = True, offvalue = False)
        myB.grid(row=5, columnspan=10, sticky=W)
        
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
        
        self.sqlbox = Text(txt_frm, wrap=NONE,
            xscrollcommand=xscrollbar.set,
            yscrollcommand=yscrollbar.set)
        self.sqlbox.grid(row=0)
        self.sqlbox.focus_set()
        self.sqlbox.insert(END, config.get('mysql', 'sql'))
        xscrollbar.config(command=self.sqlbox.xview)
        yscrollbar.config(command=self.sqlbox.yview)
        Button(self, text="Save settings", command=saveSettings).grid(row=15)


    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack()
        self.createWidgets()

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
    config.set('email', 'SMTPServer', '')
    config.set('email', 'from', '')
    config.set('email', 'to', '')
    config.set('email', 'subject', '')
                       
with open('config.ini', 'w') as f:
    config.write(f)

root = Tk()
root.title("Email Marketing")
#root.geometry('1000x700+0+0')
mysqlactive = BooleanVar()
testemailstatus = BooleanVar()
t1_stop = BooleanVar()
password = StringVar()
username = StringVar()
database = StringVar()
server = StringVar()
port = StringVar()
smtpserver = StringVar()
fromemail= StringVar()
toemail=StringVar()
subject=StringVar()
v=StringVar()

note = Notebook(root)
tab1 = Frame(note)
tab2 = Frame(note)
tab3 = Frame(note)

Label(tab3, text="INFO:\n"+info, justify=LEFT).pack(side=TOP, expand=YES, fill=BOTH)
Application(master=tab1)
MySQLTab(master=tab2)
note.add(tab1, text = "Main")
note.add(tab2, text = "MySQL")
note.add(tab3, text = "Info")
note.pack()

root.mainloop()
root.destroy()
