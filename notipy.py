import gtk
import pango
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
import mimetypes
import email
import imaplib
import email.mime.application
import smtplib
import threading
import glib
from multiprocessing import *
from ConfigParser import SafeConfigParser
from Tkinter import Tk
from tkFileDialog import askopenfilename
import sys
import os
import unidecode
import html2txt
theme="Nodoka-Aqua"
gtk.rc_parse("G:\\gtk+-bundle_2.22.1-20101229_win64\\share\\themes\\"+theme+"\\gtk-2.0\\gtkrc")
#gtk.require('2.0')
class mailThread(threading.Thread):
    def __init__(self,message,to,subject,filelist):
        threading.Thread.__init__(self)
        self.filelist=filelist
        self.message=message
        self.to=to
        self.subject=subject
        self.start() # invoke the run method
    
    def run(self):
        try:
            parser=SafeConfigParser()
            sig="""\n\n\n message sent via Notipy, personal email client.
                          https://github.com/thekindlyone/notipy
            """
            self.message+=sig
            parser.read('d:/pythonary/notipy/notipy_config.ini')
            emid=parser.get('credentials','emailid')
            pw=parser.get('credentials','password')
            smtp_server=parser.get('credentials','smtpserver')
            #print 'here'   
            #print self.to,self.message,self.subject
            msg = email.mime.Multipart.MIMEMultipart()
            msg.attach(MIMEText(self.message))
            if(self.filelist):
                for filename in self.filelist:
                    fp=open(filename,'rb')
                    #att = email.mime.application.MIMEApplication(fp.read(),_subtype="mp3")
                    att = email.mime.application.MIMEApplication(fp.read())
                    fp.close()
                    att.add_header('Content-Disposition','attachment',filename=filename)
                    msg.attach(att)
            sender = emid    
            recipients = self.to
            msg['Subject'] = self.subject
            msg['From'] = sender
            msg['To'] = ", ".join(recipients)
            s = smtplib.SMTP(smtp_server)
            s.starttls()
            s.login(msg['From'],pw)    
            s.sendmail(sender, recipients, msg.as_string())
            s.quit()
            print 'done'
        except Exception, e:
            print e.message

       


class fetchmailThread(threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)
        self.fetchedmails=None
        self.start() # invoke the run method
    
    def run(self):
        try:
            parser=SafeConfigParser()
            parser.read('d:/pythonary/notipy/notipy_config.ini')
            emid=parser.get('credentials','emailid')
            pw=parser.get('credentials','password')            
            mail = imaplib.IMAP4_SSL('imap.gmail.com')
            mail.login(emid, pw)
            #print mail.list()
            # Out: list of "folders" aka labels in gmail.
            mail.select("inbox") # connect to inbox.
            result, data = mail.uid('search', None, "UNSEEN") # search and return uids instead
            #print result
            latest_emails_uid = data[0].split()[-10:]
            emails=[]
            for uid in latest_emails_uid:
                #result, data = mail.uid('fetch',uid, '(RFC822)')
                result, data = mail.uid('fetch',uid, '(BODY.PEEK[])')
                emails.append(email.message_from_string(data[0][1]))
            self.fetchedmails=emails
            mail.close()
        except Exception, e:
            print e.message
        

 









class Entry(object):
    def __init__(self,name,address):
        self.name=name
        self.address=address
    def getdata(self):
        return(self.name,self.address)    


class mainwindow(gtk.Window):
    def __init__(self,filenamearg):
        super(mainwindow, self).__init__()
        self.set_title("notiPy") 
              
        try:
            self.set_icon_from_file("d:\\pythonary\\notipy\\notipycon.png")
        except Exception, e:
            print e.message
            sys.exit(1)
        self.options={}        
        self.options['initialdir'] = ''
        self.fetched=[]
        self.mainbox=gtk.VBox(False, 5)
        self.row1=gtk.HBox(False, 5)
        self.row2=gtk.HBox(False, 5)
        width= gtk.gdk.screen_width()/2
        height=gtk.gdk.screen_height()/1.7
        self.set_size_request(int(width), int(height))
        self.set_position(gtk.WIN_POS_CENTER)
        self.tobox=gtk.Entry()
        self.selected=[]
        #self.tobox.set_text("TO")
        self.tolabel=gtk.Label("To:")
        self.sublabel=gtk.Label("Subject:")
        #self.bodylabel=gtk.Label("Body:")
        self.subject=gtk.Entry()
        #self.subject.set_text("subject")
        self.outgoing = gtk.ScrolledWindow()
        self.textview = gtk.TextView()
        self.textview.modify_font(pango.FontDescription('serif 16'))
        self.textview.set_wrap_mode(gtk.WRAP_WORD)
        self.textview.set_left_margin(5)
        textbuffer=self.textview.get_buffer()
        #textbuffer.connect("changed",self.scrolltocursor)
        #self.combo = gtk.Combo()
        self.tooltips = gtk.Tooltips()
        self.send_button=gtk.Button()
        try:
            image_send = gtk.Image()
            image_send.set_from_file("send.png")
            image_send.show()
            self.send_button.add(image_send)
        except Exception, e:
            print e.message
        self.tooltips.set_tip(self.send_button, "Send")
        self.send_button.connect("clicked", self.sendmail)
        self.row1.pack_start(self.tolabel,fill=False,expand=False,padding=5)
        self.row1.pack_start(self.tobox,fill=True,expand=True,padding=5)
        self.row1.pack_start(self.send_button,expand=False,fill=False,padding=5)
        self.row2.pack_start(self.sublabel,expand=False,fill=False,padding=5)
        self.row2.pack_start(self.subject,expand=True,fill=True,padding=5)
        self.mainbox.pack_start(self.row1,expand=False,fill=False,padding=5)
        self.mainbox.pack_start(self.row2,expand=False,fill=False,padding=5)
        self.outgoing.set_policy(gtk.POLICY_AUTOMATIC, gtk.POLICY_AUTOMATIC)
        self.outgoing.add(self.textview)
        self.outbox=gtk.VBox(False,10)
        self.outbox.set_border_width(10)
        self.outbox.pack_start(self.outgoing,expand=True,fill=True)
        #self.alignment = gtk.Alignment(xalign=0.01, yalign=0.0, xscale=0.0, yscale=0.0)
        #self.alignment.add(self.bodylabel)
        self.row3=gtk.HBox(5,False)
        #self.row3.pack_start(self.alignment,expand=False,fill=False)
        
        self.entry_count=self.getbook();        
        bookrow1=gtk.HBox(False,1)
        bookrow2=gtk.HBox(False,1)        
        boxes = sorted(self.book.keys(), key=self.sortfun)             
        for i in range(self.entry_count):
            boxes[i].connect("toggled", self.checkcallback, self.book[boxes[i]].getdata()[0])
            if(i%2==0):
                
                bookrow1.pack_start(boxes[i],False,False,5)
                
            else:
                bookrow2.pack_start(boxes[i],False,False,5)
        
        bookbox=gtk.VBox(False,1)
        bookbox.pack_start(bookrow1,fill=False,expand=False,padding=5)
        bookbox.pack_start(bookrow2,fill=False,expand=False,padding=5)
        bookscroll=gtk.ScrolledWindow()
        bookscroll.set_policy(gtk.POLICY_AUTOMATIC, gtk.POLICY_AUTOMATIC)
        bookscroll.add_with_viewport(bookbox)        
        
        self.row3.pack_start(bookscroll,expand=True,fill=True)
        
        
        
        self.mainbox.pack_start(self.row3,expand=False,fill=False,padding=10)
        self.mainbox.pack_start(self.outbox,expand=True,fill=True)
        self.filelist=[]

            
        #liststore=gtk.ListStore(str)
        filescroll=gtk.ScrolledWindow()
        filescroll.set_shadow_type(gtk.SHADOW_ETCHED_IN)
        filescroll.set_policy(gtk.POLICY_AUTOMATIC, gtk.POLICY_AUTOMATIC)
        attbox=gtk.HBox(False,5)
        attbox.pack_start(filescroll,True,True,0)
        attachbutton=gtk.Button(stock=gtk.STOCK_OPEN)
        attachbutton.connect("clicked", self.attachmentselect)
        alignment=gtk.Alignment(xalign=1.0, yalign=1.0, xscale=0.0, yscale=0.0)
        alignment.add(attachbutton)
        attbox.pack_start(alignment,False,False,0)
        self.store=self.create_model()
        treeview=gtk.TreeView(self.store)
        treeview.set_rules_hint(True)
        treeview.connect("row-activated", self.on_activated)
        column = gtk.TreeViewColumn("Attachments")
        treeview.append_column(column)
        
        cell = gtk.CellRendererText()
        column.pack_start(cell, False)
        column.add_attribute(cell, "text", 0)
        filescroll.add(treeview)
        self.mainbox.pack_start(attbox)
        if filenamearg:
            self.filelist.append(filenamearg)            
            self.options['initialdir']= os.path.dirname(filenamearg)        
            #print self.filelist
            self.store.append([filenamearg])
        #########################################################################################
        mainbox2=gtk.VBox(False,5)
        self.inboxstore=self.create_model_inbox()
        inboxtreeview=gtk.TreeView(self.inboxstore)
        inboxtreeview.set_rules_hint(True)
        self.create_columns(inboxtreeview)
        scrollinbox=gtk.ScrolledWindow()
        scrollinbox.set_shadow_type(gtk.SHADOW_ETCHED_IN)
        scrollinbox.set_policy(gtk.POLICY_AUTOMATIC, gtk.POLICY_AUTOMATIC)
        scrollinbox.add(inboxtreeview)
        self.mailtextbox=gtk.TextView()
        self.mailtextbox.modify_font(pango.FontDescription('serif 16'))
        self.mailtextbox.set_wrap_mode(gtk.WRAP_WORD)
        self.mailtextbox.set_left_margin(5)
        self.mailtextbox.set_editable(False)
        mailtextbuffer=self.mailtextbox.get_buffer()
        inboxtreeview.connect("row-activated", self.on_activatedmail,mailtextbuffer) 
        scrolltextbox=gtk.ScrolledWindow() 
        scrolltextbox.set_shadow_type(gtk.SHADOW_ETCHED_IN)
        scrolltextbox.set_policy(gtk.POLICY_AUTOMATIC, gtk.POLICY_AUTOMATIC)
        scrolltextbox.add(self.mailtextbox)        
        refreshbutton=gtk.Button(stock=gtk.STOCK_REFRESH)
        refreshbutton.connect("clicked", self.fetchmail)
        paned=gtk.VPaned()
        paned.add1(scrollinbox)
        paned.add2(scrolltextbox) 
        paned.set_position(150)        
        mainbox2.pack_start(paned,True,True) 
        bbox=gtk.HBox(False,5)
        clearbutton=gtk.Button(stock=gtk.STOCK_CLEAR)
        clearbutton.connect("clicked", self.clearbox,mailtextbuffer)
        bbox.pack_start(clearbutton)
        bbox.pack_start(refreshbutton)      
        mainbox2.pack_start(bbox,False,False)
        
        
        
        ############################################################################################
        
        notebook = gtk.Notebook()
        notebook.set_tab_pos(gtk.POS_LEFT)
        notebook.append_page(self.mainbox, gtk.Label("COMPOSE"))
        notebook.append_page(mainbox2, gtk.Label("Inbox"))
        self.add(notebook)
        
    def checkcallback(self, widget, data=None):
        #print data,widget.get_active(),self.book[widget].getdata()[1]
        state=widget.get_active()
        name=data
        address=self.book[widget].getdata()[1]
        if state:
            self.selected.append(address)
        if not state:
            self.selected=filter(lambda a: a != address, self.selected)
        #print self.selected
    
    def sortfun(self,x):
        return self.book[x].getdata()[0]
    def getbook(self):
        self.book={}
        with open('d:/pythonary/notipy/addressbook.txt','r') as f:
            for line in f:
                line=line.strip()
                elems=line.split(',')
                name=elems[0].strip()
                address=elems[-1].strip()
                chbutton= gtk.CheckButton(name)
                self.book[chbutton]=Entry(name,address)
        return len(self.book.keys())
        
        
        
    def sendmail(self, widget):
        to=self.tobox.get_text().strip().split(',')+self.selected
        to[:] = [item.strip() for item in to]
        to=list(set(to))
        if(len(to)!=0):
            textbuffer=self.textview.get_buffer()
            start=textbuffer.get_start_iter()
            end=textbuffer.get_end_iter()
            message=textbuffer.get_text(start,end,include_hidden_chars=True)
            
            sub=self.subject.get_text()
            self.t=mailThread(message,to,sub,self.filelist)
            self.send_button.set_sensitive(False)
            glib.timeout_add_seconds(3,self.maildone)
          
    def maildone(self):
        if (not self.t.is_alive()):
            self.send_button.set_sensitive(True)
        else:
            glib.timeout_add_seconds(1,self.maildone)
        
    def attachmentselect(self,widget):
        Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
        
        filename = askopenfilename(**self.options) # show an "Open" dialog box and return the path to the selected file
        #print(filename)
        if (filename):
            self.filelist.append(filename)
            self.filelist=list(set(self.filelist))
            self.options['initialdir']= os.path.dirname(filename)        
            #print self.filelist
            self.store.append([filename])
            
    def create_model(self):
        '''create the model - a ListStore'''        
        store = gtk.ListStore(str)
        for item in self.filelist:
            store.append([item])
 
        return store

    def on_activated(self,widget,srow,scol):
        model = widget.get_model()
        #print model[srow][0]
        it= model.get_iter(srow) 
        #it=
        self.filelist=filter(lambda a: a != model[srow][0], self.filelist)
        self.store.remove(it)        

        
        
    def create_model_inbox(self):
        '''create the model - a ListStore'''        
        store = gtk.ListStore(str,str)
        #emails=[('from1','sub1'),('from2','sub2')]
        #for email in emails:
            #store.append([email[0], email[1]])
        return store
    def create_columns(self, treeView):
    
        rendererText = gtk.CellRendererText()
        column = gtk.TreeViewColumn("From", rendererText, text=0)
        #column.set_sort_column_id(0)    
        treeView.append_column(column)
        
        rendererText = gtk.CellRendererText()
        column = gtk.TreeViewColumn("Subject", rendererText, text=1)
        #column.set_sort_column_id(1)
        treeView.append_column(column)

    def fetchmail(self,widget):
        thrd=fetchmailThread()
        #thrd.start()
        thrd.join()
        
        if thrd.fetchedmails:
            for mess in reversed(thrd.fetchedmails):
                self.fetched.append((mess['From'],mess['Subject'],self.get_first_text_block(mess)))
            #print fetched
            #self.inboxstore=gtk.ListStore(str,str)
            #self.inboxstore.clear()
            for item in (self.fetched):                
                self.inboxstore.append([item[0],item[1]])
    
    def clearbox(self,widget,textbuffer):
        self.fetched=[]
        self.inboxstore.clear()
        textbuffer.set_text(' ')
    
    
    
    
    
    def on_activatedmail(self,widget,srow,scol,textbuffer):
        model = widget.get_model()        
        it= model.get_iter(srow)
        bodytext=self.fetched[int(model.get_string_from_iter(it))][2]
        #print self.fetched[int(model.get_string_from_iter(it))]
        if(bodytext==None):
            bodytext=" "
        textbuffer.set_text(bodytext)
        #self.store.remove(it)     

    def get_first_text_block(self, email_message_instance):
        maintype = email_message_instance.get_content_maintype()
        if maintype == 'multipart':
            for part in email_message_instance.get_payload():
                if part.get_content_maintype() == 'text':
                    return part.get_payload()
        elif maintype == 'text':
            return email_message_instance.get_payload()    
   

def main():
    fn=''
    if len(sys.argv)>1:
        fn=sys.argv[1]
    root=mainwindow(fn)
    root.connect("destroy", gtk.main_quit)
    #root.maximize()
    root.show_all()
    gtk.gdk.threads_init()  
    gtk.main()

if __name__ == '__main__':
    main()    
