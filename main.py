from tkinter import *
from PIL import ImageTk
from PIL import Image
from tkinter import messagebox
import tkinter as tk
from tkVideoPlayer import TkinterVideo
import webbrowser
from tkinter import Toplevel
import os
from tkinter import filedialog
import tkinter.filedialog as fdialog
from tkinter.filedialog import askopenfilename
import pandas as pd
import time
import email_function
import smtplib

class Bulk_EMAIL:
    def __init__(self,root):
        self.root=root
        self.root.title("BULK EMAIL SENDER")
        self.root.geometry("1280x750+200+50")
        # Label(self.root,text ='The Foldcurl Company',bg="White",fg="Red", font =('Courier', 15)).pack(side = BOTTOM)
        self.root.resizable(False,False)
        self.root.config(bg="White")
        
        
        # ========= Icons ========
        image = Image.open('email.png')
        img = image.resize((50,50))
        self.email_icon = ImageTk.PhotoImage(img)
        
        image2 = Image.open('setting.png')
        img2 = image2.resize((60,60))
        self.setting = ImageTk.PhotoImage(img2)

        # ========= Title ========
        
        title=Label(self.root,text="Bulk Email Sender",image=self.email_icon,padx=10,compound=LEFT,font=("Goudy Old Style",48,"bold"),bg = "#222A35",fg="Yellow",anchor="w").place(x=0,y=0,relwidth=1)
        dev=Label(self.root,text="Developed by Kawsar",font=("Times New Roman",30,"bold"),bg = "#222A35",fg="White",).place(x=700,y=15)
        des=Label(self.root,text="Use Excel File to Send Bulk Email at Once, With Just One Click. Ensure The Email Column Name Be Email",font=("Calibri (Body)",14),bg = "#FFD966",fg="Black").place(x=0,y=82,relwidth=1)

        btn_setting = Button(self.root,image=self.setting,bd=0,activebackground="#222A35",cursor="hand2",command=self.setting_window).place(x=1200,y=7)
        
        
        #--------------------
        
        slt=Label(self.root,text="Select Your Option",font=("Goudy Old Style",25,"bold"),bg="White",fg="#262626").place(x=50,y=150)
        slt1=Label(self.root,text="[Note:Single For One Person, Bulk for Multiple Person]",font=("Goudy Old Style",25,"bold")).place(x=318,y=150)
        self.var_choice = StringVar()
        single = Radiobutton(self.root,text="Single",command=self.check_single_bulk,value="single",variable=self.var_choice,activebackground="white",font=("times new roman",30,"bold"),bg="White",fg="blue").place(x=50,y=210)
        bulk = Radiobutton(self.root,text="Bulk",value="bulk",command=self.check_single_bulk,variable=self.var_choice,activebackground="white",font=("times new roman",30,"bold"),bg="White",fg="blue").place(x=300,y=210)
        self.var_choice.set("single")
        
        #---------------------
        
        to = Label(self.root, text="To (Enter Email Address)",font=("times new roman",18),bg = "white").place(x = 50,y=320)
        sub = Label(self.root, text="SUBJECT",font=("times new roman",18),bg = "white").place(x = 50,y=420)
        msg = Label(self.root, text="MESSAGE",font=("times new roman",18),bg = "white").place(x = 50,y=520)
        
        
        self.txt_to=Entry(self.root,font=("times new roman",14),bg="lightyellow")
        self.txt_to.place(x=400,y=320,width=320,height=35)
        
        self.brwose = Button(self.root,text="Import",font=("times new roman",18,"bold"),command=self.browse_file,bg="#8FAADC",activebackground="#8FAADC",activeforeground="#262626",fg="#262626",cursor="hand2",state=DISABLED)
        self.brwose.place(x=750,y=320,width=100,height=35)
        
        
        self.txt_sub=Entry(self.root,font=("times new roman",14),bg="lightyellow")
        self.txt_sub.place(x=400,y=420,width=450,height=35)
        
        self.txt_msg=Text(self.root,font=("times new roman",14),bg="lightyellow")
        self.txt_msg.place(x=400,y=520,width=580,height=160)
        
        
        
        #------ about me -------
        
        github = Button(self.root,text="GITHUB",font=("times new roman",15,"bold"),bg="#00B0F0",activebackground="#00B0F0",activeforeground="white",fg="black",cursor="hand2",command=self.github_window).place(x=1100,y=280,width=120,height=45)
        facebook = Button(self.root,text="MAIL US",font=("times new roman",15,"bold"),bg="#262626",activebackground="#262626",activeforeground="white",fg="white",cursor="hand2",command=self.support_mail).place(x=1100,y=335,width=120,height=45)
        gmail = Button(self.root,text="FACEBOOK",font=("times new roman",15,"bold"),bg="#00B0F0",activebackground="#00B0F0",activeforeground="white",fg="black",cursor="hand2",command=self.facebook_window).place(x=1100,y=390,width=120,height=45)
        github = Button(self.root,text="TUTORIAL",font=("times new roman",15,"bold"),bg="#262626",activebackground="#262626",activeforeground="white",fg="white",cursor="hand2",command=self.youtube_window).place(x=1100,y=445,width=120,height=45)
        #-------------------
        
        about = Label(self.root,text="Reach Me",font=("times new roman",20,"bold"),bg="White",relief=SUNKEN,fg="#262626").place(x=1100,y=220)
        
        
    #----- Status----
        self.total1= Label(self.root,font=("times new roman",18),bg = "white")
        self.total1.place(x = 50,y=706)
        self.sent1= Label(self.root,font=("times new roman",18),bg = "white",fg="green")
        self.sent1.place(x = 310,y=706)
        self.left1= Label(self.root,font=("times new roman",18),fg="orange",bg = "white")
        self.left1.place(x = 550,y=706)
        self.failed1= Label(self.root,font=("times new roman",18),fg="red",bg = "white")
        self.failed1.place(x = 750,y=706)
        
        clear_send = Button(self.root,text="CLEAR",command=self.clear1,font=("times new roman",18,"bold"),bg="#262626",activebackground="#262626",activeforeground="white",fg="white",cursor="hand2").place(x=1020,y=695,width=120)
        btn_send = Button(self.root,text="SEND",font=("times new roman",18,"bold"),command=self.send_mail,bg="#00B0F0",activebackground="#00B0F0",activeforeground="white",fg="black",cursor="hand2").place(x=1150,y=695,width=120)
        self.check_file()
    
    
    def email_send_funct(to_,subj_,msg_,from_,pass_):
        s = smtplib.SMTP("s,tp.gmail.com",587)
        s.starttls()
        s.login(from_,pass_)
        msg="Subject: {}\n\n{}".format(subj_,msg_)
        s.sendmail(from_,to_,msg_)
        x=s.ehlo()
        if x[0]==250:
            return "s"
        else:
            return"f"
        s.close()
            
    
    def send_mail(self):
        x = len(self.txt_msg.get('1.0',END))
        if self.txt_to.get()== "" or self.txt_sub.get()=="" or x==1:
            messagebox.showerror("Error","All Fields Are Required",parent=self.root)
        else:
            if self.var_choice.get()=='single':
                status=email_function.email_send_funct(self.txt_to.get(),self.txt_sub.get(),self.txt_msg.get('1.0',END),self.from_,self.pass_)
                if status=="s":
                    messagebox.showinfo("Success","Your Email Has been Sent",parent=self.root)
                if status=="f":
                    messagebox.showerror("Error","Your Email Has been Not Sent",parent=self.root)
            if self.var_choice.get()=="bulk":
                self.failed=[]
                self.s_count=0
                self.f_count=0
                for x in self.emails:
                    status=email_function.email_send_funct(x,self.txt_sub.get(),self.txt_msg.get('1.0',END),self.from_,self.pass_)
                    if status=="s":
                        self.s_count+=1
                    if status=="f":
                        self.f_count+=1
                    self.status_bar()
                    # time.sleep(1)
                messagebox.showinfo("Success","Your Email Has been Sent",parent=self.root)
                        
                        
    def status_bar(self):
        self.total1.config(text="STATUS: "+str(len(self.emails))+"=>>")
        self.sent1.config(text="SENT: "+str(self.s_count))
        self.left1.config(text="LEFT: "+str(len(self.emails)-(self.s_count+self.f_count)))
        self.failed1.config(text="FAILED: "+str(self.f_count))
        self.total1.update()
        self.sent1.update()
        self.left1.update()
        self.failed1.update()
        
            
    def check_single_bulk(self):
        if self.var_choice.get()=="single":
            self.txt_to.config(state=NORMAL)
            self.txt_to.delete(0,END)
            self.brwose.config(state=DISABLED)
            self.clear1()
        if self.var_choice.get()=="bulk":
            self.brwose.config(state=NORMAL)
            self.txt_to.delete(0,END)
            self.txt_to.config(state='readonly')
            
            
    def clear1(self):
        self.txt_to.config(state=NORMAL)
        self.txt_to.delete(0,END)
        self.txt_sub.delete(0,END)
        self.txt_msg.delete("1.0",END)
        self.var_choice.set("single")
        self.brwose.config(state=DISABLED)
        self.total1.config(text="")
        self.sent1.config(text="")
        self.left1.config(text="")
        self.failed1.config(text="")
    
    #------Setting Part----
    
    def setting_window(self):
        self.root2 = Toplevel()
        self.root2.title("Setting")
        self.root2.geometry("780x450+550+150")
        self.root2.focus_force()
        self.root2.grab_set()
        
        title2=Label(self.root2,text="Credentials Settings",image=self.setting,padx=10,compound=LEFT,font=("Goudy Old Style",48,"bold"),bg = "#222A35",fg="White").place(x=0,y=0,relwidth=1)
        des2=Label(self.root2,text="Enter The Email Address and Password From Which to Send All Emails",font=("Calibri (Body)",14),bg = "#FFD966",fg="Black").place(x=0,y=82,relwidth=1)
        
        fr = Label(self.root2, text="Email Address",font=("times new roman",18)).place(x = 50,y=250)
        pas = Label(self.root2, text="Password",font=("times new roman",18)).place(x = 50,y=300)
       
        self.email=Entry(self.root2,font=("times new roman",14),bg="lightyellow")
        self.email.place(x=300,y=250,width=320,height=35)
       
        self.passs=Entry(self.root2,font=("times new roman",14),bg="lightyellow")
        self.passs.place(x=300,y=300,width=320,height=35)
        
        clear = Button(self.root2,text="CLEAR",command=self.clear2,font=("times new roman",18,"bold"),bg="#262626",activebackground="#262626",activeforeground="white",fg="white",cursor="hand2").place(x=350,y=350,width=120)
        save = Button(self.root2,text="SAVE",command=self.save_setting,font=("times new roman",18,"bold"),bg="#00B0F0",activebackground="#00B0F0",activeforeground="white",fg="black",cursor="hand2").place(x=480,y=350,width=120)
        if self.email and self.passs:
            self.email.insert(0,self.email)
            self.passs.insert(0,self.passs)
        
  #---- Browse File-----
    
    def browse_file(self):
        op=filedialog.askopenfile(initialdir="/",title="Select Excel File for Emails",filetypes=(("All File","*.*"),("Excel file",".xlsx")))
        if op!=None:
            data=pd.read_excel(op.name)
            if 'Email' in data.columns:
                self.emails = list(data['Email'])
                c=[]
                for i in self.emails:
                    if pd.isnull(i)==False:
                        c.append(i)
                self.emails=c
                if len(self.emails)>0:
                    self.txt_to.config(state=NORMAL)
                    self.txt_to.delete(0,END)
                    self.txt_to.insert(0,str(op.name.split("/")[-1]))
                    self.txt_to.config(state='readonly')
                    self.total1.config(text="Total: "+str(len(self.emails)))
                    self.sent1.config(text="SENT: ")
                    self.left1.config(text="LEFT: ")
                    self.failed1.config(text="FAILED: ")
                    
                else:
                    messagebox.showerror("Error","Import with Email Columns",parent=self.root)
            else:
                messagebox.showerror("Error","Please Import a Excel File",parent=self.root)
    #----- existing file-----
    
    def check_file(self):
        if os.path.exists("LoginInfo.txt")==False:
            f=open('LoginInfo.txt','w')
            f.write(",")
            f.close()
        f2=open('LoginInfo.txt','r')
        self.credentials=[]
        for i in f2:
            self.credentials.append ( [i.split(",")[0],i.split(",")[1]] )
        self.email = self.credentials[0][0]
        self.passs = self.credentials[0][1]  
    #----- Save Config | Clear Config ------
    
    def save_setting(self):
        if self.email.get()=="" or self.passs.get()=="":
            messagebox.showerror("Error","All Fields Are Required",parent=self.root2)
        else:
            f=open('LoginInfo.txt','w')
            f.write(self.email.get()+","+self.passs.get())
            f.close()
            messagebox.showinfo("Success","You Logged in Successfully",parent=self.root2)
            self.check_file()
    
    def clear2(self):
        self.email.delete(0,END)
        self.passs.delete(0,END)    
    #------Github--------
    def github_window(self):
        github_link = "https://github.com/curl-kawsar"
        webbrowser.open_new(github_link)     
    #----- Facebook ------
    def facebook_window(self):
        facebook_link = "https://facebook.com/python.kawsar"
        webbrowser.open_new(facebook_link)    
    #----- Gmail------    
    def support_mail(self):
        to_address = "knownaskawsar@gmail.com"
        subject = "Your Subject"  
        body = "Your email body"  
        gmail_link = f"mailto:{to_address}?subject={subject}&body={body}"
        webbrowser.open_new(gmail_link)   
    #----- Youtube ----
    def youtube_window(self):
        facebook_link = "https://facebook.com/python.kawsar"
        webbrowser.open_new(facebook_link)
    
        
    
root=Tk()
Obj=Bulk_EMAIL(root)
root.mainloop()