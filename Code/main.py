from tkinter import *
import tkinter.messagebox
import tkinter.filedialog
import pandas as pd
import email_function               #created .py file
import os
class Bulk_Email:
    def __init__(self,root):
        self.root=root
        self.root.title('Email Sender Application created by RITIK KUMAR GUPTA')
        self.root.geometry('900x550+200+50')
        self.root.overrideredirect(True)
        self.root.resizable(0,0)
        self.root.config(bg='khaki2')
        self.email_icon=PhotoImage(file='email.png')
        self.setting_icon=PhotoImage(file='setting.png')




        title=Label(self.root,text='  Multiple Email Sender',image=self.email_icon,anchor='w',padx=25
                    ,compound=LEFT,font=('Goudy Old Style',48,'bold'),bg='slate gray',
                    fg='white').place(x=0,y=0,relwidth=1)
        btn_setting=Button(self.root,image=self.setting_icon,bd=0, command=self.setting_window,bg='slate gray',activebackground='slate gray',cursor='hand2').place(x=800,y=5)
        desctitle = Label(self.root, text='Use Excel File to Send the Bulk Email at once, with just one click. Ensure the Email Column Name must be Email. ',
                       font=('calibri(Body)', 13, 'bold'), bg='gold',
                      fg='black').place(x=0, y=80, relwidth=1)

        self.var_choice=StringVar()
        single=Radiobutton(self.root,text='Single',value='single',variable=self.var_choice,command=self.check_single_or_bulk,font=('times new roman',30,'bold'),activebackground='khaki2',bg='khaki2',fg='red').place(x=50,y=150)
        multiple=Radiobutton(self.root,text='Multiple',value='multiple',command=self.check_single_or_bulk,variable=self.var_choice,activebackground='khaki2',font=('times new roman',30,'bold'),bg='khaki2',fg='red').place(x=250,y=150)
        self.var_choice.set('single')

        to=Label(self.root,text='To (Email Address)',font=('times new roman',18),bg='khaki2',fg='black').place(x=50,y=250)
        subject=Label(self.root,text='SUBJECT',font=('times new roman',18),bg='khaki2',fg='black').place(x=50,y=300)
        msg=Label(self.root,text='MESSAGE',font=('times new roman',18),bg='khaki2',fg='black').place(x=50,y=350)

        self.text_to=Entry(self.root,font=('times new roman',14),bg='lightyellow')
        self.text_to.place(x=300,y=250,width=300,height=30)

        self.btn_BROWSE = Button(self.root, text='BROWSE',command=self.browse_file,font=('times new roman', 14, 'roman'),state=DISABLED, bg='slate gray', fg='black',
                           cursor='hand2', activebackground='slate gray')
        self.btn_BROWSE.place(x=630, y=245, width=120)


        self.text_subj = Entry(self.root, font=('times new roman', 14), bg='lightyellow')
        self.text_subj.place(x=300, y=300, width=450, height=30)

        self.text_msg = Text(self.root, font=('times new roman', 14), bg='lightyellow')
        self.text_msg.place(x=300, y=350, width=450, height=120)

        #status

        self.totalLabel = Label(self.root, font=('times new roman', 18), bg='khaki2', fg='black')
        self.totalLabel.place(x=10, y=510)         #totallabel=status

        self.sentLabel = Label(self.root, font=('times new roman', 18), bg='khaki2', fg='red4')
        self.sentLabel.place(x=150, y=510)

        self.leftLabel = Label(self.root, font=('times new roman', 18), bg='khaki2', fg='red4')
        self.leftLabel.place(x=250, y=510)

        self.failedLabel = Label(self.root, font=('times new roman', 18), bg='khaki2', fg='red4')
        self.failedLabel.place(x=350, y=510)


        btn_EXIT=Button(self.root,text='EXIT',font=('times new roman',18,'bold'),command=self.iExit,bg='gray30',fg='white',activebackground='#00B0F0',cursor='hand2').place(x=770,y=490,width=120)
        btn_CLEAR=Button(self.root,text='CLEAR',command=self.clear1,font=('times new roman',18,'bold'),bg='gray30',fg='white',cursor='hand2',activebackground='#00B0F0').place(x=640,y=490,width=120)
        btn_SEND=Button(self.root,text='SEND',command=self.send_email,font=('times new roman',18,'bold'),bg='gray30',fg='white',cursor='hand2',activebackground='#00B0F0').place(x=510,y=490,width=120)
        self.check_file_exists()

    def send_email(self):
        x=len(self.text_msg.get('1.0',END))

        if self.text_to.get()==''or self.text_subj.get()==''or x==1:
            tkinter.messagebox.showerror("Error",'All fields are required',parent=self.root)
        else:
            if self.var_choice.get()=='single':
                status=email_function.send_email_func(self.text_to.get(),self.text_subj.get(),self.text_msg.get('1.0',END),
                                               self.from_,self.pass_)
                if status=='s':             #s = success
                    tkinter.messagebox.showinfo("Success", 'EMAIL HAS BEEN SENT', parent=self.root)

                if status == 'f':           #f= failed
                    tkinter.messagebox.showerror("Success", 'EMAIL NOT SENT,Try Again', parent=self.root)
            if self.var_choice.get()=='multiple':
                self.failed=[]
                self.s_count=0
                self.f_count=0
                for x in self.emails:                #self.mails=total emails in excel file
                    status=email_function.send_email_func(x, self.text_subj.get(),
                                                   self.text_msg.get('1.0', END),
                                                   self.from_, self.pass_)

                    if status == 's':
                        self.s_count+=1
                    if status == 'f':
                        self.f_count+=1
                    self.status_bar()

                tkinter.messagebox.showinfo("Success", 'EMAIL HAS BEEN SENT,Please Check Status', parent=self.root)

    def status_bar(self):
        self.totalLabel.config(text='Status:' + str(len(self.emails))+'=>>')
        self.sentLabel.config(text='SENT:'+str(self.s_count))
        self.leftLabel.config(text='LEFT:'+str(len(self.emails)-(self.s_count+self.f_count)))
        self.failedLabel.config(text='FAILED:'+str(self.f_count))

        self.totalLabel.update()
        self.sentLabel.update()
        self.leftLabel.update()
        self.failedLabel.update()



    def check_single_or_bulk(self):
        if self.var_choice.get()=='single':
            self.btn_BROWSE.config(state=DISABLED)
            self.text_to.config(state=NORMAL)
            self.text_to.delete(0,END)
            self.clear1()
        if self.var_choice.get()=='multiple':
            self.btn_BROWSE.config(state=NORMAL)
            self.text_to.config(state='readonly')
            self.text_to.delete(0, END)


    def clear1(self):
        self.text_to.config(state=NORMAL)
        self.text_to.delete(0,END)
        self.text_subj.delete(0,END)
        self.text_msg.delete('1.0',END)
        self.var_choice.set('single')
        self.btn_BROWSE.config(state=DISABLED)
        self.totalLabel.config(text='')
        self.sentLabel.config(text='')
        self.leftLabel.config(text='')
        self.failedLabel.config(text='')








    def iExit(self):
        ex = tkinter.messagebox.askyesno("Notification", "Confirm if you want to exit")
        if ex:
            root.destroy()



    def setting_window(self):
        self.check_file_exists()
        self.root2=Toplevel()              #root2 is client basically
        self.root2.title("Setting")
        self.root2.geometry('650x320+350+90')
        self.root2.focus_force()
        self.root2.config(bg='thistle2')
        self.root2.grab_set()
        title2 = Label(self.root2, text='Credentials Setting', image=self.setting_icon, anchor='w', padx=25
                      , compound=LEFT, font=('Goudy Old Style', 48, 'bold'), bg='slate gray',
                      fg='white').place(x=0, y=0, relwidth=1)

        desctitle2 = Label(self.root2,
                          text='Enter the email address and password from which you want to send the emails. ',
                          font=('calibri(Body)', 13, 'bold'), bg='gold',
                          fg='black').place(x=0, y=80, relwidth=1)
        
        from_=Label(self.root2,text='Email Address',font=('times new roman',18),fg='black',bg='thistle2').place(x=50,y=150)
        pass_=Label(self.root2,text='Password',font=('times new roman',18),fg='black',bg='thistle2').place(x=50,y=200)

        self.text_from = Entry(self.root2, font=('times new roman', 14), bg='lightyellow')
        self.text_from.place(x=250, y=150, width=330, height=30)

        self.text_pass = Entry(self.root2, font=('times new roman', 14), bg='lightyellow',show='*')
        self.text_pass.place(x=250, y=200, width=330, height=30)

        btn_SAVE = Button(self.root2, text='SAVE', command=self.save_setting,font=('times new roman', 18, 'bold'), bg='gold2', fg='black',
                           cursor='hand2', activebackground='#00B0F0').place(x=300, y=260, width=120,height=30)
        btn_CLEAR = Button(self.root2, text='CLEAR',command=self.clear2, font=('times new roman', 18, 'bold'), bg='gold2', fg='black',
                          cursor='hand2', activebackground='#00B0F0').place(x=430, y=260, width=120,height=30)
        self.text_from.insert(0,self.from_)
        self.text_pass.insert(0,self.pass_)


    def browse_file(self):
        op=tkinter.filedialog.askopenfile(initialdir='C:/Users/RITIK/PycharmProjects/GUI Email sender/Code',title="Select Excel File for Emails",filetype=(('All Files','*'),('Excel Files','xls')))
        if op!=None:
            data = pd.read_excel(op.name)

            if 'Email' in data.columns:
                self.emails = list(data['Email'])
                c = []
                for i in self.emails:
                    if pd.isnull(i) == False:
                        c.append(i)
                self.emails = c
                if len(self.emails)>0:
                    self.text_to.config(state=NORMAL)
                    self.text_to.delete(0,END)
                    self.text_to.insert(0, str(op.name.split('/')[-1]))
                    self.text_to.config(state='readonly')
                    self.totalLabel.config(text='Total:'+str(len(self.emails)))
                    self.sentLabel.config(text='SENT:')
                    self.leftLabel.config(text='LEFT:')
                    self.failedLabel.config(text='FAILED:')
                else:
                    tkinter.messagebox.showerror('Error',"This file doesn't have any emails",parent=self.root )
            else:
                tkinter.messagebox.showerror('Error','Please select a file which contains Email column',parent=self.root)



    def clear2(self):
        self.text_from.delete(0,END)
        self.text_pass.delete(0,END)

    def check_file_exists(self):
        if os.path.exists('important.txt')==False:
            f = open("important.txt", 'w')
            f.write(',')
            f.close()
        f2=open('important.txt','r')
        self.credentials=[]
        for i in f2:
            self.credentials.append([i.split(',')[0],i.split(',')[1]])
            # print(self.credentials)
            self.from_=self.credentials[0][0]
            self.pass_=self.credentials[0][1]






    def save_setting(self):
        if self.text_from.get()==''or self.text_pass.get()=='':
            tkinter.messagebox.showerror("Error", 'All fields are required', parent=self.root2)

        else:
            f=open("important.txt",'w')
            f.write(self.text_from.get()+','+self.text_pass.get())
            f.close()
            tkinter.messagebox.showinfo("Notification", 'SAVED SUCCESSFULLY', parent=self.root2)
            self.check_file_exists()

        




root=Tk()
obj=Bulk_Email(root)        #the whole code's soul is this "obj" object
root.mainloop()