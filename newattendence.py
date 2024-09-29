# Importing every modules used

from tkinter import *
import tkinter.messagebox as tmsg
import csv
import time
import clipboard as clb

#  To copy absentees data to clipboard using clipboard module

def coptoclb():
    abs=""
    j=0
    for i in range(0,120):
        if lt[i].get()==0:
            j+=1
            abs=abs+str(i+1)+"\t"
    abs = abs + "\nTotal absentees : "+str(j)
    a=status_.get()
    status_.set("Copied to clipboard .")
    l51.update()
    time.sleep(0.5)
    status_.set("Copied to clipboard ...")
    l51.update()
    time.sleep(0.5)
    status_.set(a)
    l51.update()
    clb.copy(abs)

# To check how many students are present and absent 

def Chc():
    j=0
    for i in range(0,120):
        if lt[i].get()==0:
            j=j+1
    noOfAbsent.set("Present : "+str(120-j)+"\nAbsent: "+str(j))
    l31.update()

# To update data in Excel sheet

def updtExcl():

    # Checking if data already present

    fl1=open("attendence.csv",mode="r",newline="")
    reader = csv.reader(fl1,delimiter=",")
    lt2=[]
    for row in reader:
        lt2.append(row)
    dat=str(time.localtime()[2])+"_"+str(time.localtime()[1])+"_"+str(time.localtime()[0])
    for i in lt2[0]:
        if i==dat:
            msg = tmsg.askquestion("Update error !!","Attendence already updated for {}\nDo you want to exit ?".format(dat) ,icon="error")
            if msg=='yes':
                root.destroy()
            return
    
    # Updating Excel sheet because same date data not found

    status_.set("Updating...")
    l51.update()
    fl2=open("attendence.csv",mode="w",newline="")
    writer=csv.writer(fl2,delimiter=",")
    lt2[0].append(dat)
    writer.writerow(lt2[0])
    for i in range(0,120):
        lt3=lt2[i+1]
        if lt[i].get()==0:
            lt3.append("A")
            writer.writerow(lt3)
        else:
            lt3.append("P")
            writer.writerow(lt3)
    updated.set(1)
    fl2.flush()
    fl1.flush()
    fl2.close()
    fl1.close()
    time.sleep(1)
    status_.set("Updated in Excel Sheet  :] ")
    l51.update()

# To destroy GUI on demand

def toExit():
    global root1

    # Direct exit if data updated in Excel sheet

    if updated.get() == 1:
        status_2.set("Exiting..")
        l52.update()
        time.sleep(2)
        root.destroy()

    # Confirming if user want to exit without updating in Excel sheet

    else:
        if exitClicked.get()==1:
            root1.destroy()
        def root1Destroy():
            exitClicked.set(0)
            root1.destroy()
        # def dB(event):
        #     destroyBoth()
        def destroyBoth(*arg):
            root1.destroy()
            root.destroy()
        root1 = Toplevel(root)
        root1.configure(background="#8be6fc")
        root1.geometry("360x240")
        # root1.wm_iconbitmap("1.ico")
        root1.maxsize(360,240)
        root1.minsize(360,240)
        root.winfo_screenheight()
        root.winfo_screenwidth()
        root1.geometry(f"+{int((root1.winfo_screenwidth()/2)-180)}+{int((root1.winfo_screenheight()/2)-120)}")
        root1.title("Are you sure ?")
        Label(root1,text="Attendence not updated \nin Excel Sheet\n",bg="#8be6fc" ,font="lucida 20 bold").pack()
        f=Frame(root1,bg="#8be6fc")
        f.pack()
        Button(f,text="Return",bg="#8be6fc",relief=RAISED,bd=4,activebackground="#4bcbeb", font="lucida 20 ", command=root1Destroy,justify=CENTER).grid(padx=36,row=0,column=1)
        fb1=Button(f,text="Exit",bg="#8be6fc",relief=RAISED,bd=4,activebackground="#4bcbeb",font="lucida 20 ", command=destroyBoth,justify=CENTER)
        fb1.focus_set()
        fb1.bind('<Return>',destroyBoth)
        fb1.grid(padx=36,row=0,column=2)
        exitClicked.set(1)

# To add roll no. to present list

def addRNo(event):
    global checked
    if event.keysym == "Return" and str(event.widget) ==".!frame2.!frame2.!entry":
        try:
            no=int(entRoll.get())
            entRoll.set("")
            e211.focus_set()
            lt[no-1].set(1)
            checked = no-1
            ckbt[no-1].update()
        except:
            entRoll.set("Error Value")
            e211.update()
            e211.config(state=DISABLED)
            time.sleep(.5)
            entRoll.set("")
            e211.config(state=NORMAL)
            e211.focus_set()
            return
    else:
        no = event.widget.cget("text")
        if int(no) == 120:
            no = 0
        if event.keysym=="??":
            ckbt[no-1].focus_set()
            checked = no-1
            return
        lt[no-1].set(1)
        checked=no-1
        ckbt[no-1].update()
        ckbt[int(no)].focus_set()
        checked=no

# To add roll no. to absent list

def skipRNo(event):
    no = event.widget.cget("text")
    if int(no) == 120:
        no = 0
    lt[no-1].set(0)
    checked = no-1
    ckbt[no-1].update()
    ckbt[no].focus_set()
    checked = no

# To change focus between Checkbox and Entry menu

def switchFrame(event):
    global checked
    x = str(event.widget)
    print(x)
    if x==".!frame2.!frame2.!entry":
        ckbt[checked].focus_set()
    else:
        e211.focus_set()

# Help menu for all commands information

def helpDat():
    tmsg.showinfo("Help","→ Press 'left-ctrl' to switch between checkboxes\n\t and roll no entry\n\n→ Use 'arrow' keys to navigate between checkboxes\n\n→ Press 'enter' for present and '0' for absent,\n\tand move to next checkbox\n\n→ 'Show' button shows number of absent\n\tand present entities according to entry\n\n→ 'Copy' button will copy absentees roll no. and \n\tno. of absentees to clipboard\n\n→ 'Update' button will update data in excel according\n\t to entry")

# To move between checkboxes

def moveInChk(event):
    global checked
    if str(event.keysym) == "Left":
        if checked < 12:
            checked = 108+checked
        else :
            checked = checked - 12
        ckbt[checked].focus_set()
    if str(event.keysym) == "Right":
        if checked > 107 :
            checked = checked - 108
        else :
            checked = checked + 12
    if str(event.keysym) == "Up":
        if checked == 0:
            checked = 119
        else :
            checked = checked - 1
    if str(event.keysym) == "Down":
        if checked == 119:
            checked = 0
        else :
            checked = checked + 1
    ckbt[checked].focus_set()

# main GUI

if __name__ == '__main__':
    checked = 0
    root1 = int()
    root=Tk()
    root.title("Attendence Updater")
    root.geometry("1280x720")
    # root.wm_iconbitmap("1.ico")

    # Header  setup for GUI

    f1 = Frame(root,bg="#4ac6ff",height=200,width=1280)
    l1= Label(f1,text="Attendence", font="Comicsansms 50 italic", bg="#4ac6ff")
    l1.pack(side=LEFT, padx=50)
    x=str(time.localtime()[2])+ '-' + str(time.localtime()[1])+ '-' + str(time.localtime()[0])

    l2=Label(f1,text=f"Date : {x}",font="Comicsansms 30 ", bg="#4bcbeb")
    l2.pack(side=RIGHT,padx=10)
    f1.pack(side=TOP,fill="both")

    # Main body ( Checkboxex and Entry ) of GUI

    f2=Frame(root,bg="#8be6fc",width=1280,height=400)
    f3=Frame(f2,bg="#8be6fc",width=1250,height=400,bd=10,relief=RIDGE)

    lt=[]
    k=0
    j=0
    ckbt=[0]*120
    for i in range(1,121):
        lt.append(IntVar())
        k=k+1
        ckbt[i-1] = Checkbutton(f3,text=i,font="TimesRoman 12 bold", variable=lt[i-1],padx=25,cursor="cross" , activebackground="grey",bd=7,bg="#8be6fc")
        ckbt[i-1].grid(row=k,column=j)
        ckbt[i-1].bind('<0>',skipRNo)
        ckbt[i-1].bind('<Return>',addRNo)
        ckbt[i-1].bind('<Button-1>',addRNo)
        ckbt[i-1].bind("<Control_L>",switchFrame)
        ckbt[i-1].bind("<Up>",moveInChk)
        ckbt[i-1].bind("<Down>",moveInChk)
        ckbt[i-1].bind("<Left>",moveInChk)
        ckbt[i-1].bind("<Right>",moveInChk)
        if i%12 == 0:
            j=j+1
            k=0
    f3.pack(side=LEFT)

    f21 =Frame(f2,bg="#8be6fc")
    entRoll = StringVar()
    l211 = Label(f2,font="lucida 13 bold", bg="#8be6fc", text="Roll no")
    l211.pack(pady=12)
    e211 = Entry(f21,font="lucida 12", textvariable=entRoll)
    e211.pack(padx=5)
    e211.bind('<Return>',addRNo)
    f21.pack()
    f2.pack(fill=BOTH)
    e211.bind("<Control_L>",switchFrame)
    ckbt[0].focus_set()

    # Variable declaration for GUI

    updated=IntVar()
    updated.set(0)
    exitClicked=IntVar()
    exitClicked.set(0)
    
    noOfAbsent = StringVar()
    noOfAbsent.set("No. of absent/present\n students ")

    # Main functions menu of GUI

    f4=Frame(root,bg="#8be6fc",height=100)
    l31 = Label(f4,textvariable=noOfAbsent,justify=CENTER, height=1, width=20,font="lucida 20 bold",fg="black",bg="#8be6fc",bd=5,relief=SUNKEN)
    l31.pack(side=LEFT, expand=True,ipadx=15, ipady=15)
    Button(f4,text="<= Show",command=Chc,font="lucida 20 bold",activebackground="#4bcbeb", bg="#8be6fc", height=1, width=8).pack(side=LEFT,padx=28,pady=12, expand=True,ipadx=7,ipady=7)
    Button(f4,text="Copy absentees \n numbers",command=coptoclb,font="lucida 20 bold",activebackground="#4bcbeb", bg="#8be6fc", height=1, width=14).pack(side=LEFT, expand=True,ipadx=7,ipady=7,padx=25,pady=6)
    Button(f4,text="Update Excel \nSheet",command=updtExcl,font="lucida 20 bold",activebackground="#4bcbeb", bg="#8be6fc", height=1, width=12).pack(side=LEFT, expand=True,ipadx=7,ipady=7,padx=25,pady=8)
    Button(f4,text="EXIT",command=toExit,font="lucida 20 bold",activebackground="#4bcbeb", bg="#8be6fc", height=1, width=14).pack(side=RIGHT,ipadx=48, expand=True,ipady=7)
    f4.pack(fill=BOTH,expand=True)

    # Status bar variables

    status_ = StringVar()
    status_2 = StringVar()
    status_.set("Not Updated in Excel Sheet !")
    status_2.set("")

    # Status bar GUI

    f5=Frame(root,bg="#0099ff",height=20)
    l51=Label(f5,textvariable=status_,font="lucida 12 bold",fg="black", bg="#0099ff")
    l51.pack(side=LEFT)
    l52=Label(f5,textvariable=status_2,font="lucida 12 bold",fg="black", bg="#0099ff", justify=CENTER)
    l52.pack()
    f5.pack(side=BOTTOM,fill=BOTH)

    # Menubar setup for GUI

    menubar=Menu(root)
    m1=Menu(menubar,tearoff=0)
    m1.add_command(label="Help",command=helpDat)
    m1.add_separator()
    m1.add_command(label="Emergency exit !",command=root.destroy)
    menubar.add_cascade(label="Help",menu=m1)
    root.config(menu=menubar)

    # Configuring Windows exit button to work accordingly 

    root.protocol("WM_DELETE_WINDOW", toExit)

    # Executing GUI

    root.mainloop()