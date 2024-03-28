import pandas as pd
import datetime
import tkinter as tk

from tkinter import filedialog
from tkinter import  *
from tkinter import ttk
from PIL import ImageTk,Image
 
import time
import requests
from lxml import html
import webbrowser


df=pd.read_excel(r"Book2.xlsx")
try:
    
    
    k=df.values
except:
    pass
    
else:
    pass

Loc=k[:,0]
groups=k[:,1]

Location=list(Loc)
groups=list(groups)

kasi="paradiSo"
Loc=[m.lower() for m in Location]






def group_number(kasi):
    kasi=kasi.casefold()
    if kasi in Loc:
            i=Loc.index(kasi)
            return groups[i]
        

aloc=group_number(kasi)

what_stage="stage 1"
what_stage=what_stage.lower()
what_stage=what_stage.capitalize()



d=pd.read_excel(r"Book3 schedule.xlsx")

l=d.values






if what_stage in list(l[:,2]):
    stage1_rows=[]
    for i , k in enumerate(list(l[:,2])):
        if what_stage== k :
            stage1_rows.append(i)
            
                         
                         
                              
    


def make_stage(what_stage):

    if what_stage in list(l[:,2]):
        stage_rows=[]
        
        for i , k in enumerate(list(l[:,2])):
            if what_stage== k :
                stage_rows.append(i)
        return stage_rows        

stage1_rows=make_stage("Stage 1")
stage2_rows=make_stage("Stage 2")
stage3_rows=make_stage("Stage 3")
stage4_rows=make_stage("Stage 4")
stage5_rows=make_stage("Stage 5")
stage6_rows=make_stage("Stage 6")
stage7_rows=make_stage("Stage 7")
stage8_rows=make_stage("Stage 8")

stage8_rows+=stage1_rows+stage2_rows+stage3_rows+stage4_rows+ stage5_rows+stage6_rows+stage7_rows

stage7_rows+=stage1_rows+stage2_rows+stage3_rows+stage4_rows+stage5_rows+stage6_rows

stage6_rows+=stage1_rows+stage2_rows+stage3_rows+stage4_rows+stage5_rows

stage5_rows+=stage1_rows+stage2_rows+stage3_rows+stage4_rows

stage4_rows+=stage1_rows+stage2_rows+stage3_rows

stage3_rows+=stage1_rows+stage2_rows

stage2_rows+=stage1_rows


stage1_rows.sort()
stage2_rows.sort()
stage3_rows.sort()
stage4_rows.sort()
stage5_rows.sort()
stage6_rows.sort()
stage7_rows.sort()
stage8_rows.sort()





u=datetime.datetime.today()
u.day #date from Pc




u=datetime.datetime.today()
u.day #date from Pc

#reshaspe
x=50

root=tk.Tk()
canvas1=tk.Canvas(root,width=500, height= 380,bg="black",highlightthickness=0)

clicked=StringVar()
clicked.set(Location[120])

try:
    with open('last_location.txt') as prv_loc:
        last_loc=prv_loc.read()
        clicked.set(last_loc)
        prv_loc.close()
        
except:
    pass

option=tk.OptionMenu(root,clicked,*Location)
option.config(bg="black",width=14,fg="white", bd=0,highlightthickness=0,anchor="w")
canvas1.create_window(200+x,160, window=option)
canvas1.create_text(100+x,160, text="Area :",fill="white")

prev=StringVar(root)
time=StringVar(root)
time.set("@")
numbr=IntVar(root)
numbr.set(0)

# Create a function to scroll and highlight the option menu based on the entered keys
def scroll_highlight(event):
    # Get the currently entered keys
    keys = event.keysym
    
    
    name_list=[]
    if numbr.get()==0:
        prev.set(keys.lower())
    name_list=[]
   
    
    # Loop through the names in the list
    for name in Location:
        # If the currently entered keys match the first letter of the name, highlight that name in the option menu
        if keys.lower() == name[0].lower() :
            name_list.append(name)
   
    
    if prev.get()==keys.lower():
        clicked.set(name_list[numbr.get()])
        if numbr.get()+1 < len(name_list):
            
            numbr.set(numbr.get()+1)
        else:
            numbr.set(0)
    else:
        numbr.set(0)

        
option.bind("<Key>",scroll_highlight,"+")
option.focus()
stages_=[0,1,2,3,4,5,6,7,8]
clicked2=StringVar()
clicked2.set(stages_[1])

option2=tk.OptionMenu(root,clicked2,*stages_)
option2.config(bg="black",fg="white",width=1,bd=0,highlightthickness=0)

canvas1.create_window(200+x,240, window=option2)

canvas1.pack()


#aloc=group_number(clicked.get())

ui=[stage1_rows,stage2_rows,stage3_rows,stage4_rows,stage5_rows,stage6_rows
  ,stage7_rows, stage8_rows]






save=[]
store=[]
def stage_number(no):
    
    if no==0:
        lbl=tk.Label(root, text = "No Load shedding",fg="white",bg="black")
        lbl.pack()
        save.append(lbl)
        
        return
        
        
    aloc=group_number(clicked.get())
    kaosane="no"
    lis_r=ui[no-1]
    
    
    if var.get()==1:
        
        kaosane="yes"
        
    for i in lis_r:
        
        if kaosane=="no":
        
            if l[i,u.day+2]==aloc:
                frm=l[i,0]
                tot=l[i,1]

                tot=tot.strftime("%H:%M %p")
                frm=frm.strftime("%H:%M")

                store.append(frm)
                jk=[frm,r"to",tot,"Today"]
                jk= " ".join(jk)



                label=tk.Label(root, text = jk,fg="white",bg="black")

                label.place(y=22,x=10)


                h=label.pack()

                save.append(label)
            
            
        
            
            
            
        
        elif kaosane=="yes":
            
            if l[i,u.day+3]==aloc:
                
                frm=l[i,0]
                tot=l[i,1]

                tot=tot.strftime("%H:%M %p")
                frm=frm.strftime("%H:%M")

                store.append(frm)
                jk=[frm,r"to",tot,"Tomorrow"]
                jk= " ".join(jk)



                label=tk.Label(root, text = jk,fg="white",bg="black")

                label.place(y=22,x=10)


                h=label.pack()

                save.append(label)
            
    if store==[]:
        if kaosane=="yes":
            lbl=tk.Label(root, text = "No Load shedding Tomorrow ",fg="white",bg="black")
            lbl.pack()
            store.append(lbl)
        else:
            lbl=tk.Label(root, text = "No Load shedding Today ",fg="white",bg="black")
            lbl.pack()
            store.append(lbl)
            
def execute():
    if save!= []:
        
        for i in save:
            
            i.pack_forget()
            
    if store!= []:
        
        
        
        for i in store:
            
            
            if type(i)==str:
                store.remove(i)
            
            else:
                
                i.pack_forget()
                store.remove(i)
                
                
                     
        
       
    stage_number(int(clicked2.get()) ) 
    
def execute_2(event):
    if save!= []:
        
        for i in save:
            
            i.pack_forget()
            
    if store!= []:
        
        
        
        for i in store:
            
            if type(i)==str:
                store.remove(i)
            
            else:
                
                i.pack_forget()
                store.remove(i)
                
                
                     
        
       
    stage_number(int(clicked2.get()) ) 
       
    
    



class HoverButton(tk.Button):
    def __init__(self, master, **kw):
        tk.Button.__init__(self,master=master,**kw)
        self.defaultBackground = self["background"]
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)

    def on_enter(self, e):
        self['background'] = self['activebackground']

    def on_leave(self, e):
        self['background'] = self.defaultBackground



 
button1=tk.Button(root,text= "Schedule",fg="white",bg="black",width=8, command= execute , bd=0 )
button1.bind("<Return>",execute_2)
button1.focus()
canvas1.create_window(200+x,340, window= button1) 


root.iconbitmap("Vexels-Office-Bulb.ico")
canvas1.create_text(150+x,240, text="Stage :",fill="white")
root.title("Chilled Schedules")


Image_=tk.PhotoImage(master=canvas1,file="Chilled Schedules.png").subsample(17)
canvas1.create_image(200+x,90,image=Image_)

Image1=tk.PhotoImage(master=canvas1,file="BAR.png").subsample(3)
canvas1.create_image(200.35+x,356,image=Image1)
#button1.config(image=Image1,compound=CENTER)






Image4=tk.PhotoImage(master=canvas1,file="instag.png").subsample(16)
canvas1.create_image(100+x,340,image=Image4)


    
root.configure(background="black")

root.resizable(width=False,height = True)


                                        #Loadshsedding Status Check
    


power_note=tk.StringVar()  
power_title=tk.StringVar()
def check_status():
    
    try:
   
        
        option=requests.get("https://www.eskom.co.za/")
        tree=html.fromstring(option.content)
        #power_title.set(tree.xpath('//span [@class="elementor-alert-title"]/text()')[0]) #title of the eskom update window
        #power_note.set(tree.xpath('//div//span [@class="elementor-alert-description"]/text()')[0]) # eskom update
        
        try:
            option_2=requests.get("https://loadshedding.eskom.co.za/LoadShedding/GetStatus").text
            
            if option_2 =="1":

                state="NOT"
                stage="LOADSHEDDING."
                clicked2.set(0) #set stage 0
                

            elif int(option_2) > 1:


                state="LOADSHEDDING STAGE"
                stage=str(int(option_2)-1)
               
            else:
                
                state="NOT"
                stage="LOADSHEDDING"  
                clicked2.set(0) #set stage 0
        except:
            
 
            state,stage=tree.xpath('//p//b ["a href"]/text()')
        try:
            power_title.set(tree.xpath('//span [@class="elementor-alert-title"]/text()')[0])
            power_note.set(tree.xpath('//div//span [@class="elementor-alert-description"]/text()')[0])
        
        except:
            power_title.set("Loadshedding Status")
            power_note.set(state+" "+stage)
            
        status=state+" "+stage
        return status

   

    except:
        pass
    

try:
     
    o=check_status() #o=Status_Check()
  
    
    color="#f8f852"
    font_size=7
    stg_n=o.split()[-1]
    if stg_n.isnumeric():
    
        clicked2.set(int(stg_n))
except:
    
    o="Check Internet Connection"
    color="red"
    font_size=7

    

    
    


sts=tk.StringVar()  
sts.initialize(o)
#clr=tk.StringVar()  
#clr.initialize(color)




def reload():
    
    try:
        #o=Status_Check()
        color="#f8f852"
        #sts.set(o)
        #clr.set(color)
            
        o=check_status()
        
        #o=browser.find_element_by_xpath('//div//p//span [@id="lsstatus"]').text
        sts.set(o)
        
        stg_n=o.split()[-1]
        if stg_n.isnumeric():
             
            clicked2.set(int(stg_n)) 
          
        widget.configure(fg=color,font=("Helvetica",7))
    except:
        o="Check internet Connection"
        color="red" 
        sts.set(o)
        widget.configure(fg=color,font=("Helvetica",7))
        #clr.get()
    
  
      

button_reload=HoverButton(root,fg="white",bg="white",activebackground="blue",width=29,height=25, command= reload , bd=0 )
canvas1.create_window(420+x,240, window=button_reload)
  
Image19=tk.PhotoImage(master=canvas1,file="refresh boss.png").subsample(15)

canvas1.create_image(420+x,240,image=Image19)
button_reload.config(image=Image19,compound=CENTER,) 
canvas1.create_text(300+x,210,text="Status:", fill = "white")
#LS_Condition=canvas1.create_text(380+x,211,text=sts.get(), fill = color,font=("Trajan",8))
widget=tk.Label(root,textvariable=sts,fg=color,bg="black", font=("Helvetica",font_size) )
widget.pack()

canvas1.create_window(383+x,211.8,window=widget)   



def eskomupdate():
    
    
    #new=2;
    #url="https://www.twitter.com/Eskom_SA";
    #webbrowser.open(url,new=new);
    root2=tk.Tk()
    root2.title("Eskom Updates")
    #canvas=tk.Canvas(root2,height=300,width=600)
    #canvas.pack()
    try:
        position=10
        loc=[]
        indv_words=power_note.get().split()
        no_words=len(indv_words)
        while no_words>position:
            indv_words[position]=indv_words[position]+"\n" #starting inline after position
            if position+1<no_words:
                loc.append(position+1)
            position+=11
        indv=[" "+i for l,i in enumerate(indv_words) if l>0]
        indv.insert(0,indv_words[0])
        for i in loc:
            indv[i]=indv[i].strip()
        new_powernote="".join(indv) #combining all the words together
        
        root2.title("Eskom Updates - "+power_title.get())
        label=tk.Label(root2,text=new_powernote,bg="black",fg="white")
        label.pack()
    except:
        label=tk.Label(root2,text=o,bg="black",fg="white")
        label.pack()
        
    #canvas.create_text(150,150,text=power_note.get().replace(".",".\n"),justify="center")
    new=2;
    url="https://www.twitter.com/Eskom_SA";
    webbrowser.open(url,new=new);
    root2.mainloop()
    
    
   

button2=HoverButton(root,text= "Eskom updates",fg="white",bg="black",activebackground="#3333FF",activeforeground="white",width=13, command= eskomupdate , bd=0 )

canvas1.create_window(340+x,240, window= button2)   




def twitter():
    new=2;
    url="https://www.twitter.com/ChilledFridays";
    webbrowser.open(url,new=new); 
    
button3=HoverButton(root,fg="white",bg="white",activebackground="#00acee",width=35,height=30, command= twitter , bd=0 )
canvas1.create_window(300+x,340, window= button3)   

Image5=tk.PhotoImage(master=canvas1,file="ttw.png").subsample(22)
canvas1.create_image(300+x,340,image=Image5)
button3.config(image=Image5,compound=CENTER,)

def facebook():
    new=2;
    url="https://www.facebook.com/ChilledFridays";
    webbrowser.open(url,new=new); 
    
button4=HoverButton(root,fg="white",bg="white",activebackground="#3B5998",width=45,height=37, command= facebook , bd=0 )
canvas1.create_window(375+x,340, window= button4)

Image3=tk.PhotoImage(master=canvas1,file="black-facebook.png").subsample(22)
canvas1.create_image(375+x,340,image=Image3)
button4.config(image=Image3,compound=CENTER,)

def instagram():
    new=2;
    url="https://www.instagram.com/ChilledFridays";
    webbrowser.open(url,new=new); 
    
button5=HoverButton(root,fg="white",bg="white",activebackground="#DD2A7B",width=45,height=37, command= instagram , bd=0 )
canvas1.create_window(100+x,340, window= button5)

Image4=tk.PhotoImage(master=canvas1,file="instagram-12.png").subsample(16)
canvas1.create_image(100+x,340,image=Image4)
button5.config(image=Image4,compound=CENTER,)

def youtube():
    new=2;
    url="https://www.youtube.com/channel/UC_Ujffi2OBwqqPd74s2YCEQ";
    webbrowser.open(url,new=new); 
    
button6=HoverButton(root,fg="white",bg="white",activebackground="#FF0000",width=42,height=37, command= youtube , bd=0 )
canvas1.create_window(25+x,340, window= button6)

Image2=tk.PhotoImage(master=canvas1,file="transp youtube.png").subsample(15)
canvas1.create_image(25+x,340,image=Image2)
button6.config(image=Image2,compound=CENTER,)

def hearthis():
    new=2;
    url="https://www.hearthis.at/chilled-fridays";
    webbrowser.open(url,new=new); 
    
button7=HoverButton(root,fg="white",bg="white",activebackground="#1DB954", width=57,height=32,command= hearthis , bd=0 )
canvas1.create_window(200+x,290, window= button7)

Image6=tk.PhotoImage(master=canvas1,file="turntables.png").subsample(8)
canvas1.create_image(200+x,300,image=Image6)
button7.config(image=Image6,compound=RIGHT)


Image7=tk.PhotoImage(master=canvas1,file="BAR.png").subsample(2,3)

canvas1.create_image(341+x,256,image=Image7)

Image8=tk.PhotoImage(master=canvas1,file="BAR.png").subsample(4,3)

canvas1.create_image(200.49999999999996+x,256,image=Image8)

Image9=tk.PhotoImage(master=canvas1,file="BAR.png").subsample(4,3)

canvas1.create_image(199.5+x,256,image=Image9)


Image10=tk.PhotoImage(master=canvas1,file="rounded reactangles.png").subsample(1,3)


canvas1.create_image(193+x,160,image=Image10)

var=tk.IntVar()
check=tk.Checkbutton(master=root,bg= "black",activebackground="black",variable=var)

canvas1.create_window(150+x,285,window=check)
canvas1.create_text(105+x,285, text="Tomorrow :",fill="white")

#option.config(image=Image10)
#canvas1.create_image(200.5,176,image=lab)
#option2.config(image=Image8,compound=CENTER)

execute() # show the schedule

root.mainloop()
try:
    with open('last_location.txt','w') as prv_loc:
        prv_loc.write(clicked.get())
        prv_loc.close()
    
except:
    pass  