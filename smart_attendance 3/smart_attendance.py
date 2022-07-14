
from tkinter import*
try:
    import tkinter as tk
except:
    import tkinter as tk
import tkinter as tk

from PIL import Image, ImageTk
from datetime import datetime
from time import sleep

from email.mime.multipart import MIMEMultipart  
from email.mime.base import MIMEBase  
from email.mime.text import MIMEText  
from email.utils import formatdate  
from email import encoders
import smtplib,ssl

import face_recognition
import numpy as np
import sqlite3
import csv
import pandas as pd
import pickle 
import ctypes
import time
import cv2
import os
 
ctypes.windll.shcore.SetProcessDpiAwareness(1) # it increase the window clearity
root = Tk()
photo = PhotoImage(file = "icon/icon.png")
root.iconphoto(False, photo) # set the icon to the window
root.geometry("1000x800+70+70")
root.resizable(True,False)
root.title('Smart Attendance System')
root.bind("<Escape>", exit) #press escape to exit window
root.configure(bg='gray70')



lbl1=Label(root,text ="SMART ATTENDANCE SYSTEM",fg ='PaleGreen1' , font =("arial", 50),bg='black')
lbl1.place(x=0,y=0,relwidth=1)
#creating frame
frame1 = tk.Frame(root, bg="#75968b",bd=15, width=130,relief=RIDGE)
frame1.place(relx=0.07, rely=0.17, relwidth=0.39, relheight=0.30)

frame2 = tk.Frame(root, bg="#75968b",bd=15, width=130,relief=RIDGE)
frame2.place(relx=0.55, rely=0.17, relwidth=0.39, relheight=0.20)

frame3 = tk.Frame(root, bg="#75968b",bd=15, width=130,relief=RIDGE)
frame3.place(relx=0.55, rely=0.60, relwidth=0.39, relheight=0.35)

frame4 = tk.Frame(root, bg="#75968b",bd=15, width=130,relief=RIDGE)
frame4.place(relx=0.07, rely=0.50, relwidth=0.39, relheight=0.45)

frame5 = tk.Frame(root, bg="#75968b",bd=15, width=130,relief=RIDGE)
frame5.place(relx=0.55, rely=0.38, relwidth=0.39, relheight=0.20)

my_img = ImageTk.PhotoImage(Image.open('images.png'))
my_label = Label( image= my_img)
my_label.place(x=152,y=185)



lbl2=Label(root,text ="   ENROLL STUDENT   ",fg ='#7efcd0' , font =("Arial Black", 24),bg='black')
lbl2.place(x=370,y=185)

lbl3=Label(root,text ="   Click here to Start attendance      ",fg ='#7efcd0' , font =("times new roman", 30),bg='black')
lbl3.place(x=1076,y=185)

lbl16=Label(root,text ="   For Reset    ",fg ='#7efcd0' , font =("times new roman", 30),bg='black')
lbl16.place(x=1076,y=850)


########################### mail ###########################################################
def send_an_email():
            me = 'enter ur email here'     # enter your email id
            toaddr = mail_id                  # email id of person to send the mail      
            subject = "college authority"     # write Subject
                      
            msg = MIMEMultipart()  
            msg['Subject'] = subject  
            msg['From'] = me  
            msg['To'] = toaddr  
            msg.preamble = "test "   
            msg.attach(MIMEText("Attendance"))
                      
            part = MIMEBase('application', "octet-stream")  
            part.set_payload(open("attendance_data.xlsx", "rb").read())  
            encoders.encode_base64(part)  
            part.add_header('Content-Disposition', 'attachment; filename="attendance_data.xlsx"')   # File name and format name
            msg.attach(part)  
                      
             
            s = smtplib.SMTP('smtp.gmail.com', 587)  # Protocol
            s.ehlo()  
            s.starttls()  
            s.ehlo()  
            s.login(user = 'sayalitadge123@gmail.com', password = '#######')  # User id & password
            s.send_message(msg)   
            s.sendmail(me, toaddr, msg.as_string())  
            s.quit()
            
            btn9.destroy()
                
                         
            
 
###############################################################################################
                    
global btn8, btn9, password
def setTextInput1():   # save email address
    global mail_id,lbl4,btn9,lbl14
        
    if (mailid.get() == ""):
            print()
            
    else:    
        mail_id=mailid.get()
        mailid.delete(0, END)

        # lbl14=Label(root,text ="Send the mail",fg ='red3' , font =("times new roman", 30))
        # lbl14.place(x=1265,y=880)

        btn9 = Button(root, text = ' SEND ',bg='alice blue', bd = '10',width=15,height=2,font =("Arial", 10,'bold'),command = send_an_email) # it create the send button to send the mail
        btn9.place(x=1600, y=730)
        
        

lbl11=Label(root,text ="                 Send database               ",fg ='#7efcd0' , font =("times new roman", 30),bg='black')
lbl11.place(x=1074,y=615)

lbl14=Label(root,text ="  Enter email id : ",fg ='#7efcd0' , font =("times new roman", 25),bg='black')
lbl14.place(x=1076,y=680)

large_font=('Verdana',20)
mailid=Entry(root,width=15,font=large_font,fg = 'DarkOrange3') # this create entry box to write sender mail     
# mailid.grid(row=0, column=1)
mailid.place(x=1400,y=680)
mailid.focus()
    
btn3 = Button(root, text = ' SAVE ',bg='alice blue', bd = '10',width=13,height=2,font =("Arial", 10,'bold'),command = setTextInput1) # this create submit button of entry box, it submit the mail
btn3.place(x=1400, y=730)



##################################### #  start attendace ##################
def Start():
        face_encoding='face_encoding/'

        known_face_encodings = []
        known_face_names = []

        video_capture = cv2.VideoCapture(0)

        for filename in os.listdir(face_encoding):
                    known_face_names.append(filename[:-4])
                    
                    with open (face_encoding+filename, 'rb') as fp: # call the face encoding of the student inside face_encoding folder
                                
                                known_face_encodings.append(pickle.load(fp))
                                
                        
   
                        
                        
                    

        # Initialize some variables
        face_locations = []
        face_encodings = []
        face_names = []
        process_this_frame = True
        #known_face_encodings = pickle.load(open('encod.pickle', 'rb'))
        global name
        try:
                while True:
                    # Grab a single frame of video
                    ret, frame = video_capture.read()

                    # Resize frame of video to 1/4 size for faster face recognition processing
                    small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)

                    # Convert the image from BGR color (which OpenCV uses) to RGB color (which face_recognition uses)
                    rgb_small_frame = small_frame[:, :, ::-1]

                    # Only process every other frame of video to save time
                    if process_this_frame:
                        # Find all the faces and face encodings in the current frame of video
                        face_locations = face_recognition.face_locations(rgb_small_frame)
                        face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)

                        face_names = []
                        name = "Unknown"
                        
                        for face_encoding in face_encodings:
                            # See if the face is a match for the known face(s)
                            
                            matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
                            name = "Unknown"
                              
                            # Or instead, use the known face with the smallest distance to the new face
                            face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)
                            best_match_index = np.argmin(face_distances)
                            if matches[best_match_index]:
                                name = known_face_names[best_match_index]

                                conn = sqlite3.connect('attendance_database.db')                               
                                curs=conn.cursor()
                                curs.execute("SELECT college_id FROM student_table WHERE name ='%s' "%name)
                                records3 = curs.fetchall()
                                x=["Empty","Empty","Empty"]
                                for x in records3:
                                    print(x[0])

                                conn.commit()
                                curs.close()

                                curs=conn.cursor()
                                curs.execute("SELECT name FROM present_student_table WHERE name ='%s' "%name)
                                records3 = curs.fetchone()
                        
                                if records3:
                                    print()
                                else:                               
                                    now1=datetime.now()
                                    current_time=now1.strftime("%d-%m-%Y %I:%M%p")
                                    conn = sqlite3.connect('attendance_database.db')                                 
                                    curs = conn.cursor()
                                    curs.execute('INSERT INTO present_student_table(college_id, name,present_date_time) values(?,?,?)',( x[0],name,current_time)) # insert the present student in the database
                                    conn.commit()
                                    curs.close()                               
                                            
                                    face_names.append(name)

                    process_this_frame = not process_this_frame

                    # Display the results
                    for (top, right, bottom, left), name in zip(face_locations, face_names):
                        # Scale back up face locations since the frame we detected in was scaled to 1/4 size
                        top *= 4
                        right *= 4
                        bottom *= 4
                        left *= 4

                        # Draw a box around the face
                        cv2.rectangle(frame, (left, top), (right, bottom), (255, 0, 0), 4)

                        # Draw a label with a name below the face
                        cv2.rectangle(frame, (left, bottom - 35), (right, bottom), (0, 0, 255), cv2.FILLED)
                        font = cv2.FONT_HERSHEY_DUPLEX
                        cv2.putText(frame, name, (left + 6, bottom - 6), font, 1.0, (255, 255, 255), 1)

                    # Display the resulting image
                    #cv2.resizeWindow('a', 500, 500)
                    cv2.imshow('Video', frame)
                    cv2.moveWindow('Video', 1200, 160)

                    # Hit 'q' on the keyboard to quit!
                    if cv2.waitKey(1) & 0xFF == ord('q'):
                    #if n == 'p':
                     break

        # Release handle to the webcam
        except KeyboardInterrupt:
          video_capture.release()
          cv2.destroyAllWindows()


################################## enroll student information ############################################################


          
def Capture():
    global lbl4,lbl5,lbl6,lbl7,lbl8,lbl9
    
    main = Tk()
    # photo = PhotoImage(file = "icon/icon.png")
    # main.iconphoto(False, photo) # set the icon to the window
    main.geometry("1000x800+70+70")
    main.resizable(True,False)
    main.title('Smart Attendance System(author : Neeraj)')
    main.bind("<Escape>", exit) #press escape to exit window
    main.configure(bg='gray70')

    frame6 = tk.Frame(main, bg="#889e75",bd=15, width=130,relief=RIDGE)
    frame6.place(relx=0.23, rely=0.20, relwidth=0.50, relheight=0.70)

    lblx=Label(main,text ="SMART ATTENDANCE SYSTEM",fg ='PaleGreen1' , font =("arial", 50),bg='black')
    lblx.place(x=0,y=0,relwidth=1)
    
    lbl4=Label(main,text =" Enter name ",fg ='black' , font =("times new roman", 23,'bold'),bg='#889e75')
    lbl4.place(x=500,y=300)
       
    lbl5=Label(main,text =" Enter college id ",fg ='black' , font =("times new roman", 23,'bold'),bg='#889e75')
    lbl5.place(x=500,y=370)
   
    lbl6=Label(main,text =" Enter batch  ",fg ='black' , font =("times new roman", 23,'bold'),bg='#889e75')
    lbl6.place(x=500,y=440)
   
    lbl7=Label(main,text =" Enter department ",fg ='black' , font =("times new roman", 23,'bold'),bg='#889e75')
    lbl7.place(x=500,y=510)

    lbl8=Label(main,text =" Capture the image ",fg ='black' , font =("times new roman", 23,'bold'),bg='#889e75')
    lbl8.place(x=500,y=580)
            

    global name_id, c , d, j, lbl9, lbl10
    def setTextInput():
        global name_id, c, d, j, lbl9, lbl10
        
        if (name.get() == "" and college_id.get() == "" and batch.get() == "" and department.get() == ""):
            print()
        
        else:    
            name_id = name.get()
            c = college_id.get()
            d = batch.get()
            j = department.get()
            
            name.delete(0, END)
            college_id.delete(0, END)
            batch.delete(0, END)
            department.delete(0, END)

            conn = sqlite3.connect('attendance_database.db')   
            curs = conn.cursor()
            curs.execute('INSERT INTO student_table(college_id,name,batch,department) values(?,?,?,?)',( c,name_id,d,j)) # this query insert student information to the student_table
            conn.commit()
                                            


    global name
    large_font=('Verdana',20)
    name=Entry(main,width=19,font=large_font,bd=5,fg = 'DarkOrange3') # this create entry box to write name     
    # name.grid(row=0, column=1)
    name.place(x=900,y=300)
    name.focus_set()
    
    global college_id
    college_id=Entry(main,width=19,font=large_font,bd=5,fg = 'DarkOrange3')# this create entry box to write college id    
    # college_id .grid(row=1, column=0)
    college_id .place(x=900,y=370)

    global batch
    batch=Entry(main,width=19,font=large_font,bd=5,fg = 'DarkOrange3')# this create entry box to write batch      
    # batch.grid(row=2, column=0)
    batch.place(x=900,y=440)
   
    global department
    department=Entry(main,width=19,font=large_font,bd=5,fg = 'DarkOrange3')# this create entry box to write department     
    # department.grid(row=3, column=0)
    department.place(x=900,y=510)
    
    global btn2
    btn2 = Button(main, text = '   SAVE!   ',bg='black',fg='white', bd = '10',width=15,height=2,font =("Arial", 10,'bold'),command = setTextInput)# this button save student information 
    btn2.place(x=900,y=680)
    
    def Run():
        name_id = name.get()
        #timer = int() # timer
        cap = cv2.VideoCapture(0) 
        while True: 
                            ret, img = cap.read() 
                            cv2.imshow('a', img)  
                            prev = time.time()
                            timer = int() # timer

                            while timer >= 0: 
                                            ret, img = cap.read() 
                                            font = cv2.FONT_HERSHEY_SIMPLEX 
                                            cv2.putText(img, str(timer), (200, 250), font, 7, (0, 255, 255), 4, cv2.LINE_AA) 
                                            cv2.imshow('a', img)
                                            cv2.moveWindow('a', 1200, 160)
                                            cv2.waitKey(125) 
                                            cur = time.time()
                                            
                                            if cur-prev >= 1: 
                                                    prev = cur 
                                                    timer = timer-1

                                            else: 
                                                ret, img = cap.read() 
                                            cv2.imshow('a', img)
                                            cv2.moveWindow('a', 1200, 160)
                                            cv2.waitKey(100) 
                                            cv2.imwrite('clicked_photo/'+ name_id +'.jpg', img)
                                            
                                            enroll_encoding=[]
                                            img_ = face_recognition.load_image_file('clicked_photo/'+ name_id +'.jpg')
                                            try:
                                               img_encoding = face_recognition.face_encodings(img_)[0]
                                               enroll_encoding.append(img_encoding)
                                            
                                               f=open('face_encoding/'+name_id+'.txt','w+')
                                               f.close()
                                               with open('face_encoding/'+name_id+'.txt','wb') as fp:
                                                    pickle.dump(img_encoding, fp) #two argument required . only 1 given
                                            

                                            
                                            except IndexError as e:
                                               print("dlib face detector couldn't detect a face in the image you passed in")
                                               print(e)
                                            conn = sqlite3.connect('attendance_database.db')   
                                            curs = conn.cursor()
                                            curs.execute('INSERT INTO student_table(college_id, name,batch,department) values(?,?,?,?)',( c,name_id,d,j)) # this query insert student information to the student_table
                                            conn.commit()
                                            cap.release()

                                            ###################### this destroy the labels and buttons ########
                                            department.destroy()
                                            batch.destroy()
                                            college_id.destroy()
                                            name.destroy()
                                            btn2.destroy()
                                            btn5.destroy()
                                            
                                            lbl4.destroy()
                                            lbl5.destroy()
                                            lbl6.destroy()
                                            lbl7.destroy()
                                            lbl8.destroy()
                                            ##################################################################
                                            
                                            
                                            cv2.destroyAllWindows()
    global btn5                                      
    btn5 = Button(main, text = '     CAPTURE !    ',bg='alice blue', bd = '10', width=15,height=2,font =("Arial", 10,'bold'),command = Run) # this button help to enroll the face of the student
    btn5.place(x=900,y=580)

    main.state('zoomed')
    main.mainloop()

btn = Button(root, text = 'ENROLL !',bg='alice blue',fg='black', bd = '10', width=25,height=2,font =("Arial", 10,'bold'),command = Capture) # this button enroll the student information to database 
btn.place(x=450, y=270)

btn1 = Button(root, text = '    START !    ',bg='alice blue', bd = '10', width=25,height=2,font =("Arial", 10,'bold'),command = Start)  # this button start face recognition to take attendance 
btn1.place(x=1300, y=270)


############################## generate database ###################
def setTextInput2():
    global lbl18
    lbl18=Label(root,text =" Enter password ",fg ='black' , font =("times new roman", 23,'bold'),bg='#75968b')
    lbl18.place(x=150,y=700)
    
    global pass_word , password
    password=Entry(root,width=15,font=large_font,fg = 'DarkOrange3') #this create the entry box to enter the password,so i create the database   
    # password.grid(row=2, column=0)
    password.place(x=450,y=700)
    password.focus()
 
    def setTextInput3():
        pass_word = password.get()
        if pass_word == "sayali123": # from here you can change your password
            conn = sqlite3.connect('attendance_database.db')
            curs = conn.cursor()
            curs.execute("CREATE TABLE student_table(college_id int(20) not null,name text not null, batch text not null,department text not null,primary key(college_id),unique(college_id))")
            curs.execute("CREATE TABLE present_student_table(college_id int(20) not null,name TEXT not null,present_date_time datetime default CURRENT_TIMESTAMP)")
            conn.commit()
            password.destroy()
            lbl18.destroy()
            btn8.destroy()
        
        
    global btn8    
    btn8 = Button(root, text = 'Enter PSW',bg='alice blue', bd = '10',font =("Arial", 10,'bold'),command = setTextInput3)#press button causes password insert
    btn8.place(x=450, y=770)
######################################################################
                        
                   
lbl17=Label(root,text ="              Generate database             ",fg ='#7efcd0' , font =("times new roman", 30),bg='black')
lbl17.place(x=146,y=510)
btn7 = Button(root, text = 'GENERATE DB',bg='alice blue', bd = '10',width=25,height=2,font =("Arial", 10,'bold'),command = setTextInput2)
btn7.place(x=380, y=590)


############################ destroy buttons and labels ######
def reset():
    department.destroy()
    batch.destroy()
    college_id.destroy()
    name.destroy()
    lbl4.destroy()
    lbl5.destroy()
    lbl6.destroy()
    lbl7.destroy()
    lbl8.destroy()
    btn2.destroy()
    btn5.destroy()
    
    lbl18.destroy()
    btn8.destroy()                
    
    password.destroy()
    btn9.destroy()
    
    
    
    
btn6 = Button(root, text = 'RESET ALL !',bg='alice blue', bd = '10', width=25,height=2,font =("Arial", 10,'bold'),command = reset) # this button destroy buttons and labels
btn6.place(x=1500, y=850)
##########################  export to excel   ################################
def save_db():
    conn = sqlite3.connect('attendance_database.db') 
    writer = pd.ExcelWriter('attendance_data.xlsx')
    df= pd.read_sql("SELECT name FROM sqlite_master WHERE type='table'",conn)
    df['name']
    for table_name in df['name']:
        sheet_name=table_name
        SQL = "SELECT * from "+sheet_name
        dft = pd.read_sql(SQL , conn)
        # load data into excel file
        dft.to_excel(writer , sheet_name=sheet_name , index = False)
    writer.save()
    
    

      
    
      
btn10 = Button(root, text = 'EXPORT TO EXCEL',bg='alice blue', bd = '10', width=25,height=2,font =("Arial", 10,'bold'),command = save_db) # this button destroy buttons and labels
btn10.place(x=1300, y=435)    


########################## it focus the entry box when i press the arrow button #### 
def previous_widget(event):
        event.widget.tk_focusPrev().focus()
        return "break"
root.bind_class("Entry", "<Up>",previous_widget)

def next_widget(event):
        event.widget.tk_focusNext().focus()
        return "break"
root.bind_class("Entry", "<Down>",next_widget)
###################################################################################
root.state('zoomed')
root.mainloop()
