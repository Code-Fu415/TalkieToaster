from tkinter import *
import win32com.client as win32 #imports Windows module
import openpyxl as oxl

import tkinter
from tkinter import filedialog

root = tkinter.Tk()
root.minsize(600, 300)
root.title("Talkie Toaster v1.0")
path=""
mailto=""

def file_open():
  file = filedialog.askopenfile(parent=root,mode='rb',title='Choose a file')
  if file != None:
    data = file.read()
    file.close()
    x = len(data)
    print(f"I got {x} bytes from this file.")
    print(file.name)
    filepath.config(text=file.name)
    global path
    path=file.name
    loaded.config(text="")
    created.config(text="")
    global mailto
    mailto = ""


      
  
def load_wb():
  global path
  wb=oxl.load_workbook(path)
  sheet = wb['Sheet1']
  x = 2
  first_name = sheet[f"B{x}"].value
  last_name = sheet[f"C{x}"].value
  emails = []
  while first_name != None:
    first_name = sheet[f"B{x}"].value
    if first_name == None:
      break
    last_name = sheet[f"C{x}"].value
    email = f"{first_name}.{last_name}@berkeleylights.com"
    emails.append(email)
    x += 1
  recipients = ""
  for i in emails:
    recipients = recipients+i+"; "
  global mailto
  mailto=recipients
  loaded.config(text="List Loaded")  

def create_email():
  global mailto
  outlook = win32.Dispatch('outlook.application') #launches Outlook
  mail = outlook.CreateItem(0) #creates outlook email
  mail.To = mailto
  mail.HTMLBody = """<p style="font-size:15px">Stuart Yee<br>Operations<br>Berkeley Lights, Inc<br>
5858 Horton St., Suite 320<br>
Emeryville, CA 94608<br>
DIRECT: (510) 985-3104<br></p>

<p style="color:red">AUTOMATED MESSAGE:</p> If you feel you have received this message in error, please contact <a href="mailto:stuart.yee@berkeleylights.com">stuart.yee@berkeleylights.com</a>

<p style="font-size:12px">CONFIDENTIALITY NOTICE:
This email transmission, and any documents, files or previous email messages attached to it may contain confidential information that is legally privileged or is otherwise subject to certain non-disclosure restrictions. If you are not the intended recipient, or a person responsible for delivering it to the intended recipient, you are hereby notified that any disclosure, copying, distribution or use of any of the information contained in or attached to this transmission is strictly prohibited. If you have received this transmission in error, please immediately notify us by return email or by telephone at (510)  858-2855, and destroy the original transmission and its attachments without reading or saving in any manner. Thank you.</p>"""
  mail.Display(False)
  created.config(text="Check Outlook, email generated!")

  
      
      
opnbtn = tkinter.Button(master=root, text="Select File", command=file_open)
opnbtn.grid(row=1, sticky=W)

filepath = tkinter.Label(bg="white", text="", justify=LEFT, width=150)
filepath.grid(row=1, column=1, columnspan=2)

instructions = tkinter.Label(text="Howdy Doodly Doo! \nInstructions: create a spreadsheet and on Sheet1 put first names on column B starting at B2, and last names on column C", justify=LEFT)
instructions.grid(row=0, column=0, columnspan=3, pady=5)

loadbtn = tkinter.Button(master=root, text="Load Selected File", command=load_wb)  
loadbtn.grid(row=2, sticky=W)

emailbtn= tkinter.Button(master=root, text="Create Email", command=create_email)
emailbtn.grid(row=3, sticky=W)

loaded = tkinter.Label(bg="white", text="", justify=LEFT, width=150)
loaded.grid(row=2, column=1)
created = tkinter.Label(bg="white", text="", justify=LEFT, width=150)
created.grid(row=3, column=1)

canvas = tkinter.Canvas()
canvas.grid(row=4, column=1)
img = tkinter.PhotoImage(file="talkie.gif")
canvas.create_image(20, 20, anchor="nw", image=img)

root.mainloop()