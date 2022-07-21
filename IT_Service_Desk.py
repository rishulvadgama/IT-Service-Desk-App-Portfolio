from tkinter import *
import os
import subprocess
from tkinter import ttk
import webbrowser
from ctypes import windll

windll.shcore.SetProcessDpiAwareness(1)


root=Tk()
root.title(" IT Service Desk")
root.iconbitmap('mclaren.ico')
#root.geometry("800x550")
root['background']='#F48500'

def ExecuteFunction(process, result):
    result.delete('1.0', END)
    p = subprocess.Popen(process,shell=True, stdout=subprocess.PIPE,stderr=subprocess.PIPE, universal_newlines=True)
    if p.stdout:
        for line in p.stdout:
            result.insert(END, line)
    if p.stderr:
        for line in p.stderr:
            result.insert(END, line)
    result.config(state=DISABLED)

def ConnectionStatus():   
    response = os.popen(f"ping tagmclarengroup.com").read()
    if "Received = 4" in response:
        ConnectionText.insert(END, 'You are Online')
    else:
        ConnectionText.insert(END, 'You are Offline')

def AVA():
    webbrowser.open_new("https://mclaren.service-now.com/sp")

logo = PhotoImage(file="McLaren_Logo_1.png")
logo_label = Label(image=logo,borderwidth=0, highlightthickness=0)
logo_label.grid(row=0, column=0)

Title=Label(root, text="IT Service Desk", bg="#F48500", justify=CENTER)
Title.config(font=('McLaren Bespoke Bold',16))

Space=Label(root, text=" ", bg="#F48500")
Space2=Label(root, text=" ", bg="#F48500")
Space3=Label(root, text=" ", bg="#F48500")
Space3.config(font=('McLaren Bespoke',2))

HostnameLabel=Label(root, text="   Hostname:    ", bg="#F48500", anchor="w")
HostnameLabel.config(font=('McLaren Bespoke',11))
HostnameText=Text(root, width=80, height=1, bg="white", borderwidth=0)

IpaddressLabel=Label(root, text="   IPv4 Address:    ", bg="#F48500", anchor="w")
IpaddressLabel.config(font=('McLaren Bespoke', 11))
IpaddressText=Text(root,width=80, height=1, bg="white", borderwidth=0)

UserLabel=Label(root, text="   Username:", bg="#F48500", anchor="w")
UserLabel.config(font=('McLaren Bespoke', 11))
UserText=Text(root, width=80, height=1, bg="white", borderwidth=0)

BiosLabel=Label(root, text="   BIOS Version:", bg="#F48500", anchor="w")
BiosText=Text(root, width=80, height=1, bg="white", borderwidth=0)
BiosLabel.config(font=('McLaren Bespoke', 11))

BuildLabel=Label(root, text="   Build Number:", bg="#F48500")
BuildLabel.config(font=('McLaren Bespoke', 11))
BuildText=Text(root, width=80, height=1, bg="white", borderwidth=0)

ModelLabel=Label(root, text="   Model No:", bg="#F48500")
ModelLabel.config(font=('McLaren Bespoke', 11))
ModelText=Text(root, width=80, height=1, bg="white", borderwidth=0)

DomainLabel=Label(root, text="   Domain:", bg="#F48500")
DomainLabel.config(font=('McLaren Bespoke', 11))
DomainText=Text(root, width=80, height=1, bg="white", borderwidth=0)

SerialLabel=Label(root, text="   Serial Number: ", bg="#F48500")
SerialLabel.config(font=('McLaren Bespoke', 11))
SerialText=Text(root, width=80, height=1, bg="white", borderwidth=0)

ConnectionLabel=Label(root, text="   Status: ", bg="#F48500")
ConnectionLabel.config(font=('McLaren Bespoke', 11))
ConnectionText=Text(root, width=80, height=1, bg="white", borderwidth=0)

MacLabel=Label(root, text="   MAC Address: ", bg="#F48500")
MacLabel.config(font=('McLaren Bespoke',11))
MacText=Text(root, width=80, height=4, bg="white", borderwidth=0)

ContactInfo1=Label(root, text="Have you got an IT issue?", bg="#F48500")
ContactInfo1.config(font=('McLaren Bespoke', 14))

OpenAVA=ttk.Button(root, text="Open AVA", command=AVA)

Title.grid(row=0, column=0, columnspan=2)

Space.grid(row=1, column=0)

ConnectionLabel.grid(row=2, column=0, sticky=W)
ConnectionText.grid(row=2, column=1, padx=25)

HostnameLabel.grid(row=3, column=0, sticky=W)
HostnameText.grid(row=3, column=1, columnspan=2)

DomainLabel.grid(row=4, column=0, sticky=W)
DomainText.grid(row=4, column=1)

IpaddressLabel.grid(row=5, column=0, sticky=W)
IpaddressText.grid(row=5, column=1)

UserLabel.grid(row=7, column=0, sticky=W)
UserText.grid(row=7, column=1)

BiosLabel.grid(row=8,column=0, sticky=W)
BiosText.grid(row=8, column=1)

BuildLabel.grid(row=9,column=0, sticky=W)
BuildText.grid(row=9, column=1)


ModelLabel.grid(row=10, column=0, sticky=W)
ModelText.grid(row=10, column=1)

SerialLabel.grid(row=11, column=0, sticky=W)
SerialText.grid(row=11, column=1)

MacLabel.grid(row=6,column=0, sticky=NW)
MacText.grid(row=6, column=1, pady=5)

Space2.grid(row=12, column=0)

ContactInfo1.grid(row=13, column=0, columnspan=2)

Space3.grid(row=14, column=0)

OpenAVA.grid(row=15, column=0, columnspan=2, ipadx=30, ipady=15, pady=25)

ExecuteFunction('hostname', HostnameText)
ExecuteFunction("""for /f "tokens=2 delims=[]" %a in ('ping -n 1 -4 ""') do @echo %a""", IpaddressText)
ExecuteFunction("whoami", UserText)
ExecuteFunction('wmic os get buildnumber | findstr /r /v "Build"', BuildText)
ExecuteFunction('WMIC CSPRODUCT GET NAME | findstr /r /v "Name"', ModelText)
ExecuteFunction('wmic computersystem get domain | findstr /r /v "Domain"', DomainText)
ExecuteFunction('wmic bios get serialnumber | findstr /r /v "Serial"', SerialText)
ConnectionStatus()
ExecuteFunction('getmac | findstr /r /v "=" | findstr /r /v "name" | findstr /r /v "N/A | findstr /r /v "Physical" | findstr /r /v "Media', MacText)
ExecuteFunction('wmic bios get smbiosbiosversion | findstr /r /v "SMBIOSBIOSVersion"', BiosText)

root.mainloop()