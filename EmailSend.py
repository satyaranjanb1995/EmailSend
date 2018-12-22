#!/usr/bin/env python
# coding: utf-8

# In[1]:


import smtplib


# In[2]:


smtpObj = smtplib.SMTP('smtp-mail.outlook.com', 587)


# In[3]:


type(smtpObj)


# In[4]:


smtpObj.ehlo()


# In[5]:


smtpObj.starttls()


# In[6]:


smtpObj.login('satyaranjanbehera735@outlook.com','satya12345',initial_response_ok=True)


# In[7]:


def prompt(prompt):
    return input(prompt).strip()

fromaddr = prompt("From: ")
toaddrs  = prompt("To: ").split()
subject = prompt("Subject: ")
print("Enter message, end with ^D (Unix) or ^Z (Windows):")

# Add the From: and To: headers at the start!
msg = subject + "\n" + ("From: %s\r\nTo: %s\r\n\r\n"
       % (fromaddr, ", ".join(toaddrs)))
while True:
    try:
        line = input()
    except EOFError:
        break
    if not line:
        break
    msg = msg + line

print("Message length is", len(msg))


# In[8]:


smtpObj.sendmail(fromaddr, toaddrs, msg)


# In[9]:


smtpObj.quit()

