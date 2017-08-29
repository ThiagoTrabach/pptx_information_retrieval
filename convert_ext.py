
# coding: utf-8

# In[1]:

import sys, os, glob  
import win32com.client  
import pptx, docx

def convert_ppt(files, path, formatType = 24):  
    powerpoint = win32com.client.Dispatch("Powerpoint.Application") 
    newname = os.path.splitext(files)[0] 
    deck = powerpoint.Presentations.Open(os.path.join(path, files))
    deck.SaveAs(newname, formatType)
    deck.Close()
    powerpoint.Quit()

def convert_word(files, path, formatType = 12):  
    word = win32com.client.Dispatch("Word.Application") 
    newname = os.path.splitext(files)[0] 
    deck = word.Documents.Open(os.path.join(path, files))
    deck.SaveAs(newname, formatType)
    deck.Close()
    word.Quit()


# In[ ]:



