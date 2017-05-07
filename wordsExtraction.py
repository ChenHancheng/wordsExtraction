#! python3
import re
import win32com
import time
from win32com.client import Dispatch,constants

def get_voca():
    fp = open(r'F:\精读单词总结\a.txt', "r", encoding = 'utf-8')
    content = fp.read()
    s = re.findall(r'@.*//|@.*@', content, re.M)
    return s

def get_spell(voca):
    m=re.findall('@(.*)@',voca)
    if len(m) == 1:
        return m[0]
    else:
        return -1
    
def get_pronun(voca):
    m=re.findall('(/.*/)/',voca)
    if len(m) == 1:
        return m[0]
    elif len(m) == 0:
        return ' '
    else:
        return -1
        
w = win32com.client.Dispatch('Word.Application')
w.Visible = 1
w.DisplayAlerts = 0
olddoc = w.Documents.Open(r"F:\精读单词总结\!template.docx")
olddoc.SaveAs('F:\\精读单词总结\\' + time.strftime('%Y.%m.%d') + r".docx")
olddoc.Close()
# newdoc.Close()
newdoc = w.Documents.Open('F:\\精读单词总结\\' + time.strftime('%Y.%m.%d') + r".docx")
newdoc.Sections[0].Headers[0].Range.Text = 'IELTS 5, TEST 3'
voca = get_voca()
for i in range(len(voca)):
    newdoc.Tables[0].Cell(i+1,1).Range.Text = get_spell(voca[i]) + '\n' +  get_pronun(voca[i])
    newdoc.Tables[0].Cell(i+1,2).Range.Text = '\n'
    newdoc.Tables[0].Rows.Add();
#doc.Close()
#/w.Quit()

