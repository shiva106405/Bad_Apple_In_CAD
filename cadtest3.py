import pythoncom
import win32com.client

def cad_run():
    global mp1,doc1,wincad
    wincad=win32com.client.Dispatch("AutoCAD.Application")
    doc1=wincad.ActiveDocument
    mp1=doc1.ModelSpace

#以下为cad接口要求的python数据类型转换函数：
def vtPnt(x,y,z=0):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8,(x,y,z))
def vtObject(obj):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_DISPATCH,obj)
def vtFloat(list):
    return  win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8,list)
def vtInt(list):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_I2,list)
def vtVariant(list):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_VARIANT,list)

class cad_line:
    def __init__(self,start_pt,end_pt):
        self.first_pt=start_pt
        self.second_pt=end_pt
    def line_draw(self):
        reval1=mp1.Addline(self.first_pt,self.second_pt)

