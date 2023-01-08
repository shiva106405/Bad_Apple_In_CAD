import pythoncom
import win32com.client

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