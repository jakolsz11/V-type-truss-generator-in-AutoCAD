import win32com.client
import pythoncom

def APoint(x, y, z = 0):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

def ADouble(xyz):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (xyz))

def variants(object):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, (object))