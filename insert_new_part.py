import win32com.client as Client
import pythoncom

sw = Client.Dispatch("SldWorks.Application")
number = input("Number: ")
xl = Client.Dispatch('Excel.Application')
name = xl.Worksheets(1).Range('A:A').Find(int(number)).Cells(1,2).Value
#name = input("Name: ")
path = "E:\Data\\2020 Documents\Subsystems\Chassis\Cartesian\\"+ number+"_"+name+".SLDPRT"
part = sw.ActiveDoc
arg = Client.VARIANT(pythoncom.VT_I4|pythoncom.VT_BYREF,0)
if not part.SaveToFile3(path,2,2,False,'',arg,arg):
	print("Failure")
else:
	sw.CloseDoc("")
	print("Created file "+ number+"_"+name+".SLDPRT")