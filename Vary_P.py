# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import win32com.client as client
import pythoncom
sw = client.Dispatch("SldWorks.Application")
cosmos = sw.GetAddInObject("SldWorks.Simulation")
cwork = cosmos.CosmosWorks
act_doc = cwork.ActiveDoc
studymgr = act_doc.StudyManager
static = studymgr.GetStudy(0)
lrmgr = static.LoadsAndRestraintsManager
x = 0;
arg1 = client.VARIANT(pythoncom.VT_BYREF|pythoncom.VT_I4,0)
#l= lrmgr.GetLoadsAndRestraints(0,arg1)
part = sw.ActiveDoc
selMgr = part.SelectionManager
print(selMgr.GetSelectionPoint2(1,-1))
#swFeat = selMgr.GetSelectedObject6(1,-1)
#sketch = swFeat.GetSpecificFeature
#points = sketch.GetSketchPoints2
#point = points[0]
#print(point.GetCoords)
xl = client.Dispatch("Excel.Application")
xl.WorkBooks.Add()
xl.Visible = 1
sheet = xl.ActiveSheet
for i in range(0,lrmgr.Count):
    p = [500,1000,2000,10000,30000,60000,100000]
    l = lrmgr.GetLoadsAndRestraints(i,arg1)
    if(l.Type ==1):
        for i in range(len(p)):
            l.PressureBeginEdit()
            l.Value = float(p[i])
            l.PressureEndEdit
            static.RunAnalysis
            res = static.Results
            coords = client.VARIANT(pythoncom.VT_VARIANT|pythoncom.VT_ARRAY,[3.64935207,4.21521187,4.78179598])
            third = client.VARIANT(pythoncom.VT_DISPATCH,None)
            disp = res.GetDisplacementAtPoints(3,1,third,selMgr.GetSelectionPoint2(1,-1),0,arg1)
            xl.Cells(i+1,1).Value = p[i]
            xl.Cells(i+1,2).Value = disp[1]
            print(disp)
xl.Visible = 1