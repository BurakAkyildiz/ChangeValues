' ----------------------------------------------
' Script Recorded by ANSYS  Version 2016.2.0
' 13:26:24  Oca 08, 2017
' ----------------------------------------------
Dim oAnsoftApp
Dim oDesktop
Dim oProject
Dim oDesign
Dim oEditor
Dim oModule
Set oAnsoftApp = CreateObject("Ansoft.ElectronicsDesktop")
Set oDesktop = oAnsoftApp.GetAppDesktop()
oDesktop.RestoreWindow
Set oProject = oDesktop.SetActiveProject("NIHAI-5kW-YHJ-12102016")
Set oDesign = oProject.SetActiveDesign("5kW-End")
Set oEditor = oDesign.SetActiveEditor("Machine")
oEditor.ChangeProperty Array("NAME:AllTabs", Array("NAME:Pole", Array("NAME:PropServers",  _
  "Rotor:Pole"), Array("NAME:ChangedProps", Array("NAME:Magnet Thickness", "Value:=", "2mm"))))
oEditor.ChangeProperty Array("NAME:AllTabs", Array("NAME:Pole", Array("NAME:PropServers",  _
  "Rotor:Pole"), Array("NAME:ChangedProps", Array("NAME:Magnet Width", "Value:=", "333333mm"))))
oEditor.ChangeProperty Array("NAME:AllTabs", Array("NAME:Rotor", Array("NAME:PropServers",  _
  "Rotor"), Array("NAME:ChangedProps", Array("NAME:Inner Diameter", "Value:=", "4mm"))))
oProject.Save
oDesign.AnalyzeAll
Set oModule = oDesign.GetModule("ReportSetup")
oModule.CreateReport "XY Plot 1", "RMxprt", "Rectangular Plot",  _
  "Setup1 : Performance", Array("Domain:=", "Angle"), Array("Ang:=", Array("All")), Array("X Component:=",  _
  "Ang", "Y Component:=", Array("Efficiency")), Array()
oModule.DeleteReports Array("XY Plot 1")
oModule.CreateReport "Data Table 1", "RMxprt", "Data Table",  _
  "Setup1 : Performance", Array("Domain:=", "Parameter"), Array("X:=", Array("All")), Array("X Component:=",  _
  "X", "Y Component:=", Array("PowerFactorParameter")), Array()
oModule.CreateReport "Data Table 2", "RMxprt", "Data Table",  _
  "Setup1 : Performance", Array("Domain:=", "Parameter"), Array("X:=", Array("All")), Array("X Component:=",  _
  "X", "Y Component:=", Array("EfficiencyParameter")), Array()
oModule.DeleteReports Array("Data Table 1")
oModule.ExportToFile "Data Table 2", "C:/Users/by/Desktop/TEZ/1/output.txt"
oProject.Save
