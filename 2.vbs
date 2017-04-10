' ----------------------------------------------
' Script Recorded by ANSYS Electronics Desktop Version 2016.2.0
' 23:36:11  Nis 09, 2017
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
Set oProject = oDesktop.SetActiveProject("N04-100-110-00003KW-1500RPM")
Set oDesign = oProject.SetActiveDesign("N04-100-110-00003KW-1500RPM")
Set oModule = oDesign.GetModule("Optimetrics")
oModule.EditSetup "ParametricSetup1", Array("NAME:ParametricSetup1", "IsEnabled:=",  _
  true, Array("NAME:StartingPoint"), "Sim. Setups:=", Array("Setup1"), Array("NAME:Sweeps", Array("NAME:SweepDefinition", "Variable:=",  _
  "e", "Data:=", "0.1mm", "OffsetF1:=", false, "Synchronize:=", 0), Array("NAME:SweepDefinition", "Variable:=",  _
  "o", "Data:=", "1mm", "OffsetF1:=", false, "Synchronize:=", 0), Array("NAME:SweepDefinition", "Variable:=",  _
  "mt", "Data:=", "1mm", "OffsetF1:=", false, "Synchronize:=", 0)), Array("NAME:Sweep Operations"), Array("NAME:Goals", Array("NAME:Goal", "ReportType:=",  _
  "RMxprt", "Solution:=", "Setup1 : Performance", Array("NAME:SimValueContext", "Domain:=",  _
  "Parameter"), "Calculation:=", "EfficiencyParameter", "Name:=",  _
  "EfficiencyParameter", Array("NAME:Ranges"))))
oProject.Save
oDesign.AnalyzeAll
oModule.ExportOptimetricsResult "ParametricSetup1",  _
  "C:/Users/by/Desktop/pmsm_son/output.txt", false
