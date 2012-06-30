Attribute VB_Name = "ModMain"
Option Explicit
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Versionstring


Public Sub Main()
InitCommonControls
Versionstring = "Version " & App.Major & "." & App.Minor & "." & "0" & "." & App.Revision & " " & "Pre-Alpha"
Load FrmSplash
FrmSplash.Show
Load FrmWizard
FrmSplash.Hide
FrmWizard.Show
Unload FrmSplash
End Sub

Public Sub quit()
Dim f As Form
For Each f In Forms
Unload f
Next
End
End Sub
