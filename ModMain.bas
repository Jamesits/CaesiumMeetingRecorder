Attribute VB_Name = "ModMain"
Option Explicit
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Versionstring
Public m_transparencyKey As Long
Public Enable7Feature As Boolean
Public FrmClock As Form
Public FrmDock As Form
Public FrmSplash As Form
Public FrmWizard As Form

Public Sub Main()
InitCommonControls
Dim cOS As New clsOS
Enable7Feature = IIf(CInt(Left(cOS.OS_Version, 1)) < 7, False, True)
Set cOS = Nothing
Versionstring = "Version " & App.Major & "." & App.Minor & "." & "0" & "." & App.Revision & " " & "Pre-Alpha"
m_transparencyKey = RGB(0, 0, 0)
SelectForm
End Sub
Private Sub SelectForm()
If Enable7Feature Then
Load FrmSplash
FrmSplash.Show
Load FrmWizard
FrmSplash.Hide
FrmWizard.Show
Unload FrmSplash
Else
Load FrmSplashXP
FrmSplashXP.Show
Load FrmWizardXP
FrmSplashXP.Hide
FrmWizardXP.Show
Unload FrmSplashXP
End If
End Sub
Public Sub quit()
Dim f As Form
For Each f In Forms
Unload f
Next
End
End Sub
