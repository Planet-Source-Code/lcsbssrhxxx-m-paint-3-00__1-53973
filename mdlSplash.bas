Attribute VB_Name = "mdlSplash"
'This Module Is The Sub Main For The Program (Where The Program Starts From)
'So It Will Do The Fallowing Events On Start Up Of M-Paint,
'This Is The Code For The Splash Screen To Show On The Start Up Then
'Hides It.

Sub Main()
'Show Splash Screen
frmSplash.Show
'Refresh Splash Screen
frmSplash.Refresh
'Unload Splash Screen
Unload frmSplash
'Show Main Form
frmMain.Show
End Sub
