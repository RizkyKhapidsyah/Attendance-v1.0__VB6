Attribute VB_Name = "Module1"
Global it As Integer
Dim Admin As New frmAdmin
Dim Main As New frmMain
Dim Log As New frmLogin
Public Sub LoadLoginForm()
Load Log
Log.Left = Screen.Width / 2 - Log.Width / 2
Log.Top = Screen.Height / 2 - Log.Height / 2
End Sub
Public Sub UnLoadLoginForm()
Unload Log
End Sub
Public Sub MainFormsLoad()
Load Main
Main.Left = 1000
Main.Top = 1000
End Sub
Public Sub MainFormsUnload()
Unload Main
End Sub
Public Sub AdminFormLoad()
Load Admin
Admin.Left = 5000
Admin.Top = 3000
End Sub
Public Sub AdminFormUnload()
Unload Admin
End Sub
