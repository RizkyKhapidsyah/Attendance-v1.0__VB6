VERSION 5.00
Begin VB.Form frmChange 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   Icon            =   "frmChange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   3360
      MaxLength       =   20
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtPasswd 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3360
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdOk 
      Height          =   495
      Left            =   3120
      Picture         =   "frmChange.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ok"
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdCanel 
      Height          =   495
      Left            =   3840
      Picture         =   "frmChange.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancel"
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Login Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1800
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1800
      TabIndex        =   4
      Top             =   960
      Width           =   1260
   End
End
Attribute VB_Name = "frmChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim check As Boolean
Dim RS As New ADODB.Recordset

Private Sub cmdCanel_Click()
Unload Me
End Sub
Private Sub cmdOk_Click()
Validate
If check = True Then
SavePass
MsgBox "Password change successfully", vbInformation, "Attendance"
MsgBox "Your new Login is" & " " & txtUserName.Text & " " & "and new password is" & " " & txtPasswd.Text, vbInformation, "Attendance"
Unload Me
Else
MsgBox "Unable to change password", vbCritical, "Attendance"
End If
End Sub
Private Sub Form_Load()
ModuleConn.ConnPass
check = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
ModuleConn.DissConnPass
Unload Me
frmMain.Show
End Sub
Private Sub Validate()
If txtUserName.Text = "" And txtPasswd.Text = "" Then
MsgBox "Please enter valid username and password", vbCritical, "Attendance"
check = False
End If
End Sub
Private Sub SavePass()
If RS.State = adStateOpen Then
RS.Close
End If
RS.CursorLocation = adUseClient
RS.CursorType = adOpenStatic
RS.LockType = adLockOptimistic
RS.Open "select * from pass", CN1
If RS.RecordCount > 0 Then
RS.MoveFirst
RS("uu") = Trim(txtUserName.Text)
RS("pp") = Trim(txtPasswd.Text)
RS.Update
End If
End Sub
Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdOk_Click
End If
End Sub
