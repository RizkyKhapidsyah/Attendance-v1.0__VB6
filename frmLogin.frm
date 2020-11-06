VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login Window"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7020
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCanel 
      Height          =   495
      Left            =   3600
      Picture         =   "frmLogin.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdOk 
      Height          =   495
      Left            =   2880
      Picture         =   "frmLogin.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtPasswd 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   3360
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Line Line14 
      X1              =   240
      X2              =   960
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line13 
      X1              =   240
      X2              =   960
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line12 
      X1              =   240
      X2              =   960
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line11 
      X1              =   240
      X2              =   960
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line10 
      X1              =   240
      X2              =   960
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line9 
      X1              =   240
      X2              =   960
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line8 
      X1              =   240
      X2              =   960
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line7 
      X1              =   960
      X2              =   960
      Y1              =   480
      Y2              =   1200
   End
   Begin VB.Line Line6 
      X1              =   840
      X2              =   840
      Y1              =   480
      Y2              =   1200
   End
   Begin VB.Line Line5 
      X1              =   720
      X2              =   720
      Y1              =   480
      Y2              =   1200
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   240
      Y1              =   480
      Y2              =   1200
   End
   Begin VB.Line Line3 
      X1              =   360
      X2              =   360
      Y1              =   480
      Y2              =   1200
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   480
      Y1              =   480
      Y2              =   1200
   End
   Begin VB.Line Line1 
      X1              =   600
      X2              =   600
      Y1              =   480
      Y2              =   1200
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "frmLogin.frx":091E
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      TabIndex        =   3
      Top             =   960
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login Name"
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
      TabIndex        =   2
      Top             =   480
      Width           =   1020
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Private Sub cmdCanel_Click()
End
End Sub
Private Sub cmdOk_Click()
If rs.State = adStateOpen Then
rs.Close
End If

rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic

rs.Open "select * from pass where uu= '" & txtUserName.Text & "' and pp='" & txtPasswd.Text & "'", CN1
If rs.RecordCount > 0 Then
Unload Me
frmMain.Show
Else
MsgBox "Incorrect Login Attempt ", vbCritical, "Attendance"
End If
End Sub
Private Sub Form_Load()
ModuleConn.ConnPass
End Sub
Private Sub Form_Unload(Cancel As Integer)
ModuleConn.DissConnPass
Unload Me
End Sub
Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdOk_Click
End If
End Sub
