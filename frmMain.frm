VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu                                              "
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Change Password"
      Height          =   975
      Left            =   3240
      Picture         =   "frmMain.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Quit"
      Height          =   975
      Left            =   4320
      Picture         =   "frmMain.frx":1794
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Report"
      Height          =   975
      Left            =   2160
      Picture         =   "frmMain.frx":1A9E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OutGoing"
      Height          =   975
      Left            =   1080
      Picture         =   "frmMain.frx":1DA8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton frmMain 
      Caption         =   "InComing"
      Height          =   975
      Left            =   0
      Picture         =   "frmMain.frx":20B2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
frmOutgoing.Show
End Sub
Private Sub Command2_Click()
Unload Me
frmReport.Show
End Sub
Private Sub Command3_Click()
End
End Sub
Private Sub Command4_Click()
Unload Me
frmChange.Show
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
Private Sub frmMain_Click()
Unload Me
frmIncoming.Show
End Sub
