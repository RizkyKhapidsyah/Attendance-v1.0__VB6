VERSION 5.00
Begin VB.Form frmAdmin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administrator"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   Icon            =   "frmAdmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Change Password"
      Height          =   975
      Left            =   3840
      Picture         =   "frmAdmin.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete User"
      Height          =   975
      Left            =   2880
      Picture         =   "frmAdmin.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "Create New User"
      Height          =   975
      Left            =   1920
      Picture         =   "frmAdmin.frx":149E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Month Wise Report"
      Height          =   975
      Left            =   960
      Picture         =   "frmAdmin.frx":17A8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Modify Attendance"
      Height          =   975
      Left            =   0
      Picture         =   "frmAdmin.frx":1AB2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
