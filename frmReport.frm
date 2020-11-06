VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmReport 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attendance Report"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7995
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHtml 
      Height          =   615
      Left            =   6120
      Picture         =   "frmReport.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Save as"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdToday 
      Height          =   615
      Left            =   5280
      Picture         =   "frmReport.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Today's Report"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   4680
      Picture         =   "frmReport.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "View all Attendance"
      Top             =   120
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3975
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7011
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BackColor       =   12648384
      ColumnHeaders   =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         Locked          =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSearch 
      Height          =   615
      Left            =   3960
      Picture         =   "frmReport.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Search Now"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdView 
      Height          =   615
      Left            =   6720
      Picture         =   "frmReport.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "View Report"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   615
      Left            =   7320
      Picture         =   "frmReport.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Print Report to Printer"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtfrom 
      Height          =   285
      Left            =   840
      MaxLength       =   10
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Top             =   0
      Width           =   435
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   120
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2160
      TabIndex        =   7
      Top             =   360
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   405
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset

Private Sub cmdHtml_Click()
With rptdynamic
            Set .DataSource = Nothing
                .DataMember = ""
            Set .DataSource = RS.DataSource
                With .Sections("Section1").Controls
                    For i = 1 To .Count
                        If TypeOf .Item(i) Is RptTextBox Then
                             .Item(i).DataMember = ""
                            .Item(i).DataField = RS.Fields(i - 1).Name
                        End If
                    Next i
                End With
               rptdynamic.ExportReport
            End With

End Sub

Private Sub cmdPrint_Click()
With rptdynamic
            Set .DataSource = Nothing
                .DataMember = ""
            Set .DataSource = RS.DataSource
                With .Sections("Section1").Controls
                    For i = 1 To .Count
                        If TypeOf .Item(i) Is RptTextBox Then
                             .Item(i).DataMember = ""
                            .Item(i).DataField = RS.Fields(i - 1).Name
                        End If
                    Next i
                End With
               rptdynamic.PrintReport (True)
            End With
MsgBox "Report sent to printer ", vbInformation, "Attendance"
End Sub

Private Sub cmdSearch_Click()
If RS.State = adStateOpen Then
RS.Close
End If
If txtfrom.Text = "" Or txtTo.Text = "" Then
MsgBox "Please Enter valid dates ", vbCritical, "Attendance"
Exit Sub
Else
If Not IsDate(txtfrom.Text) Then
MsgBox "Please enter a valid date in it", vbCritical, "Attendance"
Exit Sub
Else
If Not IsDate(txtTo.Text) Then
MsgBox "Please enter a valid date in it", vbCritical, "Attendance"
Exit Sub
Else
RS.Open "select * from attend where tdate between #" & CDate(txtfrom.Text) & "# and #" & CDate(txtTo.Text) & "#", CN
If RS.RecordCount > 0 Then
RS.MoveFirst
Set DataGrid1.DataSource = RS
Else
MsgBox "Could not found Attendance between" & "  " & txtfrom.Text & " and " & txtTo.Text, vbCritical, "Attendance"
End If
End If
End If
End If
End Sub
Private Sub cmdSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdSearch_Click
End If
End Sub
Private Sub cmdToday_Click()
If RS.State = adStateOpen Then
RS.Close
End If
RS.Open "select * from attend where tdate= #" & Date & "#", CN
If RS.RecordCount > 0 Then
RS.MoveFirst
Set DataGrid1.DataSource = RS
End If
End Sub

Private Sub cmdView_Click()
 With rptdynamic
            Set .DataSource = Nothing
                .DataMember = ""
            Set .DataSource = RS.DataSource
                With .Sections("Section1").Controls
                    For i = 1 To .Count
                        If TypeOf .Item(i) Is RptTextBox Then
                            'The datamember should be always blank while creating dynamic data reports
                            .Item(i).DataMember = ""
                            .Item(i).DataField = RS.Fields(i - 1).Name
                        End If
                    Next i
                End With
               .Show
       End With
End Sub
Private Sub Command1_Click()
FillGrid
End Sub
Private Sub Form_Load()
ModuleConn.AttendConn
FillGrid
End Sub
Private Sub Form_Unload(Cancel As Integer)
ModuleConn.CloseConn
Unload Me
frmMain.Show
End Sub
Private Sub FillGrid()
If RS.State = adStateOpen Then
RS.Close
End If

RS.CursorLocation = adUseClient
RS.CursorType = adOpenStatic
RS.LockType = adLockOptimistic

RS.Open "select * from attend", CN
If RS.RecordCount > 0 Then
RS.MoveFirst
Set DataGrid1.DataSource = RS
End If
End Sub
Private Sub txtTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdSearch_Click
End If
End Sub
