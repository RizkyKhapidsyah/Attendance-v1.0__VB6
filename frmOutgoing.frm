VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmOutgoing 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Outgoing"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   Icon            =   "frmOutgoing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   3480
      Picture         =   "frmOutgoing.frx":2C5A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Save"
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   375
      Left            =   4080
      Picture         =   "frmOutgoing.frx":2F64
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancel "
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   375
      Left            =   3720
      Picture         =   "frmOutgoing.frx":326E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Add New Attendance"
      Top             =   3960
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5520
      Top             =   2280
   End
   Begin VB.TextBox txtOutgoing 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   9
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtIncoming 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin VB.ComboBox cboName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox txtTDate 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7095
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   12515
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12648384
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "name"
         Caption         =   "Employee Name"
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
         MarqueeStyle    =   3
         Locked          =   -1  'True
         BeginProperty Column00 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Line Line3 
      X1              =   3855
      X2              =   3840
      Y1              =   360
      Y2              =   480
   End
   Begin VB.Line Line2 
      X1              =   3600
      X2              =   3840
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   3840
      Y1              =   1200
      Y2              =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Employee Names who had left today so far"
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   0
      Width           =   3075
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Outgoing Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2400
      TabIndex        =   10
      Top             =   1800
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
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
      Left            =   2760
      TabIndex        =   8
      Top             =   3000
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Today's Date"
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
      Left            =   3360
      TabIndex        =   7
      Top             =   960
      Width           =   1140
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IncomingTime"
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
      Left            =   4320
      TabIndex        =   6
      Top             =   1800
      Width           =   1185
   End
End
Attribute VB_Name = "frmOutgoing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsGrid As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim RS As New ADODB.Recordset
Private Sub cboName_Click()
If rs1.State = adStateOpen Then
rs1.Close
End If
rs1.CursorLocation = adUseClient
rs1.CursorType = adOpenStatic
rs1.LockType = adLockOptimistic
rs1.Open "select * from attend where name='" & cboName.Text & "' and tdate= #" & Date & "#", CN
If rs1.RecordCount > 0 Then
rs1.MoveFirst
txtIncoming.Text = rs1.Fields("incoming").Value
End If
End Sub
Private Sub cmdAdd_Click()
If Trim(cboName.Text) = "" Then
MsgBox "Please select a name from list", vbCritical, "Attendance"
Exit Sub
End If
cmdAdd.Visible = False
cmdSave.Visible = True
cmdCancel.Visible = True
End Sub

Private Sub cmdCancel_Click()
cmdAdd.Visible = True
cmdSave.Visible = False
cmdCancel.Visible = False
End Sub
Private Sub cmdSave_Click()
If rs1.State = adStateOpen Then
rs1.Close
End If
If cboName.Text = Trim("") Then
MsgBox "Please select a name from List ", vbCritical, "Attendance"
Exit Sub
End If
rs1.Open "select * from attend where name='" & cboName.Text & "' and tdate= #" & Date & "# and outgoing is null", CN
If rs1.RecordCount > 0 Then
rs1.MoveFirst
txtIncoming.Text = rs1.Fields("incoming").Value
rs1("outgoing") = Trim(txtOutgoing.Text)
rs1.Update
cboName.Text = ""
RefreshOutgoing
FillGrid
cmdAdd.Visible = True
cmdCancel.Visible = False
cmdSave.Visible = False
Else
MsgBox "Duplicate or wrong Entry !!! Could not add twice outgoing for" & "  " & cboName.Text, vbCritical, "Attendance"
cmdAdd.Visible = True
cmdCancel.Visible = False
cmdSave.Visible = False
Exit Sub
End If
End Sub
Private Sub Form_Load()
ModuleConn.AttendConn
RefreshOutgoing
FillGrid
End Sub
Private Sub Form_Unload(Cancel As Integer)
ModuleConn.CloseConn
Unload Me
frmMain.Show
End Sub
Private Sub RefreshOutgoing()
If RS.State = adStateOpen Then
RS.Close
End If
RS.CursorLocation = adUseClient
RS.CursorType = adOpenDynamic
RS.LockType = adLockOptimistic
cboName.Clear
RS.Open "select distinct name from attend", CN
If RS.RecordCount > 0 Then
        RS.MoveFirst
        While Not RS.EOF
        cboName.AddItem RS("name")
        RS.MoveNext
        Wend
RS.Close
End If
End Sub
Private Sub Timer1_Timer()
Timer1.Enabled = True
txtTDate.Text = Format(Now, "dd/mm/yyyy")
txtOutgoing.Text = Format(Now, "hh:mm:ss AM/PM")
End Sub
Private Sub FillGrid()
If RsGrid.State = adStateOpen Then
RsGrid.Close
End If
RsGrid.CursorLocation = adUseClient
RsGrid.CursorType = adOpenStatic
RsGrid.LockType = adLockOptimistic

RsGrid.Open "select name from attend where tdate = #" & Date & "# and (NOT(outgoing IS NULL))", CN
If RsGrid.RecordCount > 0 Then
RsGrid.MoveFirst
Set DataGrid1.DataSource = RsGrid
End If
End Sub
