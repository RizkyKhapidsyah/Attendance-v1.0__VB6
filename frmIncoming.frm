VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmIncoming 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Incoming"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   Icon            =   "frmIncoming.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7095
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   12515
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12648384
      ColumnHeaders   =   -1  'True
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
   Begin VB.CommandButton cmdAdd 
      Height          =   375
      Left            =   3600
      Picture         =   "frmIncoming.frx":2C5A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Add New Attendance"
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   375
      Left            =   3960
      Picture         =   "frmIncoming.frx":33C4
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancel "
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   3360
      Picture         =   "frmIncoming.frx":36CE
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Save"
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
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
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1440
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
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5640
      Top             =   3720
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
      Height          =   270
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Line Line3 
      X1              =   3840
      X2              =   3720
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
      X1              =   2160
      X2              =   3840
      Y1              =   960
      Y2              =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Employees Name who has come  office Today"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   0
      Width           =   3345
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
      TabIndex        =   7
      Top             =   1200
      Width           =   1185
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
      Left            =   2280
      TabIndex        =   6
      Top             =   1200
      Width           =   1140
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
      TabIndex        =   5
      Top             =   2640
      Width           =   1365
   End
End
Attribute VB_Name = "frmIncoming"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cmd As New Command
Dim RsGrid As New Recordset
Dim rs1 As New ADODB.Recordset
Dim RS As New ADODB.Recordset
Private Sub cmdAdd_Click()
If (Trim(cboName.Text = "")) Then
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
If cboName.Text = Trim("") Then
MsgBox "Name could not be blank ", vbCritical, "Attendance"
Exit Sub
End If
FindIncoming
ddd = txtTDate.Text
rs1.Open "select * from attend where name='" & cboName.Text & "' and tdate= #" & CDate(txtTDate.Text) & "#", CN
If Not rs1.EOF Then
MsgBox "Could not add the same attendance twice for" & "  " & cboName.Text, vbCritical, "Attendance"
cmdAdd.Visible = True
cmdSave.Visible = False
cmdCancel.Visible = False
rs1.Close
Exit Sub
End If
rs1.Close
If RS.State = adStateOpen Then
RS.Close
End If
RS.Open "select * from attend", CN
RS.AddNew
RS("name") = cboName.Text
RS("incoming") = CDate(txtIncoming.Text)
RS("tdate") = CDate(txtTDate.Text)
RS.Update
FillGrid
RS.Close
cmdAdd.Visible = True
cmdSave.Visible = False
cmdCancel.Visible = False
End Sub
Private Sub Form_Load()
ModuleConn.AttendConn
RefreshIncoming
FillGrid
End Sub
Private Sub Form_Unload(Cancel As Integer)
ModuleConn.CloseConn
Unload Me
frmMain.Show
End Sub
Private Sub Timer1_Timer()
Timer1.Enabled = True
txtIncoming.Text = Format(Now, "hh:mm:ss AM/PM")
txtTDate.Text = Format(Now, "dd/mm/yyyy")
End Sub
Private Sub RefreshIncoming()
If RS.State = adStateOpen Then
RS.Close
End If
RS.CursorLocation = adUseClient
RS.CursorType = adOpenStatic
RS.LockType = adLockOptimistic
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
Private Sub FindIncoming()
If rs1.State = adStateOpen Then
rs1.Close
End If
rs1.CursorLocation = adUseClient
rs1.CursorType = adOpenDynamic
rs1.LockType = adLockOptimistic
End Sub
Private Sub FillGrid()
If RsGrid.State = adStateOpen Then
RsGrid.Close
End If
RsGrid.CursorLocation = adUseClient
RsGrid.CursorType = adOpenStatic
RsGrid.LockType = adLockOptimistic

RsGrid.Open "select name from attend where tdate = #" & Date & "#", CN
If RsGrid.RecordCount > 0 Then
RsGrid.MoveFirst
Set DataGrid1.DataSource = RsGrid
End If
End Sub
