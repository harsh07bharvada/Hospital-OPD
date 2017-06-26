VERSION 5.00
Begin VB.Form Home 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Home"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12885
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   12885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAvailablity 
      BackColor       =   &H8000000B&
      Caption         =   "Today's Schedule"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   3855
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H8000000B&
      Caption         =   "Update Entries"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   3255
   End
   Begin VB.CommandButton cmdInsert 
      BackColor       =   &H8000000B&
      Caption         =   "New Registration"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   3255
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H8000000B&
      Caption         =   "Search Entries"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   3855
   End
   Begin VB.Line Line5 
      X1              =   1800
      X2              =   11040
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line4 
      X1              =   11040
      X2              =   11040
      Y1              =   3240
      Y2              =   5640
   End
   Begin VB.Line Line3 
      X1              =   1800
      X2              =   11040
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line2 
      X1              =   1800
      X2              =   1800
      Y1              =   3240
      Y2              =   5640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000006&
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   3840
      X2              =   9840
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label lblRecDoc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total Doctors:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Top             =   7320
      Width           =   3375
   End
   Begin VB.Label lblRecPat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total Patients:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   7320
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "St Thomas' Hospital Desk"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   735
      Left            =   3720
      TabIndex        =   4
      Top             =   1080
      Width           =   6255
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   1095
      Left            =   8760
      TabIndex        =   3
      Top             =   6840
      Width           =   2535
   End
End
Attribute VB_Name = "Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sql As String
Dim query As String
Dim dtmTest As Date
Dim timetest As Date
Public patientNum As Integer

Private Sub cmdAvailablity_Click()
Me.Hide
Form3.Show
End Sub

Private Sub cmdInsert_Click()
Me.Hide
Insert.Show
End Sub

Private Sub cmdSearch_Click()
Me.Hide
Search.Show
End Sub

Private Sub cmdUpdate_Click()
Me.Hide
Update.Show
End Sub

Private Sub Form_Activate()
dtmTest = DateValue(Now)
timetest = TimeValue(Now)
Label7.Caption = Str(timetest) + Str(dtmTest)
lblRecPat.Caption = "Total Patients : "
lblRecDoc.Caption = "Total Doctors : "
cn.Open ("Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + App.Path + "\opdDbms.mdb;")
sql = "select patientID from patientDetails"
query = "select DocID from Doctor"

Set rs = cn.Execute(sql)
While rs.EOF = False And rs.BOF = False
patientNum = rs(0)
rs.MoveNext
Wend
lblRecPat.Caption = lblRecPat.Caption + "  " + Str(patientNum)
Set rs = Nothing

Set rs = cn.Execute(sql)
lblRecDoc.Caption = lblRecDoc.Caption + "  " + Str(rs(0))
Set rs = Nothing
cn.Close
End Sub

Private Sub Timer1_Timer()
dtmTest = DateValue(Now)
timetest = TimeValue(Now)
Label7.Caption = Str(timetest) + Str(dtmTest)
End Sub
