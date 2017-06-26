VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form3"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12585
   LinkTopic       =   "Form3"
   ScaleHeight     =   8295
   ScaleWidth      =   12585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton back 
      Appearance      =   0  'Flat
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Label lblAppointment 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   3615
      Left            =   9720
      TabIndex        =   14
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label lblDepart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   3615
      Left            =   7320
      TabIndex        =   13
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label lblArrive 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   3615
      Left            =   5280
      TabIndex        =   12
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Depart time"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   495
      Left            =   7320
      TabIndex        =   11
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblDept 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   3615
      Left            =   3000
      TabIndex        =   10
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Arrive Time"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   495
      Left            =   5280
      TabIndex        =   9
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Line Line10 
      X1              =   7200
      X2              =   7200
      Y1              =   2160
      Y2              =   7560
   End
   Begin VB.Label lblDocName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   3615
      Left            =   720
      TabIndex        =   8
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label dayOut 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   495
      Left            =   10440
      TabIndex        =   7
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label dateOut 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Day:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   495
      Left            =   9600
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000006&
      X1              =   5160
      X2              =   5160
      Y1              =   2160
      Y2              =   7560
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000006&
      X1              =   9480
      X2              =   9480
      Y1              =   2160
      Y2              =   7560
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000006&
      X1              =   2880
      X2              =   2880
      Y1              =   2160
      Y2              =   7560
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Appointments Left"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   735
      Left            =   9720
      TabIndex        =   3
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Doctor Name"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000A&
      X1              =   480
      X2              =   12000
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000A&
      X1              =   480
      X2              =   12000
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000A&
      X1              =   12000
      X2              =   12000
      Y1              =   2160
      Y2              =   7560
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000A&
      X1              =   480
      X2              =   12000
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000A&
      X1              =   480
      X2              =   480
      Y1              =   2160
      Y2              =   7560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000006&
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   4320
      X2              =   8280
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblheading 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Today's Schedule"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   615
      Left            =   4440
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim con1 As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim query As String
Dim query1 As String
Dim Avail As Long
Dim int1 As Integer
Dim temp As String
Dim length As Integer
Dim daynum As Integer
Dim count1 As Integer




Private Sub back_Click()
Me.Hide
Home.Show
End Sub

Private Sub Form_Activate()
dateOut.Caption = DateValue(Now)
daynum = Weekday(DateValue(Now), vbMonday)

Select Case daynum
Case 7
dayOut.Caption = "Sunday"
Case 1
dayOut.Caption = "Monday"
Case 2
dayOut.Caption = "Tuesday"
Case 3
dayOut.Caption = "Wednesday"
Case 4
dayOut.Caption = "Thursday"
Case 5
dayOut.Caption = "Friday"
Case 6
dayOut.Caption = "Saturday"
End Select

con.Open ("Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + App.Path + "\opdDbms.mdb;")

query = "select * from Doctor"
Set rs = con.Execute(query)
con1.Open ("Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + App.Path + "\opdDbms.mdb;")
query1 = "select * from patientDetails where patientDate= #" + dateOut.Caption + "#"
Set rs1 = con1.Execute(query1)
While rs.EOF = False And rs.BOF = False

Avail = rs(2)
temp = Val(rs(2))
length = Len(temp)


While length <> 0
int1 = Avail Mod 10
Avail = Avail - int1

Avail = Avail / 10

If int1 = daynum Then
lblDocName = lblDocName + vbCrLf + rs(1)
lblDept = lblDept + vbCrLf + rs(8)
lblArrive = lblArrive + vbCrLf & rs(3)
lblDepart = lblDepart + vbCrLf & rs(4)


While rs1.EOF = False And rs1.BOF = False
If rs(1) = rs1(7) Then
count1 = count1 + 1
End If
rs1.MoveNext
Wend
count1 = 10 - count1
lblAppointment = lblAppointment + vbCrLf + Str(count1)
End If
count1 = 0
length = length - 1

Wend

rs.MoveNext

Wend
Set rs = Nothing
con.Close
Set rs1 = Nothing
con1.Close

End Sub


