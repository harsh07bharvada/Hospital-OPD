VERSION 5.00
Begin VB.Form Update 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11850
   FillColor       =   &H00404080&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton back 
      BackColor       =   &H8000000E&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox textout 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Index           =   2
      Left            =   8160
      TabIndex        =   25
      Top             =   2400
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox textout 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Index           =   9
      Left            =   8040
      TabIndex        =   24
      Top             =   6240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox textout 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Index           =   8
      Left            =   8040
      TabIndex        =   23
      Top             =   5040
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox textout 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Index           =   7
      Left            =   720
      TabIndex        =   21
      Top             =   5040
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox textout 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Index           =   6
      Left            =   840
      TabIndex        =   22
      Top             =   6360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton btnUpdate 
      BackColor       =   &H80000016&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox textout 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Index           =   5
      Left            =   4200
      TabIndex        =   14
      Top             =   6240
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox textout 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Index           =   4
      Left            =   720
      TabIndex        =   12
      Top             =   3600
      Visible         =   0   'False
      Width           =   10215
   End
   Begin VB.TextBox textout 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Index           =   3
      Left            =   4080
      TabIndex        =   10
      Top             =   4920
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox textout 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Index           =   1
      Left            =   4080
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox textout 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton btnEnter 
      BackColor       =   &H80000016&
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      MaskColor       =   &H80000010&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox textID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Text            =   "PatientID"
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Height          =   975
      Left            =   -1005
      TabIndex        =   20
      Top             =   960
      Width           =   975
   End
   Begin VB.Label patientlbl 
      Alignment       =   2  'Center
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   375
      Index           =   9
      Left            =   8400
      TabIndex        =   18
      Top             =   5760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label patientlbl 
      Alignment       =   2  'Center
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   375
      Index           =   8
      Left            =   8280
      TabIndex        =   17
      Top             =   4440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label patientlbl 
      Alignment       =   2  'Center
      Caption         =   "Doctor"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   375
      Index           =   7
      Left            =   1080
      TabIndex        =   16
      Top             =   4440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label patientlbl 
      Alignment       =   2  'Center
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   375
      Index           =   6
      Left            =   1080
      TabIndex        =   15
      Top             =   5760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label patientlbl 
      Alignment       =   2  'Center
      Caption         =   "Alternative Contact"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   495
      Index           =   5
      Left            =   3960
      TabIndex        =   13
      Top             =   5760
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label patientlbl 
      Alignment       =   2  'Center
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   375
      Index           =   4
      Left            =   4320
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label patientlbl 
      Alignment       =   2  'Center
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   375
      Index           =   3
      Left            =   4800
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label patientlbl 
      Alignment       =   2  'Center
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   375
      Index           =   2
      Left            =   8160
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label patientlbl 
      Alignment       =   2  'Center
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label patientlbl 
      Alignment       =   2  'Center
      Caption         =   "PatientID"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblHeading 
      Caption         =   "Please enter your PatientId below:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   2400
      Width           =   5295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   3360
      X2              =   8520
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblUpdate 
      Alignment       =   2  'Center
      Caption         =   "UPDATE PATIENT DETAILS"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim query As String
Dim flag As Integer
Private Sub back_Click()
Me.Hide
Home.Show
End Sub
Private Sub btn_Click()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim query As String
cn.Open ("Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + App.Path + "\opddbms.mdb;")
query = "select * from opd"
Set rs = cn.Execute(query)
MsgBox rs(0)

Set rs = Nothing
cn.Close
End Sub
Private Sub btnEnter_Click()

If textID.Text = "" Then
MsgBox "Enter PatientID for searching.."
Exit Sub
End If

query = "select * from patientDetails where patientID = '" + textID.Text + " ' "
Set rs = cn.Execute(query)
If rs.EOF = True Or rs.BOF = True Then
MsgBox "No Records Found"
flag = 0
ElseIf rs(0) = textID.Text Then
flag = 1
lblheading.Visible = False
textID.Visible = False
btnEnter.Visible = False

detailsshow
Else
MsgBox "Cant found"
flag = 0
End If
End Sub

Private Sub btnUpdate_Click()
cn.Execute "Update patientDetails set patientID = '" + textout(0).Text + "', patientName = '" + textout(1).Text + "', patientDept = '" + textout(2).Text + "', patientContact = '" + textout(3).Text + "', patientAddress = '" + textout(4).Text + "', patientAltContact = '" + textout(5).Text + "', patientTime = #" + textout(6).Text + "#, patientDoc = '" + textout(7).Text + "', patientAmount = '" + textout(8).Text + "', patientDate = #" + textout(9).Text + "# where patientID='" + textout(0).Text + "'"
MsgBox "done"
cn.Close
Invoice.activeUser = Val(textout(0).Text)
Me.Hide
Invoice.Show
End Sub
Private Sub Form_Activate()
If Not IsNumeric(textID.Text) Then
btnEnter.SetFocus
End If
cn.Open ("Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + App.Path + "\opdDbms.mdb;")
End Sub
Private Sub textID_GotFocus()
textID.Text = ""
End Sub
Private Sub detailsshow()
Dim X As Integer
For X = 0 To 9
patientlbl(X).Visible = True
textout(X).Visible = True
textout(X).Text = rs(X)
Next
btnUpdate.Visible = True
End Sub

Private Sub textout_Change(Index As Integer)
If Index = 8 Then
Dim flag As Boolean
flag = IsNumeric(textout(Index).Text)
If flag = False Then
Beep
MsgBox "Enter a number!", vbCritical
textout(8).Text = "0"
End If
End If
If Index = 0 Then
textout(Index).Enabled = False
End If
If Index = 9 Then
If Not IsDate(textout(Index).Text) Then
Beep
MsgBox "Enter a date!", vbCritical
End If
End If
If Index = 6 Then
If Not IsDate(textout(Index).Text) Then
Beep
MsgBox "Enter  Time!", vbCritical
End If
End If
End Sub

