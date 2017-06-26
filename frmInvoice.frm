VERSION 5.00
Begin VB.Form Invoice 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14745
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   14745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Update"
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
      Left            =   9240
      TabIndex        =   22
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Confirm"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   21
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000006&
      X1              =   14640
      X2              =   120
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000006&
      X1              =   7800
      X2              =   7800
      Y1              =   2160
      Y2              =   6000
   End
   Begin VB.Line Line5 
      X1              =   14640
      X2              =   14640
      Y1              =   1440
      Y2              =   6000
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   120
      Y1              =   1440
      Y2              =   6000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Patient And Hospital Details "
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   495
      Left            =   240
      TabIndex        =   23
      Top             =   1560
      Width           =   14295
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000006&
      X1              =   0
      X2              =   14760
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000006&
      X1              =   0
      X2              =   14760
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblout 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   495
      Index           =   8
      Left            =   10320
      TabIndex        =   20
      Top             =   5040
      Width           =   3015
   End
   Begin VB.Label lblout 
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
      ForeColor       =   &H80000003&
      Height          =   495
      Index           =   9
      Left            =   10320
      TabIndex        =   19
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label lblout 
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
      ForeColor       =   &H80000003&
      Height          =   495
      Index           =   7
      Left            =   10320
      TabIndex        =   18
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label lblout 
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
      ForeColor       =   &H80000003&
      Height          =   495
      Index           =   6
      Left            =   10560
      TabIndex        =   17
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label lblout 
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
      ForeColor       =   &H80000003&
      Height          =   495
      Index           =   5
      Left            =   3720
      TabIndex        =   16
      Top             =   4320
      Width           =   3735
   End
   Begin VB.Label lblout 
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
      ForeColor       =   &H80000003&
      Height          =   495
      Index           =   4
      Left            =   3600
      TabIndex        =   15
      Top             =   5040
      Width           =   3975
   End
   Begin VB.Label lblout 
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
      ForeColor       =   &H80000003&
      Height          =   495
      Index           =   3
      Left            =   3720
      TabIndex        =   14
      Top             =   3720
      Width           =   3735
   End
   Begin VB.Label lblout 
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
      ForeColor       =   &H80000003&
      Height          =   495
      Index           =   2
      Left            =   10200
      TabIndex        =   13
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label lblout 
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
      ForeColor       =   &H80000003&
      Height          =   495
      Index           =   1
      Left            =   3720
      TabIndex        =   12
      Top             =   3120
      Width           =   3735
   End
   Begin VB.Label lblout 
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
      ForeColor       =   &H80000003&
      Height          =   495
      Index           =   0
      Left            =   3720
      TabIndex        =   11
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Label lblstatic 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Bill Amount"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   375
      Index           =   9
      Left            =   8280
      TabIndex        =   10
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label lblstatic 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H8000000A&
      Height          =   375
      Index           =   8
      Left            =   8400
      TabIndex        =   9
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label lblstatic 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H8000000A&
      Height          =   375
      Index           =   7
      Left            =   8400
      TabIndex        =   8
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblstatic 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H8000000A&
      Height          =   375
      Index           =   6
      Left            =   8400
      TabIndex        =   7
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lblstatic 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Alternate Contact"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Index           =   5
      Left            =   480
      TabIndex        =   6
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label lblstatic 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H8000000A&
      Height          =   375
      Index           =   4
      Left            =   600
      TabIndex        =   5
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lblstatic 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H8000000A&
      Height          =   375
      Index           =   3
      Left            =   600
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblstatic 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H8000000A&
      Height          =   375
      Index           =   2
      Left            =   8400
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblstatic 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H8000000A&
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblstatic 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Patient ID"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderStyle     =   4  'Dash-Dot
      X1              =   6120
      X2              =   8040
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblInvoice 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Invoice"
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
      Height          =   495
      Left            =   6240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Invoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim query As String
Public activeUser As String

Private Sub cmdCancel_Click()
Me.Hide
Update.Show
End Sub

Private Sub cmdOK_Click()
Me.Hide
Home.Show
End Sub

Private Sub Form_Activate()
cn.Open ("Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + App.Path + "\opdDbms.mdb;")
query = "select * from patientDetails where patientID = '" + activeUser + " ' "
Set rs = cn.Execute(query)
Dim i As Integer
For i = 0 To 9
lblout(i).Caption = rs(i)
Next i
Set rs = Nothing
cn.Close
End Sub
