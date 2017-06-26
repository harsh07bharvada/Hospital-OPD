VERSION 5.00
Begin VB.Form Search 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Search"
   ClientHeight    =   9360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16155
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   16155
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton back 
      BackColor       =   &H8000000B&
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   1320
      Width           =   975
   End
   Begin VB.OptionButton optPat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Patient's Details"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   1440
      Width           =   3495
   End
   Begin VB.OptionButton optDoc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Doctor's Details"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   420
      Left            =   8640
      TabIndex        =   0
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Frame framePat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Patient's Details"
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   600
      TabIndex        =   18
      Top             =   2160
      Width           =   15015
      Begin VB.ComboBox CmbSearch 
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
         ForeColor       =   &H80000002&
         Height          =   465
         ItemData        =   "frmSearch.frx":0000
         Left            =   2040
         List            =   "frmSearch.frx":000A
         TabIndex        =   39
         Text            =   "Search by"
         Top             =   600
         Width           =   3375
      End
      Begin VB.CommandButton cmdSearch1 
         Appearance      =   0  'Flat
         Caption         =   "Search"
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
         Left            =   9720
         TabIndex        =   20
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox txtPatid 
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
         ForeColor       =   &H80000006&
         Height          =   495
         Left            =   5640
         TabIndex        =   19
         Top             =   600
         Width           =   3615
      End
      Begin VB.Line Line14 
         BorderColor     =   &H8000000A&
         X1              =   2280
         X2              =   13560
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Line Line13 
         BorderColor     =   &H8000000A&
         X1              =   10080
         X2              =   10080
         Y1              =   1440
         Y2              =   4800
      End
      Begin VB.Line Line12 
         BorderColor     =   &H8000000A&
         X1              =   2280
         X2              =   2280
         Y1              =   1440
         Y2              =   5760
      End
      Begin VB.Line Line11 
         BorderColor     =   &H8000000A&
         X1              =   7200
         X2              =   7200
         Y1              =   1440
         Y2              =   4800
      End
      Begin VB.Line Line10 
         BorderColor     =   &H8000000A&
         X1              =   13560
         X2              =   13560
         Y1              =   1440
         Y2              =   5760
      End
      Begin VB.Line Line9 
         BorderColor     =   &H8000000A&
         X1              =   120
         X2              =   13560
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Line Line8 
         BorderColor     =   &H8000000A&
         X1              =   120
         X2              =   120
         Y1              =   1440
         Y2              =   5760
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000010&
         X1              =   120
         X2              =   13560
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label LblIdP1 
         BackColor       =   &H00FFFFFF&
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
         Left            =   10200
         TabIndex        =   42
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label LblIdP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   375
         Left            =   7800
         TabIndex        =   41
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lblDateP1 
         BackColor       =   &H00FFFFFF&
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
         Left            =   10200
         TabIndex        =   38
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Label lblDate 
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
         ForeColor       =   &H80000006&
         Height          =   495
         Left            =   7800
         TabIndex        =   37
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label lblAltContact1 
         BackColor       =   &H00FFFFFF&
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
         Height          =   615
         Left            =   10200
         TabIndex        =   36
         Top             =   4080
         Width           =   2775
      End
      Begin VB.Label lblAltContact 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Contact 2"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   495
         Left            =   7680
         TabIndex        =   35
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label lblNameP 
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
         ForeColor       =   &H80000006&
         Height          =   375
         Left            =   600
         TabIndex        =   34
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblDepartment 
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
         ForeColor       =   &H80000006&
         Height          =   375
         Left            =   480
         TabIndex        =   33
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label lblDocName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Doctor Name"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   615
         Left            =   7680
         TabIndex        =   32
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label lblEntryTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Entry Time"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   375
         Left            =   480
         TabIndex        =   31
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label lblPhNumberP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Contact 1"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   375
         Left            =   480
         TabIndex        =   30
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label lblAddressP 
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
         ForeColor       =   &H80000006&
         Height          =   375
         Left            =   480
         TabIndex        =   29
         Top             =   4920
         Width           =   1455
      End
      Begin VB.Label lblBill 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Billing Amount"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   615
         Left            =   3600
         TabIndex        =   28
         Top             =   6000
         Width           =   2175
      End
      Begin VB.Label lblNameP1 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3120
         TabIndex        =   27
         Top             =   2040
         Width           =   3975
      End
      Begin VB.Label lblDepartment1 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3120
         TabIndex        =   26
         Top             =   2760
         Width           =   3975
      End
      Begin VB.Label lblEnterTime1 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3120
         TabIndex        =   25
         Top             =   3480
         Width           =   3975
      End
      Begin VB.Label lblDocName1 
         BackColor       =   &H00FFFFFF&
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
         Left            =   10200
         TabIndex        =   24
         Top             =   3360
         Width           =   2775
      End
      Begin VB.Label lblPhNumberP1 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3120
         TabIndex        =   23
         Top             =   4200
         Width           =   3975
      End
      Begin VB.Label lblAddressP1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   3120
         TabIndex        =   22
         Top             =   4920
         Width           =   9855
      End
      Begin VB.Label lblAmountP1 
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
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5880
         TabIndex        =   21
         Top             =   6000
         Width           =   4695
      End
   End
   Begin VB.Frame frameDoc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Doctor's Details"
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   600
      TabIndex        =   2
      Top             =   2400
      Width           =   14895
      Begin VB.CheckBox chkDay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sun"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   6
         Left            =   7680
         TabIndex        =   51
         Top             =   2760
         Width           =   735
      End
      Begin VB.CheckBox chkDay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sat"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   5
         Left            =   6840
         TabIndex        =   50
         Top             =   2760
         Width           =   855
      End
      Begin VB.CheckBox chkDay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fri"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   4
         Left            =   6120
         TabIndex        =   49
         Top             =   2760
         Width           =   735
      End
      Begin VB.CheckBox chkDay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Thurs"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   3
         Left            =   5160
         TabIndex        =   48
         Top             =   2760
         Width           =   975
      End
      Begin VB.CheckBox chkDay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Wed"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   47
         Top             =   2760
         Width           =   855
      End
      Begin VB.CheckBox chkDay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tue"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   1
         Left            =   3360
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   46
         Top             =   2760
         Width           =   735
      End
      Begin VB.CheckBox chkDay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Mon"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   45
         Top             =   2760
         Width           =   975
      End
      Begin VB.ComboBox CmbSearch1 
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
         ForeColor       =   &H80000003&
         Height          =   465
         ItemData        =   "frmSearch.frx":002E
         Left            =   2040
         List            =   "frmSearch.frx":0038
         TabIndex        =   40
         Text            =   "Search by"
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtDocid 
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
         ForeColor       =   &H80000006&
         Height          =   495
         Left            =   6720
         TabIndex        =   4
         Top             =   840
         Width           =   3615
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10920
         TabIndex        =   3
         Top             =   720
         Width           =   2415
      End
      Begin VB.Line Line16 
         BorderColor     =   &H8000000A&
         X1              =   11400
         X2              =   11400
         Y1              =   1680
         Y2              =   5640
      End
      Begin VB.Line Line15 
         BorderColor     =   &H8000000A&
         X1              =   2280
         X2              =   2280
         Y1              =   1680
         Y2              =   5640
      End
      Begin VB.Line Line7 
         BorderColor     =   &H8000000A&
         X1              =   9360
         X2              =   9360
         Y1              =   1680
         Y2              =   5640
      End
      Begin VB.Line Line6 
         BorderColor     =   &H8000000A&
         X1              =   14760
         X2              =   14760
         Y1              =   1680
         Y2              =   5640
      End
      Begin VB.Line Line4 
         BorderColor     =   &H8000000A&
         X1              =   120
         X2              =   14760
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         X1              =   120
         X2              =   14760
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   120
         X2              =   120
         Y1              =   1680
         Y2              =   5640
      End
      Begin VB.Label lblId1 
         BackColor       =   &H00FFFFFF&
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
         Height          =   495
         Left            =   11760
         TabIndex        =   44
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label lblId 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   375
         Left            =   9960
         TabIndex        =   43
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lblDegrees1 
         BackColor       =   &H00FFFFFF&
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
         Height          =   615
         Left            =   2520
         TabIndex        =   17
         Top             =   4800
         Width           =   6135
      End
      Begin VB.Label lblAddress1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   615
         Left            =   2520
         TabIndex        =   16
         Top             =   3840
         Width           =   6135
      End
      Begin VB.Label lblPhNumber1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   615
         Left            =   11760
         TabIndex        =   15
         Top             =   4680
         Width           =   2775
      End
      Begin VB.Label lblPatTime1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   615
         Left            =   11760
         TabIndex        =   14
         Top             =   3720
         Width           =   3255
      End
      Begin VB.Label lblTiming1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   615
         Left            =   11760
         TabIndex        =   13
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Label lblName1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   495
         Left            =   2400
         TabIndex        =   12
         Top             =   1920
         Width           =   6135
      End
      Begin VB.Label lblDegrees 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Qualification"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label lblAddress 
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
         ForeColor       =   &H80000006&
         Height          =   615
         Left            =   600
         TabIndex        =   10
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label lblPhNumber 
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
         ForeColor       =   &H80000006&
         Height          =   375
         Left            =   9840
         TabIndex        =   9
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label lblTiming 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Arrival Time"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   855
         Left            =   9840
         TabIndex        =   8
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label lblPatTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Departure Time"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   735
         Left            =   9840
         TabIndex        =   7
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label lblAvailable 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Available Days"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   855
         Left            =   600
         TabIndex        =   6
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label lblName 
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
         ForeColor       =   &H80000006&
         Height          =   495
         Left            =   600
         TabIndex        =   5
         Top             =   1920
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Search Record"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   615
      Left            =   6840
      TabIndex        =   52
      Top             =   360
      Width           =   2655
   End
   Begin VB.Line Line3 
      X1              =   3120
      X2              =   3120
      Y1              =   8760
      Y2              =   8880
   End
End
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim query As String
Dim DocAvail As Integer
Private Sub back_Click()
Me.Hide
Home.Show
End Sub
Private Sub cmdSearch_Click()
con.Open ("Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + App.Path + "\opdDbms.mdb;")
If CmbSearch1.ListIndex = 0 Then
query = "select * from Doctor where docID='" + txtDocid.Text + "'"
Else: query = "select * from Doctor where docName='" + txtDocid.Text + "'"
End If
Set rs = con.Execute(query)
If rs.EOF = False And rs.BOF = False Then
lblId1.Caption = rs(0)
lblName1.Caption = rs(1)
DocAvail = Val(rs(2))
Dim Index As Integer
While DocAvail <> 0
 Index = DocAvail Mod 10
 chkDay(Index - 1).Value = Checked
 DocAvail = DocAvail / 10
Wend
lblTiming1.Caption = rs(3)
lblPatTime1.Caption = rs(4)
lblPhNumber1.Caption = rs(5)
lblAddress1.Caption = rs(6)
lblDegrees1.Caption = rs(7)
Else: MsgBox "No Record Found", vbExclamation
End If
Set rs = Nothing
con.Close
End Sub

Private Sub cmdSearch1_Click()
con.Open ("Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + App.Path + "\opdDbms.mdb;")
If CmbSearch.ListIndex = 0 Then
query = "select * from patientDetails where patientID= '" + txtPatid.Text + "'"
Else: query = "select * from patientDetails where patientName= '" + txtPatid.Text + "'"
End If
Set rs = con.Execute(query)
If rs.EOF = False And rs.BOF = False Then
LblIdP1.Caption = rs(0)
lblNameP1.Caption = rs(1)
lblDepartment1.Caption = rs(2)
lblPhNumberP1.Caption = rs(3)
lblAddressP1.Caption = rs(4)
lblAltContact1.Caption = rs(5)
lblEnterTime1.Caption = rs(6)
lblDocName1.Caption = rs(7)
lblAmountP1.Caption = rs(8)
lblDateP1.Caption = rs(9)
Else: MsgBox "No Record Found", vbExclamation
End If
Set rs = Nothing
con.Close
End Sub

Private Sub Combo1_Change()
MsgBox "hello"
End Sub

Private Sub Form_Activate()
optPat.Value = True
framePat.Visible = True
frameDoc.Visible = False
End Sub

Private Sub optDoc_Click()
If optDoc.Value = True Then
framePat.Visible = False
frameDoc.Visible = True
End If
End Sub

Private Sub optPat_Click()
If optPat.Value = True Then
framePat.Visible = True
frameDoc.Visible = False
End If
End Sub

