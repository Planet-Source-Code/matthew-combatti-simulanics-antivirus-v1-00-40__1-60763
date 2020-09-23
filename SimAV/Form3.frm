VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   Caption         =   "Scanning In Progress"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   2535
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SimulanicsAntivirus.XPFrame XPContainer1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4471
      HeaderDarkColor =   14737632
      BorderColor     =   16777215
      Caption         =   "Progress"
      Begin SimulanicsAntivirus.lvButtons_H lvButtons_H1 
         Height          =   375
         Left            =   6000
         TabIndex        =   11
         ToolTipText     =   "Cancel Scanning and Return To Main AV Window"
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "Cancel"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin SimulanicsAntivirus.Check Check1 
         Height          =   300
         Left            =   1560
         TabIndex        =   7
         Top             =   60
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   529
      End
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   5400
         Top             =   2040
      End
      Begin SimulanicsAntivirus.lvButtons_H Command1 
         Height          =   375
         Left            =   6000
         TabIndex        =   4
         Top             =   2040
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "Start"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label lblEndTime 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         Height          =   255
         Left            =   3240
         TabIndex        =   15
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblStartTime 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Ending Time:"
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Time:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Threats Found:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Minimize AV Windows Upon Scanning Initialization && Reshow upon Completion"
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   120
         Width           =   5775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   4440
         TabIndex        =   6
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "File Size (KB)"
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "0 of 0"
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Files To Scan:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   7575
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Select Case Command1.Caption
        Case "Start"
        If Check1.Checked = True Then
            Form1.WindowState = 1
            Form3.WindowState = 1
        Else
        End If
            Form1.Enabled = True
            Form3.Command1.Caption = "Please Wait..."
            DoEvents
            lblStartTime.Caption = Time$
            Form1.SearchDirectory Label2.Caption
           
        Case "Finished"
            Timer1.Enabled = False
            Form1.Enabled = True
            Command1.Caption = "Start"
            Label2.Caption = ""
            Form1.Show
            Form3.Hide
        Case "Please Wait..."
            Exit Sub
            
    End Select


End Sub

Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub lvButtons_H1_Click()

            Form1.txtYesNo.Text = "No"
            Timer1.Enabled = False
            Form1.Enabled = True
            Command1.Caption = "Start"
            Label2.Caption = ""
            Form1.Show
            Form3.Hide
            Close #1
            
End Sub

Private Sub Timer1_Timer()
If Label1.Caption = "Scanning is Now Complete" Then
    Command1.Caption = "Finished"
Else
End If
End Sub

