VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Directory To Scan"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7095
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SimulanicsAntivirus.XPFrame XPFrame1 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   6165
      BorderColor     =   0
      Caption         =   "Select Directory To Scan"
      Begin VB.DirListBox Dir1 
         Height          =   2565
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   2655
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2655
      End
   End
   Begin SimulanicsAntivirus.XPFrame XPFrame2 
      Height          =   3495
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   6165
      BorderColor     =   16777215
      Caption         =   "Scan Info:"
      Begin SimulanicsAntivirus.lvButtons_H Command1 
         Height          =   375
         Left            =   360
         TabIndex        =   5
         ToolTipText     =   "Go Back To Main AV Window"
         Top             =   3000
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
      Begin SimulanicsAntivirus.lvButtons_H Command2 
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         ToolTipText     =   "Procede To Scan Selected Directory"
         Top             =   3000
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "Begin"
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
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "C:\"
         Height          =   1215
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   3975
      End
      Begin VB.Image Image1 
         Height          =   1440
         Left            =   1440
         Picture         =   "Form2.frx":0000
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   1425
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Hide
End Sub

Private Sub Command2_Click()

Form2.Hide
Form1.txtYesNo.Text = "Yes"
Form3.Label1.Caption = Form1.File1.Path
Form3.Command1.Caption = "Start"
Form3.Label8.Caption = "0"
Form3.Label3.Caption = "0 of " & Form1.File1.ListCount
Form3.Show
DoEvents
End Sub

Private Sub Dir1_Change()
Form1.File1.Path = Form2.Dir1.Path
Label2.Caption = Dir1.Path
Form3.Label1.Caption = Dir1.Path

End Sub

Private Sub Dir1_Click()
Form1.File1.Path = Form2.Dir1.Path
Label2.Caption = Dir1.Path
End Sub

Private Sub Dir1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Caption = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
Label2.Caption = Dir1.Path
End Sub

