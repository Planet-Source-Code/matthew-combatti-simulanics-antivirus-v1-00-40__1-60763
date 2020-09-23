VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "About Simulanics Antivirus"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   5400
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SimulanicsAntivirus.XPFrame XPFrame1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9551
      Caption         =   ""
      Begin SimulanicsAntivirus.lvButtons_H lvButtons_H1 
         Height          =   375
         Left            =   7200
         TabIndex        =   2
         Top             =   4800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Finished"
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
         Caption         =   $"Form4.frx":0000
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   8535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"Form4.frx":00CD
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   1
         Top             =   4080
         Width           =   8295
      End
      Begin VB.Image Image1 
         Height          =   3840
         Left            =   600
         Picture         =   "Form4.frx":0181
         ToolTipText     =   "Owned and Operated By Matthew Combatti"
         Top             =   360
         Width           =   7575
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lvButtons_H1_Click()
Form4.Hide

End Sub
