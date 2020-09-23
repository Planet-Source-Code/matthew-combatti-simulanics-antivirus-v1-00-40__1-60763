VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form5 
   Caption         =   "Update Virus Definitions"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   ScaleHeight     =   2865
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SimulanicsAntivirus.XPFrame XPFrame1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5106
      Caption         =   "Automatic Updates"
      Begin SimulanicsAntivirus.lvButtons_H lvButtons_H1 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Close"
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
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   5280
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin SimulanicsAntivirus.lvButtons_H lvButtons_H2 
         Height          =   375
         Left            =   4320
         TabIndex        =   3
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Begin Update"
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "You must be connected to the internet to Update.  If you are not connected to the internet, do so now..."""
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Database Update Information Center"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1560
         Width           =   5775
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub GetUpdateVersion()
Dim UpdateVersion As String
    
End Sub


Public Sub GetUpdate()
Dim UpDate As String
    lvButtons_H1.Enabled = False
    lvButtons_H2.Enabled = True
    lblInfo.Caption = "Update Being Retrieved"
    UpDate = Inet1.OpenURL("http://members.lycos.co.uk/simulanics/update/signatures.txt")
    DoEvents
    lblInfo.Caption = "Backing Up Existing Signature Database"
    FileCopy App.Path & "\signatures.db", App.Path & "\signatures.db.old"
    DoEvents
    Kill App.Path & "\signatures.db"
    DoEvents
    Open App.Path & "\signatures.db" For Output As #11
        Print #11, UpDate
    Close #11
    lblInfo.Caption = "Update is Complete. Restart the Antivirus to load the new Database."
    lvButtons_H1.Enabled = True
End Sub


Private Sub Form_Load()
Form1.Enabled = False
End Sub


Private Sub lvButtons_H1_Click()
Form5.Hide
Form1.Enabled = True
End Sub

Private Sub lvButtons_H2_Click()
GetUpdate
End Sub
