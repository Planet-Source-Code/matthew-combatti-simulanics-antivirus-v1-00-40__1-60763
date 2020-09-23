VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SimAV Signature Creator"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   ControlBox      =   0   'False
   Enabled         =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin SIMAVSigCreator.XPFrame XPFrame4 
      Height          =   2775
      Left            =   4920
      TabIndex        =   15
      Top             =   3720
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4895
      Caption         =   "Signature Database Editor"
      Begin SIMAVSigCreator.lvButtons_H Command2 
         Height          =   495
         Left            =   360
         TabIndex        =   22
         Top             =   2160
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   873
         Caption         =   "Add Signature and Virus Name to Simulanics Antivirus Database"
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
      Begin SIMAVSigCreator.Search txtVirus 
         Height          =   315
         Left            =   1080
         TabIndex        =   21
         Top             =   1560
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         Text            =   ""
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Be sure to look for a unique signature to the virus only. It's not that hard!"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Virus Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1640
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   5280
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "What is the name of the Virus this Signature comes from?"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   5175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Signature:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Select A Signature From the left by double clicking on it, then procede."
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   5175
      End
   End
   Begin SIMAVSigCreator.XPFrame XPFrame3 
      Height          =   2415
      Left            =   4920
      TabIndex        =   9
      Top             =   1320
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4260
      Caption         =   "Obtaining The Signature"
      Begin SIMAVSigCreator.lvButtons_H lvButtons_H2 
         Height          =   495
         Left            =   4080
         TabIndex        =   25
         Top             =   1800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "Stop Reading File"
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
      Begin SIMAVSigCreator.lvButtons_H cmdRefresh 
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   873
         Caption         =   "Gather Signatures From Selected File"
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
      Begin SIMAVSigCreator.Search txtFilename 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   556
         Text            =   ""
      End
      Begin SIMAVSigCreator.lvButtons_H Command1 
         Height          =   375
         Left            =   3840
         TabIndex        =   11
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Browse For File"
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
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected File:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Begin By Selecting A File To Scan For Signatures."
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   3855
      End
   End
   Begin SIMAVSigCreator.XPFrame XPFrame2 
      Height          =   1335
      Left            =   4920
      TabIndex        =   6
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2355
      Caption         =   "Signature Settings"
      Begin SIMAVSigCreator.lvButtons_H lvButtons_H1 
         Height          =   495
         Left            =   4320
         TabIndex        =   24
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         Caption         =   "Close"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   4
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
         cBack           =   16744576
      End
      Begin VB.TextBox ValidCharSet 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "abcdefghijklmnopqrstuvwxyzüöäÜÖÄß_-' 1234567890"
         Top             =   720
         Width           =   5145
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valid Character Set To Search With:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2580
      End
   End
   Begin MSComDlg.CommonDialog cmd1 
      Left            =   0
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pProgressBar 
      AutoRedraw      =   -1  'True
      Height          =   285
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   10275
      TabIndex        =   0
      Top             =   6480
      Visible         =   0   'False
      Width           =   10335
   End
   Begin SIMAVSigCreator.XPFrame XPFrame1 
      Height          =   6495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   11456
      Caption         =   "Signatures List"
      Begin SIMAVSigCreator.lvButtons_H cmdSearch 
         Height          =   735
         Left            =   3600
         TabIndex        =   5
         Top             =   5640
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1296
         Caption         =   "Go"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   7
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
         Image           =   "frmMain.frx":030A
         ImgSize         =   48
         cBack           =   -2147483633
      End
      Begin SIMAVSigCreator.Search txtSearch 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   5880
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         Text            =   ""
      End
      Begin VB.ListBox lStrings 
         Appearance      =   0  'Flat
         Height          =   5100
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   4665
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   5925
         Width           =   555
      End
   End
   Begin VB.Menu menuEdit 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuCountThis 
         Caption         =   "Count This"
      End
      Begin VB.Menu mnuRemoveAll 
         Caption         =   "Remove All"
      End
      Begin VB.Menu mnuRemoveDuplicates 
         Caption         =   "Remove Duplicates"
      End
      Begin VB.Menu mnuRemoveThis 
         Caption         =   "Remove This"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear List"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FileName As String
Dim FileNr



Private Sub cmdRefresh_Click()
  lStrings.Clear
  lStrings.SetFocus
  ReadFile
End Sub



Private Sub cmdSearch_Click()
  Dim i As Long, J As Long
  If Len(txtSearch) = 0 Then
    'If txtSearch.Enabled = False Then txtSearch.Enabled = True
    txtSearch.SetFocus
    Exit Sub
  End If
  cmdSearch.Enabled = False
  'txtSearch.Enabled = False
  J = lStrings.ListIndex + 1
  If J < 0 Or J = lStrings.ListCount Then J = 0
  For i = J To lStrings.ListCount - 1
    If InStr(UCase(lStrings.List(i)), UCase(txtSearch)) Then
      lStrings.ListIndex = i
      Exit For
    End If
    DoEvents
  Next i
  cmdSearch.Enabled = True
  'txtSearch.Enabled = True
End Sub

Private Sub Command1_Click()
cmd1.ShowOpen
FileName = cmd1.FileName
txtFilename.Text = FileName
End Sub

Private Sub Command2_Click()
Dim y As Integer
Dim x As String

On Error GoTo FIXIT
GoTo STARTER
FIXIT:
    MsgBox "Signature was not added. The signature file could not be located." & vbCrLf & "Be sure that SimAV Signature Creator.exe is in the SimAVSigCreator Folder in the Simulanics Antivirus Folder!", vbOKOnly, "Internal Error ~ Try Again"
    Exit Sub
    
STARTER:


Open App.Path & "\signatures.db" For Append As #1
    Print #1, txtVirus.Text & " '~#~' " & Label2.Caption
Close #1
txtVirus.Text = ""
Label2.Caption = ""
MsgBox "Virus Signature Added Successfully", vbOKOnly, "Simulanics Virus Signature Creator"

End Sub

Private Sub Command3_Click()

End Sub

Private Sub ReadFile()
  Dim Buffer As String, BufferLen As Long, FileNr As Integer
  Dim Tmp As String, i As Long
  FileNr = FreeFile
  On Local Error Resume Next
  Open FileName For Input As #FileNr
  If Err Then
    Err.Clear
    Exit Sub
  End If
  Close #FileNr
  cmdRefresh.Enabled = False
  ValidCharSet.Enabled = False
  pProgressBar.Visible = True
  pProgressBar.Cls
  Open FileName For Binary Access Read As #FileNr
  BufferLen = 256
  Do Until EOF(FileNr)
    If BufferLen > LOF(FileNr) - Loc(FileNr) Then
      BufferLen = LOF(FileNr) - Loc(FileNr)
      If BufferLen < 1 Or Err Then Exit Do
    End If
    pProgressBar.Line (0, 0)-(pProgressBar.Width / LOF(FileNr) * Loc(FileNr), _
                              pProgressBar.Height), vbBlue, BF
    pProgressBar.Refresh
    Buffer = Space(BufferLen)
    Get #FileNr, , Buffer
    For i = 1 To Len(Buffer)
      If InStr(ValidCharSet, LCase(Mid(Buffer, i, 1))) Then
        Tmp = Tmp & Mid(Buffer, i, 1)
      Else
        If Len(Tmp) > 3 Then
          lStrings.AddItem Trim(Tmp)
          
        End If
        Tmp = ""
      End If
    Next i
    DoEvents
  Loop
  Close FileNr
  If Err Then Err.Clear
  pProgressBar.Visible = False
  cmdRefresh.Enabled = True
  ValidCharSet.Enabled = True
End Sub

Private Sub Form_Load()
Me.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub lStrings_DblClick()
  Dim lIndex As Long
  lIndex = lStrings.ListIndex + 1
  If lIndex Then
    Label2.Caption = lStrings.Text
  End If
End Sub

Private Sub lStrings_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim lIndex As Long
  lIndex = lStrings.ListIndex + 1
  Select Case KeyCode
  Case 46 '[DEL]'
    If Shift = 0 Then
      mnuRemoveAll_Click
    Else
      mnuRemoveDuplicates_Click
      TrySelectItem lIndex
    End If
  Case 10, 13 '[ENTER]'
    lStrings_DblClick
  End Select
End Sub

Private Sub lStrings_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then
    PopupMenu menuEdit
  End If
End Sub

Private Sub lvButtons_H1_Click()
Unload Me
End
End Sub

Private Sub lvButtons_H2_Click()
Close #FileNr
MsgBox "Gathering of Signatures from " & FileName & " was terminated by the user.", vbOKOnly, "SimAV Signature Creator"

End Sub

Private Sub mnuClear_Click()
  lStrings.Clear
End Sub

Private Sub mnuCopy_Click()
  Dim lIndex As Long
  lIndex = lStrings.ListIndex + 1
  If lIndex Then
    Clipboard.Clear
    Clipboard.SetText lStrings
  End If
End Sub

Private Sub mnuCountThis_Click()
  Dim ItemText As String
  Dim i As Long
  Dim c As Long
  
  i = lStrings.ListIndex + 1
  If i Then
    ItemText = UCase(Trim(lStrings))
    For i = 0 To lStrings.ListCount - 1
      If UCase(Trim(lStrings.List(i))) = ItemText Then c = c + 1
    Next i
    MsgBox "Count of """ & ItemText & """ is " & c
  End If
End Sub

Private Sub mnuRemoveAll_Click()
  Dim lIndex As Long
  lIndex = lStrings.ListIndex + 1
  If lIndex Then
    RemoveItem lStrings
    TrySelectItem lIndex - 1
  End If
End Sub

Private Sub mnuRemoveDuplicates_Click()
  Dim lIndex As Long
  lIndex = lStrings.ListIndex + 1
  If lIndex Then
    RemoveItem lStrings, lIndex
    TrySelectItem lIndex - 1
  End If
End Sub

Private Sub mnuRemoveThis_Click()
  Dim lIndex As Long
  lIndex = lStrings.ListIndex + 1
  If lIndex Then
    lStrings.RemoveItem lIndex - 1
    TrySelectItem lIndex - 1
  End If
End Sub



Private Sub RemoveItem(ByVal ItemText As String, Optional ByVal ExceptIndex As Long)
  Dim i As Long
  ItemText = UCase(Trim(ItemText))
  While i < lStrings.ListCount
    If i + 1 <> ExceptIndex Then
      If UCase(Trim(lStrings.List(i))) = ItemText Then
        lStrings.RemoveItem i
        i = i - 1
      End If
    End If
    i = i + 1
  Wend
End Sub


Private Sub TrySelectItem(ByVal lIndex As Long)
  If lStrings.ListCount < 1 Then Exit Sub
  If lStrings.ListCount < (lIndex + 1) Then lIndex = lStrings.ListCount - 1
  If lIndex < 0 Then lIndex = 0
  lStrings.ListIndex = lIndex
End Sub

