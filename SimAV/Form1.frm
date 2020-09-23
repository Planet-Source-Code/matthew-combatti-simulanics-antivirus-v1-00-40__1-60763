VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simulanics Antivirus"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7920
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin SimulanicsAntivirus.XPFrame XPFrame1 
      Height          =   1215
      Left            =   0
      TabIndex        =   19
      Top             =   5520
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   2143
      Caption         =   "Control Panel"
      Begin SimulanicsAntivirus.lvButtons_H lvButtons_H6 
         Height          =   855
         Left            =   5400
         TabIndex        =   27
         ToolTipText     =   "Open SimAV Signature Creator"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1508
         Caption         =   "Database Editor"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "Form1.frx":030A
         ImgSize         =   48
         cBack           =   -2147483633
      End
      Begin SimulanicsAntivirus.lvButtons_H cmdVault 
         Height          =   855
         Left            =   0
         TabIndex        =   20
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1508
         Caption         =   "Virus Vault"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "Form1.frx":335C
         ImgSize         =   48
         cBack           =   -2147483633
      End
      Begin SimulanicsAntivirus.lvButtons_H lvButtons_H2 
         Height          =   855
         Left            =   4200
         TabIndex        =   21
         ToolTipText     =   "Visit http://www.simulanics.com"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1508
         Caption         =   "Web Site"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "Form1.frx":63AE
         ImgSize         =   48
         cBack           =   -2147483633
      End
      Begin SimulanicsAntivirus.lvButtons_H lvButtons_H3 
         Height          =   855
         Left            =   2760
         TabIndex        =   22
         ToolTipText     =   "Download the Latest Signature Database"
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1508
         Caption         =   "Update Definitions"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "Form1.frx":9400
         ImgSize         =   48
         cBack           =   -2147483633
      End
      Begin SimulanicsAntivirus.lvButtons_H lvButtons_H4 
         Height          =   855
         Left            =   1320
         TabIndex        =   23
         ToolTipText     =   "Got A Virus? What's it do? Find Out Here"
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1508
         Caption         =   "Virus Information"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   4
         Image           =   "Form1.frx":C452
         ImgSize         =   48
         cBack           =   -2147483633
      End
      Begin SimulanicsAntivirus.lvButtons_H lvButtons_H5 
         Height          =   855
         Left            =   6600
         TabIndex        =   24
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1508
         Caption         =   "About"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "Form1.frx":F4A4
         ImgSize         =   32
         cBack           =   -2147483633
      End
   End
   Begin SimulanicsAntivirus.lvButtons_H cmdSearchDirectory 
      Height          =   1455
      Left            =   1440
      TabIndex        =   3
      ToolTipText     =   "Scan All Files Within A Directory"
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2566
      Caption         =   "Scan Directory"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Form1.frx":FBF6
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin SimulanicsAntivirus.lvButtons_H cmdSearch 
      Height          =   1455
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Scan a Single File"
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2566
      Caption         =   "Scan File"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "Form1.frx":12C48
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin VB.Frame VirusTable 
      Caption         =   "Virus Table"
      Height          =   6495
      Left            =   8160
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.TextBox txtYesNo 
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Text            =   "Yes"
         Top             =   6120
         Width           =   1095
      End
      Begin VB.ListBox List1 
         Height          =   5520
         Left            =   2640
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
      Begin VB.FileListBox File1 
         Height          =   5550
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Infected"
         Height          =   255
         Left            =   2640
         TabIndex        =   26
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Scanning Directory"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSComDlg.CommonDialog CMDLG1 
      Left            =   7440
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SimulanicsAntivirus.XPFrame Frame2 
      Height          =   1455
      Left            =   2880
      TabIndex        =   5
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2566
      Caption         =   "Browse For File"
      Begin SimulanicsAntivirus.Search txtFile 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         Text            =   ""
      End
      Begin SimulanicsAntivirus.lvButtons_H Command1 
         Height          =   375
         Left            =   3600
         TabIndex        =   7
         ToolTipText     =   "Select A File To Scan"
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Browse"
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
      Begin VB.Label lblFileName 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "File Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
   End
   Begin SimulanicsAntivirus.XPFrame Frame1 
      Height          =   1215
      Left            =   0
      TabIndex        =   11
      Top             =   4320
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   2143
      Caption         =   "Antivirus Definition Database"
      Begin SimulanicsAntivirus.lvButtons_H lvButtons_H1 
         Height          =   615
         Left            =   6720
         TabIndex        =   18
         ToolTipText     =   "Close Simulanics AV"
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         Caption         =   "Exit"
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
      Begin VB.Label lblvtot 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Virus Signatures"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "DB Info 2"
         Height          =   255
         Left            =   4320
         TabIndex        =   13
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label lblAVDB 
         BackStyle       =   0  'Transparent
         Caption         =   "DB Info"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   4095
      End
   End
   Begin SimulanicsAntivirus.XPFrame Frame3 
      Height          =   2895
      Left            =   0
      TabIndex        =   8
      Top             =   1440
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5106
      Caption         =   "Scan Log"
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7200
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":15C9A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":168EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":16C0A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2415
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   4260
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         PictureAlignment=   5
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   0
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Virus Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "File Location"
            Object.Width           =   2540
         EndProperty
         Picture         =   "Form1.frx":174E6
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   7095
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Scanning is in Progress Please Wait!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   600
            TabIndex        =   10
            Top             =   600
            Width           =   5895
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWNOACTIVE = 4
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_RESTORE = 9
 
Dim Signature As String
Dim SigName As String
Dim TotalSigs As Integer
Dim DispDate As String
Dim RunCheck As Integer
Dim Intermediary As String
Dim SigSet(10000) As String
Dim NameSet(10000) As String
Dim TotSigs As Integer
Private b1() As Byte
Private b2() As Byte
Dim lItem As ListView
Dim AA

Private Sub cmdSearch_Click()
On Error Resume Next
Dim FileToScan As Integer
Dim totfile As Integer

ListView1.ListItems.Clear

If txtFile.Text = "" Then
    MsgBox "Please Select A File To Scan Before Attempting To Scan!", vbOKOnly, "File Selection?"
    Exit Sub
Else
End If

ListView1.Visible = False
Frame4.Visible = True
DoEvents
totfile = 0
FileToScan = 0

        
'Open txtFile.Text For Binary As #1
'    AA = Input(LOF(1), 1)
'Close #1
AA = GetFileQuick(txtFile.Text, True)

DoEvents

For FileToScan = 1 To TotSigs

If ((InStr(1, AA, SigSet(FileToScan), vbTextCompare)) > 0) = True Then
    'txtLOG.Text = txtLOG.Text & NameSet(FileToScan) & vbTab & vbTab & vbTab & vbTab & txtFile.Text & vbCrLf
    ListView1.ListItems.Add(, , NameSet(FileToScan), 3, 3).ListSubItems.Add , , txtFile.Text
    
End If


Next FileToScan

Frame4.Visible = False

txtFile.Text = ""
ListView1.Visible = True
Frame4.Visible = False
End Sub

Public Sub SearchDirectory(Direct As String)
On Error Resume Next
Dim FileToScan As Integer
Dim totfile As Integer
Dim FileLocation As String

File1.Path = Direct

For FileToScan = 0 To (File1.ListCount - 1)
'NextFileToScan:
        'totfile = 0
        File1.Selected(FileToScan) = True
        
If txtYesNo.Text = "No" Then
    Close #1
    Exit Sub
End If


FileLocation = File1.Path & "\" & File1.FileName
FileLocation = Replace(FileLocation, "\\", "\")


Form3.Label3.Caption = (FileToScan + 1) & " of " & Form1.File1.ListCount 'scanning progress
Form3.Label5.Caption = GetRoundedKB(FileLen(FileLocation))


Form3.Label1.Caption = FileLocation
'Open FileLocation For Binary As #1
'    AA = Input(LOF(1), 1)
'Close #1
AA = GetFileQuick(FileLocation, True)

DoEvents

For totfile = 1 To TotSigs

If ((InStr(1, AA, SigSet(totfile), vbTextCompare)) > 0) = True Then
    Form3.Label8.Caption = Form3.Label8.Caption + 1
    ListView1.ListItems.Add(, , NameSet(totfile), 3, 3).ListSubItems.Add , , FileLocation
    
End If

Next totfile

Next FileToScan
Form3.lblEndTime.Caption = Time$
Form3.Command1.Caption = "Finished"
Form3.Label1.Caption = "Scanning is Now Complete"


End Sub

Private Sub cmdSearchDirectory_Click()
ListView1.ListItems.Clear

Form2.Show

End Sub

Private Sub Command1_Click()
ListView1.ListItems.Clear
CMDLG1.Filter = "All Files (*.*)|*.*"
CMDLG1.ShowOpen
txtFile.Text = CMDLG1.FileName

End Sub

Public Function CheckDBDate()
On Error GoTo FixTheProblem
Dim DBDate As String
Dim OLDDate As Date
Dim NewDate As Date
GoTo GetStarted

FixTheProblem:
    Select Case Err.Number
        Case 62
            Close #2
            Kill App.Path & "\signatures.db"
            DoEvents
            FileCopy App.Path & "\signatures.db.old", App.Path & "\signatures.db"
            DoEvents
            Resume Next
    End Select

GetStarted:
Open App.Path & "\signatures.db" For Input As #2
    Input #2, DBDate
Close #2

OLDDate = DBDate
NewDate = Date$

DispDate = DateDiff("D", OLDDate, NewDate, vbUseSystemDayOfWeek, vbUseSystem)

If DispDate = 0 Then
    lblAVDB.Caption = "Antivirus Database Updated Today"
    Exit Function
ElseIf DispDate < 0 Then 'if system time has been edited then unload AV
    MsgBox "Please fix your system time and reload the Antivirus", vbOKOnly, "Antivirus"
    Unload Me
    End
Else
End If

If DispDate < 15 Then
    lblAVDB.Caption = "Antivirus Database is Up to Date."
Else
    lblAVDB.Caption = "Avtivirus Database is Outdated."
End If


End Function

Public Function LoadAVDB()


Dim Vari As String
TotSigs = 0

Open App.Path & "\signatures.db" For Input As #1
    Do
        Line Input #1, Signature
    Loop Until Signature = "##Signatures##"
        
    Do
        TotSigs = TotSigs + 1
        Line Input #1, Vari
        
    NameSet(TotSigs) = Split(Vari, " '~#~' ")(0)
    SigSet(TotSigs) = Split(Vari, " '~#~' ")(1)
        'Debug.Print SigSet(TotSigs) & "    " & NameSet(TotSigs)
        
    Loop Until EOF(1)
Close #1

End Function


Private Sub Form_Load()
Dim Totv As Integer
TotSigs = 0
Totv = 0


    With ListView1
        .ColumnHeaders(1).Width = (ListView1.Width / 2.5)
        .ColumnHeaders(2).Width = ListView1.Width
        
    End With


CheckDBDate
DoEvents
LoadAVDB
DoEvents
Open App.Path & "\signatures.db" For Input As #1
    Do
        Input #1, Signature
    Loop Until Signature = "##Signatures##"
    Do
        Line Input #1, Signature
        Totv = Totv + 1
    Loop Until EOF(1)
Close #1

lblvtot.Caption = Totv
Label2.Caption = "Database is " & DispDate & " days old."
    
    
End Sub


Private Sub lvButtons_H1_Click()
Unload Me
End
End Sub


Private Sub lvButtons_H2_Click()
Dim opbrowser As Long

opbrowser = ShellExecute(Me.hWnd, "open", "http://www.simulanics.com", "", "C:\", SW_SHOWNORMAL)

End Sub

Private Sub lvButtons_H3_Click()
Form5.Show
End Sub

Private Sub lvButtons_H4_Click()
Dim opbrowser As Long

opbrowser = ShellExecute(Me.hWnd, "open", "http://vil.nai.com/vil/", "", "C:\", SW_SHOWNORMAL)

End Sub

Private Sub lvButtons_H5_Click()
Form4.Show
End Sub

Private Sub lvButtons_H6_Click()
On Error GoTo SIGEDITORNOTFOUND

GoTo OpenEditor:

SIGEDITORNOTFOUND:
    MsgBox "The Signature Creator was not found. Be sure that 'SimAV Signature Creator.exe' is in the Simulanics Antivirus Folder.", vbOKOnly, "Could Not Be Found"
    Exit Sub
    
OpenEditor:
    Shell App.Path & "\SimAV Signature Creator.exe", vbNormalFocus
End Sub


