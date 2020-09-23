VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgCom 
      Left            =   6120
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Parsed Text: "
      Height          =   1935
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   6615
      Begin VB.TextBox txtParsed 
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Top             =   240
         Width           =   6375
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5925
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton Command4 
         Caption         =   "Clear All"
         Height          =   375
         Left            =   2400
         TabIndex        =   14
         Top             =   5400
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Caption         =   "Options: "
         Height          =   1095
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   6615
         Begin VB.CommandButton Command5 
            Caption         =   "Open A File"
            Height          =   375
            Left            =   5520
            TabIndex        =   16
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtDelimCh 
            Height          =   285
            Left            =   1080
            TabIndex        =   11
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox txtDelim 
            Height          =   285
            Left            =   1080
            TabIndex        =   10
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "\p = Line Break   \t = Tab  "
            Height          =   255
            Left            =   1560
            TabIndex        =   15
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Change to:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Delimiter:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         Height          =   375
         Left            =   3480
         TabIndex        =   6
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Clear Results"
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   5400
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Parse"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   5400
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Text To Parse: "
         Height          =   1935
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   6615
         Begin VB.TextBox txtParse 
            Height          =   1575
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   7
            Top             =   240
            Width           =   6375
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iCommas As Integer
Dim i As Integer
Dim iPos As Integer
Dim iPosStart As Integer
Dim iLen As Integer
Dim X As Integer
Dim iFile As Integer

Dim sDelim As String
Dim strBefore(64333) As String
Dim strAfter As String
Dim sFile As String


Public Function parseStr(g_strText) As Boolean
  'parse that string
  On Error GoTo Errchk
  
  Dim bRetVal As Boolean
  bRetVal = True
  iPosStart = 1
  iLen = Len(g_strText)

  'make sure there is a string...
  If (iLen > 0) Then
      iPos = InStr(iPosStart, g_strText, sDelim)
      'k now lets make sure we found it
      If (iPos <> 0) Then
        iPosStart = iPos
        strBefore(i) = Left(g_strText, iPos - 1)
        'MsgBox ("strBefore:" & strBefore & vbCrLf & "i:" & i)
        
        strAfter = Right(g_strText, iLen - iPosStart)
        'MsgBox (strAfter)
      Else
      'didnt find anymore... so we are done
        bRetVal = False
        Exit Function
      End If
  Else
    bRetVal = False
    Exit Function
  End If
  
  'keep moving AFTER the comma
  g_strText = strAfter
  'MsgBox ("strBefore:" & strBefore & vbCrLf & "i:" & i)
  parseStr = bRetVal
  i = i + 1
  iCommas = iCommas + 1
  
Exit Function
Errchk:
MsgBox ("Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description & vbCrLf & "Error Source: " & Err.Source)
End Function

Private Sub Command1_Click()
 On Error GoTo Errchk
'parse the string while it returns a true value _
 meaning there are still commas in the file
 
    'set variables in use back to nothing
    txtParsed.Text = ""
    i = 0
    X = 0
    iCommas = 0
    
    'set string
    sDelim = txtDelim.Text
    sDelimch = txtDelimCh.Text
    
    'change certain chars
    If sDelimch = "\p" Then
        sDelimch = vbCrLf
    ElseIf sDelimch = "\t" Then
        sDelimch = vbTab
    End If
    
    'initialize the function
    g_strText = txtParse.Text
    
    'make sure there is text to process
    If Len(g_strText) < 1 Then
        MsgBox ("There must be text in the text box")
        Exit Sub
    End If
        
    'go at it
    While (parseStr(g_strText))
        DoEvents
    Wend
    
    'populate the results textbox
    For X = 0 To i - 1
        txtParsed.Text = txtParsed.Text & strBefore(X) & sDelimch
    Next
    
    'grab the end of the string
    txtParsed.Text = txtParsed.Text & strAfter
    
    'display in status bar
    StatusBar1.Panels(1).Text = "Delimiter #: " & i
    StatusBar1.Panels(2).Text = "Delimiter: " & Chr(34) & sDelim & Chr(34) & "   Changed to: " & Chr(34) & sDelimch & Chr(34)
Exit Sub
Errchk:
MsgBox ("Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description & vbCrLf & "Error Source: " & Err.Source)
End Sub

Private Sub Command2_Click()
 On Error GoTo Errchk
 
    txtParsed.Text = ""
    
Exit Sub
Errchk:
MsgBox ("Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description & vbCrLf & "Error Source: " & Err.Source)
End Sub

Private Sub Command3_Click()
 On Error GoTo Errchk
 
    Unload Me
    
Exit Sub
Errchk:
MsgBox ("Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description & vbCrLf & "Error Source: " & Err.Source)
End Sub

Private Sub Command4_Click()
 On Error GoTo Errchk
 
    txtParsed.Text = ""
    txtParse.Text = ""
    
Exit Sub
Errchk:
MsgBox ("Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description & vbCrLf & "Error Source: " & Err.Source)
End Sub

Private Sub Command5_Click()
 On Error GoTo Errchk
 
    iFile = FreeFile
    
    dlgCom.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    dlgCom.ShowOpen
    'msgBox (dlgCommon.FileName)
    sFile = dlgCom.FileName
    
    Open sFile For Input As #iFile
    
        txtParse.Text = Input(LOF(iFile), iFile)
   
    Close #iFile

Exit Sub
Errchk:
MsgBox ("Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description & vbCrLf & "Error Source: " & Err.Source)
End Sub

Private Sub Form_Load()
    frmMain.Caption = "[HM] String Parse - v" & App.Major & "." & App.Minor
    txtDelim.Text = ","
    txtDelimCh.Text = "\p"
    txtParse.Text = "This is,my,test,for,you."
    
End Sub

