VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTransStr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Auto Translate - Strings"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frmTransStr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbLang 
      Height          =   315
      ItemData        =   "frmTransStr.frx":000C
      Left            =   1800
      List            =   "frmTransStr.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Width           =   2175
   End
   Begin VB.ListBox lstO 
      Height          =   2085
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   960
      Width           =   3975
   End
   Begin VB.CommandButton cmdSelAll 
      Caption         =   "Select &All"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelNone 
      Caption         =   "Select &None"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdTranslate 
      Caption         =   "&Translate"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar progress 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Languages (from, to):"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Select what strings you want to translate:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   3855
   End
End
Attribute VB_Name = "frmTransStr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sALang() As String, iLangs As Integer

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub cmdSelAll_Click()
    Dim i As Integer
    
    For i = 0 To lstO.ListCount - 1
        lstO.Selected(i) = True
    Next
    
    lstO.ListIndex = 0
End Sub

Private Sub cmdSelNone_Click()
    Dim i As Integer
    
    For i = 0 To lstO.ListCount - 1
        lstO.Selected(i) = False
    Next
    
    lstO.ListIndex = 0
End Sub

Private Function TranslateText(sText As String, sLang As String) As String
    Dim sTmp As String, x As Integer, y As Integer
    
    On Local Error Resume Next
    
    sTmp = mUrlSource.DownloadUrlSource("http://babelfish.altavista.com/tr?urltext=" & sText & "&lp=" & sLang)
    
    x = InStr(sTmp, "<div style=padding:10px;>")
    y = InStr(x, sTmp, "</div>")
    
    sTmp = Mid(sTmp, x, y - x)
    sTmp = Mid(sTmp, InStr(sTmp, ">") + 1)
    
    sTmp = Replace(sTmp, Chr(10), " ")
    
    TranslateText = Trim(sTmp)
    
    Sleep 1000
End Function

Private Sub cmdTranslate_Click()
    If cmbLang.ListIndex = -1 Then Exit Sub
    
    Dim sTemp As String, i As Integer, sTrans As String, i2 As Integer
    
    cmbLang.Enabled = False
    lstO.Enabled = False
    cmdSelAll.Enabled = False
    cmdSelNone.Enabled = False
    cmdTranslate.Enabled = False
    
    progress.Max = lstO.ListCount
    
    For i = 0 To lstO.ListCount - 1
        DoEvents
        If lstO.Selected(i) Then
            If Trim(LPGStrings(lstO.ItemData(i)).String) <> "" Then
                sTemp = Trim(LPGStrings(lstO.ItemData(i)).String)
                sTemp = Replace(sTemp, "&", "")
                sTrans = TranslateText(sTemp, sALang(cmbLang.ListIndex + 1))
                LPGStrings(lstO.ItemData(i)).String = sTrans
            End If
        End If
        i2 = i2 + 1
        DoEvents
        progress.Value = i2
    Next
    
    MsgBox "Translation finished!", vbInformation
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim sLine As String, i As Integer

    Open App.Path & "\translator.lang" For Input As #1
        Do
            Input #1, sLine
            If Trim(sLine) <> "" And InStr(sLine, "|") > 0 Then
                iLangs = iLangs + 1
                ReDim Preserve sALang(iLangs)
                sALang(iLangs) = Left(sLine, InStr(sLine, "|") - 1)
                cmbLang.AddItem Mid(sLine, InStr(sLine, "|") + 1)
            End If
        Loop Until EOF(1)
    Close #1
    
    For i = 1 To iStrings
        If Trim(LPGStrings(i).String) <> "" Then
            lstO.AddItem LPGStrings(i).Name
            lstO.ItemData(lstO.NewIndex) = i
        End If
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmTransStr = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTransStr = Nothing
End Sub

