VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGridS 
   Caption         =   "Grid View"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7050
   Icon            =   "frmGridS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   431
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   470
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtGrid 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   5760
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid fGrid 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   9975
      _Version        =   393216
      RowHeightMin    =   300
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmGridS.frx":08E2
   End
End
Attribute VB_Name = "frmGridS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Recebe_Texto()
    With fGrid
        txtGrid.Top = (.CellTop / 15) + .Top + 3
        txtGrid.Left = (.CellLeft / 15) + .Left + 3
        
        txtGrid.Width = .CellWidth / 15 - 3
        txtGrid.Height = 16
        txtGrid.Text = fGrid.Text
        txtGrid.Visible = True
        txtGrid.SelStart = 0
        txtGrid.SelLength = Len(txtGrid)

        txtGrid.SetFocus
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim i As Integer
    
    For i = 1 To fGrid.Rows - 1
        fGrid.Row = i
        If fGrid.RowData(i) > 0 Then
            fGrid.Col = 1
            LPGStrings(fGrid.RowData(i)).String = fGrid.Text
        End If
    Next
End Sub

Private Sub Form_Load()
    Caption = "Grid View - Strings"
    
    With fGrid
        .ColWidth(0) = 1500
        .ColWidth(1) = 2500
        .Rows = 2
    End With
    
    Dim i As Integer, j As Integer
    
    j = 1
    fGrid.Rows = iStrings + 1
    
    fGrid.Row = 1
    For i = 1 To iStrings
        fGrid.RowData(j) = i
        fGrid.Col = 0
        fGrid.Text = LPGStrings(i).Name
        fGrid.Col = 1
        fGrid.Text = LPGStrings(i).String
        j = j + 1
        If j < fGrid.Rows Then fGrid.Row = j
    Next
End Sub

Private Sub Form_Resize()
    fGrid.Width = Me.ScaleWidth
    fGrid.Height = Me.ScaleHeight - 41
    cmdSave.Top = fGrid.Height + 10
    cmdClose.Top = fGrid.Height + 10
End Sub

Private Sub fGrid_Click()
    If fGrid.Rows = 1 Then Exit Sub
    Recebe_Texto
End Sub

Private Sub fGrid_KeyPress(KeyAscii As Integer)
    On Local Error GoTo ERRO
    With fGrid
        Select Case KeyAscii
            Case vbKeyReturn
                If .Col = .Cols - 1 Then
                    .Row = .Row + 1
                    .Col = 1
                Else
                    .Col = .Col + 1
                End If
            Case vbKeyBack
                If Trim(.Text) <> "" Then
                    .Text = Mid(.Text, 1, Len(.Text) - 1)
                End If
            Case Is < 32
            Case Else
                If .Col = 0 Or .Row = 0 Then
                    Exit Sub
                Else
                    .Text = .Text & Chr(KeyAscii)
                End If
        End Select
    End With
ERRO:
    
End Sub

Private Sub txtGrid_KeyPress(KeyAscii As Integer)
    On Local Error GoTo ERRO
    If KeyAscii = 13 Then
        fGrid.Text = txtGrid.Text
        txtGrid.Text = ""
        txtGrid.Visible = False
        If fGrid.Col = fGrid.Cols - 1 Then
            fGrid.Row = fGrid.Row + 1
            fGrid.Col = 0
        Else
            fGrid.Col = fGrid.Col + 1
        End If
    End If
ERRO:
    
End Sub

Private Sub txtGrid_LostFocus()
    txtGrid.Text = ""
    txtGrid.Visible = False
End Sub
