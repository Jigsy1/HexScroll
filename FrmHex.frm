VERSION 5.00
Begin VB.Form FrmHex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hex Color Scroller"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrNoteClear 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   120
      Top             =   1200
   End
   Begin VB.CommandButton CmdCopy 
      Caption         =   "C&opy"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox TxtBlue 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3600
      TabIndex        =   5
      Text            =   "0"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox TxtGreen 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3600
      TabIndex        =   4
      Text            =   "0"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox TxtRed 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3600
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox TxtHex 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "#000000"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox TxtRGB 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0,0,0"
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin VB.HScrollBar HsBlue 
      Height          =   255
      Left            =   120
      Max             =   255
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.HScrollBar HsGreen 
      Height          =   255
      Left            =   120
      Max             =   255
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
   Begin VB.HScrollBar HsRed 
      Height          =   255
      Left            =   120
      Max             =   255
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Begin VB.Menu menuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu menuAbout 
      Caption         =   "&About"
      Begin VB.Menu menuLicense 
         Caption         =   "&License"
         Shortcut        =   ^L
      End
      Begin VB.Menu menuSep 
         Caption         =   "-"
      End
      Begin VB.Menu menuWebsite 
         Caption         =   "&Website..."
         Shortcut        =   ^W
      End
   End
End
Attribute VB_Name = "FrmHex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function updateValues(redValue, greenValue, blueValue)
  Me.BackColor = RGB(redValue, greenValue, blueValue)
  If TxtRed.Text <> redValue Then TxtRed.Text = redValue
  If TxtGreen.Text <> greenValue Then TxtGreen.Text = greenValue
  If TxtBlue.Text <> blueValue Then TxtBlue.Text = blueValue
  TxtHex.Text = "#" & IIf(Len(Hex(redValue)) <= 1, "0" & Hex(redValue), Hex(redValue)) & IIf(Len(Hex(greenValue)) <= 1, "0" & Hex(greenValue), Hex(greenValue)) & IIf(Len(Hex(blueValue)) <= 1, "0" & Hex(blueValue), Hex(blueValue))
  TxtRGB.Text = redValue & "," & greenValue & "," & blueValue
End Function

Private Sub CmdClose_Click()
  End
End Sub

Private Sub CmdCopy_Click()
  On Error GoTo endCopy
    Clipboard.Clear
    ' `-> I doubt this has any use; but just incase...
    Clipboard.SetText TxtHex.Text
    Me.Caption = Me.Caption & " - Copied to clipboard"
    tmrNoteClear.Enabled = True
  Exit Sub

endCopy:
  MsgBox "Failed to copy to clipboard.", vbExclamation, "Error"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then
    Randomize
    Call updateValues(Int(Val(Rnd * 256)), Int(Val(Rnd * 256)), Int(Val(Rnd * 256)))
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then End
End Sub

Private Sub Form_Load()
  Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
  Me.Tag = Me.Caption
  ' `-> Store the title in "memory" for easy changing if someone clicks copy.
  Call updateValues(HsRed.Value, HsGreen.Value, HsBlue.Value)
End Sub

Private Sub HsBlue_Change()
  Call updateValues(HsRed.Value, HsGreen.Value, HsBlue.Value)
End Sub

Private Sub HsGreen_Change()
  Call updateValues(HsRed.Value, HsGreen.Value, HsBlue.Value)
End Sub

Private Sub HsRed_Change()
  Call updateValues(HsRed.Value, HsGreen.Value, HsBlue.Value)
End Sub

Private Sub Form_Terminate()
  End
  ' `-> I doubt this has any use; but just incase...
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub menuExit_Click()
  End
End Sub

Private Sub menuLicense_Click()
  frmAbout.Visible = True
End Sub

Private Sub menuWebsite_Click()
  CreateObject("WScript.Shell").Run frmAbout.lblLink.Caption
End Sub

Private Sub tmrNoteClear_Timer()
  Me.Caption = Me.Tag
  tmrNoteClear.Enabled = False
End Sub

Private Sub TxtBlue_Change()
  If IsNull(TxtBlue.Text) = False And TxtBlue.Text <> "" And IsNumeric(TxtBlue.Text) = True Then
    If TxtBlue.Text >= 0 And TxtBlue.Text <= 255 Then
      HsBlue.Value = TxtBlue.Text
      Call updateValues(HsRed.Value, HsGreen.Value, HsBlue.Value)
    End If
  End If
End Sub

Private Sub TxtGreen_Change()
  If IsNull(TxtGreen.Text) = False And TxtGreen.Text <> "" And IsNumeric(TxtGreen.Text) = True Then
    If TxtGreen.Text >= 0 And TxtGreen.Text <= 255 Then
      HsGreen.Value = TxtGreen.Text
      Call updateValues(HsRed.Value, HsGreen.Value, HsBlue.Value)
    End If
  End If
End Sub

Private Sub TxtRed_Change()
  If IsNull(TxtRed.Text) = False And TxtRed.Text <> "" And IsNumeric(TxtRed.Text) = True Then
    If TxtRed.Text >= 0 And TxtRed.Text <= 255 Then
      HsRed.Value = TxtRed.Text
      Call updateValues(HsRed.Value, HsGreen.Value, HsBlue.Value)
    End If
  End If
End Sub

' EOF
