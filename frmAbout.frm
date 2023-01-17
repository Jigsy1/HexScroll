VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   3760
      Width           =   1335
   End
   Begin VB.Frame fmeLicense 
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   7455
      Begin VB.TextBox txtLicense 
         Height          =   3255
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   0
         Text            =   "frmAbout.frx":0000
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.Label lblLink 
      Caption         =   "https://github.com/Jigsy1/HexScroll/"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   1
      ToolTipText     =   "Opens in your browser"
      Top             =   3840
      Width           =   2775
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOkay_Click()
  Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
  Me.Caption = Me.Caption & " " & FrmHex.Tag
End Sub

Private Sub Form_Terminate()
  End
  ' `-> I doubt this has any use; but just incase...
End Sub

Private Sub lblLink_Click()
  CreateObject("WScript.Shell").Run lblLink.Caption
End Sub

' EOF
