VERSION 5.00
Begin VB.Form fFind 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Find "
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4335
   ControlBox      =   0   'False
   Icon            =   "fFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.OptionButton opNext 
      Caption         =   "Next"
      Height          =   195
      Left            =   2340
      TabIndex        =   5
      Top             =   840
      Value           =   -1  'True
      Width           =   645
   End
   Begin VB.OptionButton opFirst 
      Caption         =   "First"
      Height          =   195
      Left            =   2340
      TabIndex        =   4
      Top             =   585
      Width           =   600
   End
   Begin VB.CheckBox ckWholeWord 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Whole Word Only"
      Height          =   195
      Left            =   435
      TabIndex        =   1
      Top             =   705
      Value           =   1  'Aktiviert
      Width           =   1590
   End
   Begin VB.TextBox txFind 
      Height          =   300
      Left            =   435
      MaxLength       =   40
      TabIndex        =   0
      Top             =   165
      Width           =   2775
   End
   Begin VB.CommandButton btCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   3375
      TabIndex        =   3
      Top             =   615
      Width           =   870
   End
   Begin VB.CommandButton btOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   3375
      TabIndex        =   2
      Top             =   150
      Width           =   870
   End
   Begin VB.Image img 
      Height          =   240
      Left            =   105
      Picture         =   "fFind.frx":000C
      Top             =   195
      Width           =   240
   End
End
Attribute VB_Name = "fFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btCancel_Click()

    Tag = ""
    Hide
    
End Sub

Private Sub btOK_Click()

    LastSrchFor = Trim$(txFind)
    Tag = LastSrchFor
    Hide
    
End Sub

Private Sub ckWholeWord_Click()

    On Error Resume Next
      txFind.SetFocus
    On Error GoTo 0

End Sub

Private Sub Form_Load()

    txFind = LastSrchFor
    ckWholeWord = WholeWord
    txFind.SelStart = 0
    txFind.SelLength = txFind.MaxLength

End Sub

Private Sub opFirst_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If opFirst Then
        LastFoundIndex = 0
    End If

End Sub

Private Sub txFind_Change()
    
    opFirst = True
    
End Sub

':) Ulli's VB Code Formatter V2.4.4 (21.10.2001 11:53:55) 1 + 48 = 49 Lines
