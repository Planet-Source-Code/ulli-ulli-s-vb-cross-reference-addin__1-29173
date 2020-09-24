VERSION 5.00
Begin VB.Form fTidy 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   585
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3120
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   39
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   208
   ShowInTaskbar   =   0   'False
   Begin VB.Image img 
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Left            =   60
      Picture         =   "fTidy.frx":0000
      Top             =   45
      Width           =   570
   End
   Begin VB.Label lbTdy 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Tidying up Virtual Memory Please wait..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Index           =   0
      Left            =   765
      TabIndex        =   0
      Top             =   105
      Width           =   2220
   End
   Begin VB.Label lbTdy 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   750
      TabIndex        =   1
      Top             =   90
      Width           =   2220
   End
End
Attribute VB_Name = "fTidy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    lbTdy(1) = lbTdy(0)

End Sub

':) Ulli's VB Code Formatter V2.4.4 (21.10.2001 11:53:54) 1 + 8 = 9 Lines
