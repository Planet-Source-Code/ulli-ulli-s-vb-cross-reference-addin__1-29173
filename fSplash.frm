VERSION 5.00
Begin VB.Form fSplash 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4620
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   Icon            =   "fSplash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picMenu 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      Height          =   240
      Left            =   2370
      Picture         =   "fSplash.frx":000C
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   2
      Top             =   105
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.OptionButton opDummy 
      Height          =   195
      Left            =   4710
      TabIndex        =   1
      Top             =   825
      Width           =   180
   End
   Begin VB.Image img 
      BorderStyle     =   1  'Fest Einfach
      Height          =   765
      Left            =   195
      Picture         =   "fSplash.frx":034E
      Top             =   188
      Width           =   825
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading X-Ref-AddIn..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1230
      TabIndex        =   0
      Top             =   450
      Width           =   2340
   End
End
Attribute VB_Name = "fSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

Private Sub Form_Load()

    LastWindowState = vbNormal
    WholeWord = vbChecked
    
End Sub

':) Ulli's VB Code Formatter V2.4.4 (21.10.2001 11:53:58) 2 + 9 = 11 Lines
