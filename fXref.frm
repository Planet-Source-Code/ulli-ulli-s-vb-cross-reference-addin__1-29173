VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fXref 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   6795
   ClientLeft      =   2190
   ClientTop       =   2235
   ClientWidth     =   5430
   Icon            =   "fXref.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   5430
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComctlLib.ImageList imgList 
      Left            =   2100
      Top             =   3270
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   29
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":030A
            Key             =   "Cla"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":041E
            Key             =   "FilC"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":0532
            Key             =   "FilO"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":0646
            Key             =   "Con"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":075A
            Key             =   "Des"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":086E
            Key             =   "For"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":0982
            Key             =   "Chi"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":0AA2
            Key             =   "MDI"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":0BB6
            Key             =   "Mod"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":0CCA
            Key             =   "PrjAEx"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":0DDE
            Key             =   "PrjCtl"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":0EF2
            Key             =   "PrjDll"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":1006
            Key             =   "PrjStd"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":1126
            Key             =   "Pro"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":1246
            Key             =   "Res"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":1362
            Key             =   "Rel"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":1476
            Key             =   "Dsc"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":158A
            Key             =   "Lbr"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":169E
            Key             =   "Rfr"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":17B2
            Key             =   "Unk"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":18CE
            Key             =   "Use"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":19EA
            Key             =   "Ref"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":1AFE
            Key             =   "Dup"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":1C1A
            Key             =   "Tol"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":1D2E
            Key             =   "Var"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":1E42
            Key             =   "Cst"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":1F56
            Key             =   "Eve"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":206A
            Key             =   "Prp"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fXref.frx":217E
            Key             =   "Sub"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   4155
      Top             =   405
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "RTF"
      DialogTitle     =   "Save Cross Reference"
      FileName        =   "XREF.RTF"
   End
   Begin RichTextLib.RichTextBox rtfXRef 
      Height          =   270
      Left            =   4155
      TabIndex        =   6
      Top             =   900
      Visible         =   0   'False
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   476
      _Version        =   393217
      TextRTF         =   $"fXref.frx":2292
   End
   Begin MSComctlLib.ProgressBar pgb 
      Height          =   240
      Left            =   90
      TabIndex        =   3
      Top             =   765
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.PictureBox picColorkey 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   2430
      ScaleHeight     =   945
      ScaleWidth      =   2850
      TabIndex        =   2
      ToolTipText     =   "Color Key"
      Top             =   300
      Width           =   2910
   End
   Begin MSComctlLib.TreeView tvwRef 
      Height          =   5235
      Left            =   90
      TabIndex        =   0
      Top             =   1455
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   9234
      _Version        =   393217
      Indentation     =   176
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "imgList"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgUlli 
      BorderStyle     =   1  'Fest Einfach
      Height          =   765
      Left            =   90
      Picture         =   "fXref.frx":2367
      Top             =   300
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Image LED 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   195
      Index           =   1
      Left            =   105
      Picture         =   "fXref.frx":404D
      Top             =   435
      Width           =   195
   End
   Begin VB.Label lblDupl 
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
      ForeColor       =   &H000000C0&
      Height          =   420
      Left            =   1155
      TabIndex        =   7
      Top             =   540
      Width           =   1140
   End
   Begin VB.Label lblDirty 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Save your Project"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   90
      TabIndex        =   5
      ToolTipText     =   "You have changed something..."
      Top             =   1125
      Width           =   1530
   End
   Begin VB.Label lblLoading 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   345
      TabIndex        =   4
      Top             =   435
      Width           =   2010
   End
   Begin VB.Label lbl 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Color Key"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2445
      TabIndex        =   1
      Top             =   75
      Width           =   675
   End
   Begin VB.Image LED 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   195
      Index           =   0
      Left            =   105
      Picture         =   "fXref.frx":4513
      Top             =   435
      Width           =   195
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSave 
         Caption         =   "&Export"
         Begin VB.Menu mnuColor 
            Caption         =   "in &Color"
         End
         Begin VB.Menu mnuBW 
            Caption         =   "in &Black && White"
         End
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuxRef 
      Caption         =   "x&Ref"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy to Clipboard"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExpand 
         Caption         =   "E&xpand all"
      End
      Begin VB.Menu mnuCollapse 
         Caption         =   "C&ollapse all"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&?"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMail 
         Caption         =   "&Send Mail to Author..."
      End
   End
End
Attribute VB_Name = "fXref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z 'we're 32 bit
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Private Const CRL = "Cross Reference Listing "

Private RightMargin
Private BottomMargin
Private i, j
Private OutCount
Private MsgType

Private myColorKey(1 To 6)  As String
Private myProjectName       As String
Private InsForBMP           As String 'six spaces icon width if with icons
Private Text                As String
Private MemberType          As String
Private ClipCopy            As Boolean

Private IsColored           As OLE_COLOR

Private Node1               As MSComctlLib.Node
Private Node2               As MSComctlLib.Node
Private Node3               As MSComctlLib.Node
Private Node4               As MSComctlLib.Node
Private Node5               As MSComctlLib.Node

Private Sub Form_Load()
     
    Caption = App.ProductName
    RightMargin = ScaleWidth - tvwRef.Width
    BottomMargin = ScaleHeight - tvwRef.Height
    IsColored = vbWhite
    ClipCopy = False
    imgUlli.ToolTipText = App.LegalCopyright
    
End Sub

Private Sub Form_Paint()
    
    With picColorkey
        .Cls
        .ForeColor = vbBlack
        If CompoFound Then
            picColorkey.Print " Project"; Tab(21);
            .ForeColor = vbRed
            picColorkey.Print " Component"
            For i = 1 To 6 Step 2
                .ForeColor = QBColor(i)
                picColorkey.Print " "; myColorKey(i); Tab(21);
                .ForeColor = QBColor(i + 1)
                picColorkey.Print " "; myColorKey(i + 1)
            Next i
          Else 'COMPOFOUND = 0
            picColorkey.Print " Project has no Components"
        End If
    End With 'PICCOLORKEY
    
End Sub

Public Property Let Colorkey(Index As Long, nuColorKey As String)

    myColorKey(Index) = nuColorKey
    
End Property

Private Sub Form_Resize()

    On Error Resume Next
      tvwRef.Width = ScaleWidth - RightMargin
      tvwRef.Height = ScaleHeight - BottomMargin
      picColorkey.Left = tvwRef.Left + tvwRef.Width - picColorkey.Width
      lbl.Left = picColorkey.Left
      pgb.Width = lbl.Left - RightMargin
      If WindowState <> vbMinimized Then
          LastWindowState = WindowState
      End If
    On Error GoTo 0
    
End Sub

Private Sub Form_Terminate()

    Screen.MousePointer = vbDefault
    Do
    Loop Until ShowCursor(True) >= 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If ClipCopy Then
        Select Case MsgBox("You have copied the Cross Reference Listing to the Clipboard. Do you want to keep it?", vbQuestion Or vbDefaultButton2 Or vbYesNoCancel, Caption)
          Case vbNo
            Clipboard.Clear
            ClipCopy = False
          Case vbCancel
            Cancel = True
        End Select
    End If
    If Not Cancel Then 'about to exit - tidy up
        Visible = False
        DoEvents
        myProjectName = ""
        LastWindowState = vbNormal
    End If

End Sub

Private Sub imgUlli_Click()

    mnuAbout_Click

End Sub

Private Sub mnuAbout_Click()
    
    With App
        ShellAbout Me.hWnd, "About " & .ProductName & "#Operating System:", AppDetails & vbCrLf & .LegalCopyright, Me.Icon.Handle
    End With 'APP
    
End Sub

Private Sub mnuClose_Click()

    Unload Me
    
End Sub

Private Sub mnuCollapse_Click()
    
    tvwRef.Visible = False
    For i = 1 To tvwRef.Nodes.Count
        tvwRef.Nodes(i).Expanded = False
    Next i
    tvwRef.Nodes(1).Expanded = True
    tvwRef.Visible = True
    tvwRef.Nodes(1).EnsureVisible

End Sub

Private Sub mnuCopy_Click()

  'To Clipboard
  
  Dim NC As String
  
    NC = vbCrLf & "'"
    Do
    Loop Until ShowCursor(False) < 0
    Clipboard.Clear
    imgUlli.Visible = False
    pgb.Value = 0
    OutCount = 0
    lblDupl.Visible = False
    pgb.Visible = True
    Enabled = False
    lblLoading = "Copying"
    lblLoading.Visible = True
    DoEvents
    Set Node1 = tvwRef.Nodes(1)
    Text = "'" & CRL & "for " & String$(Val(Mid$(Node1.Key, 2, 1)), vbTab) & Node1.Text & NC
    If Node1.Children Then
        Set Node2 = Node1.Child.FirstSibling
        Do Until Node2 Is Nothing
            Text = Text & NC & String$(Val(Mid$(Node2.Key, 2, 1)), vbTab) & Node2.Text & NC
            Inc OutCount, 100
            pgb.Value = OutCount / tvwRef.Nodes.Count
            If Node2.Children Then
                Set Node3 = Node2.Child.FirstSibling
                Do Until Node3 Is Nothing
                    If Len(Node3.Tag) Then
                        MemberType = " " & myColorKey(Val(Node3.Tag)) & ")"
                      Else 'LEN(NODE3.TAG) = 0
                        MemberType = ")"
                    End If
                    Text = Text & NC & String$(Val(Mid$(Node3.Key, 2, 1)), vbTab) & Node3.Text & NC
                    Inc OutCount, 100
                    pgb.Value = OutCount / tvwRef.Nodes.Count
                    If Node3.Children Then
                        Set Node4 = Node3.Child.FirstSibling
                        Do Until Node4 Is Nothing
                            Text = Text & String$(Val(Mid$(Node4.Key, 2, 1)), vbTab) & ZeroSuppress(Node4.Text) & NC
                            Inc OutCount, 100
                            pgb.Value = OutCount / tvwRef.Nodes.Count
                            If Node4.Children Then
                                LED(1).Visible = Not LED(1).Visible
                                DoEvents
                                Set Node5 = Node4.Child.FirstSibling
                                Do Until Node5 Is Nothing
                                    Text = Text & String$(Val(Mid$(Node5.Key, 2, 1)), vbTab) & ZeroSuppress(Node5.Text) & NC
                                    Inc OutCount, 100
                                    pgb.Value = OutCount / tvwRef.Nodes.Count
                                    Set Node5 = Node5.Next
                                Loop
                            End If
                            Set Node4 = Node4.Next
                        Loop
                    End If
                    Set Node3 = Node3.Next
                Loop
            End If
            Set Node2 = Node2.Next
        Loop
    End If
    Text = Text & NC & "End of " & CRL & "for " & Left$(Node1.Text, InStr(Node1.Text, " ")) & _
           NC & "Created by " & AppDetails & NC & App.LegalCopyright
    Clipboard.SetText Text
    ClipCopy = True
    lblLoading.Visible = False
    imgUlli.Visible = True
    pgb.Visible = False
    lblDupl.Visible = True
    Enabled = True
    Do
    Loop Until ShowCursor(True) >= 0
    MsgBox CRL & "transferred to Clipboard.", vbInformation, Caption
    
End Sub

Private Sub mnuExpand_Click()
    
    Screen.MousePointer = vbHourglass
    On Error Resume Next
      tvwRef.Visible = False
      For i = 1 To tvwRef.Nodes.Count
          tvwRef.Nodes(i).Expanded = True
      Next i
      tvwRef.Visible = True
      tvwRef.Nodes(1).EnsureVisible
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub mnuFind_Click()

  Dim NodeText As String
    
    With fFind
        .Move Left + 600, Top + 600
        .opNext = LastFoundIndex
        .opFirst = Not .opNext
        .Show vbModal
        Text = UCase$(.Tag)
        WholeWord = .ckWholeWord
    End With 'FFIND
    Unload fFind
    Screen.MousePointer = vbHourglass
    DoEvents
    If Len(Text) Then
        mnuCollapse_Click
        With tvwRef
            Enabled = False
            For i = LastFoundIndex + 1 To .Nodes.Count
                NodeText = UCase$(.Nodes(i).Text)
                j = InStr(NodeText, " ")
                If j Then
                    NodeText = Left$(NodeText, j - 1)
                End If
                If WholeWord = vbChecked Then
                    If NodeText = Text Then
                        Exit For '>---> Next
                      ElseIf InStr(NodeText, ".") Then 'NOT NODETEXT...
                        If Left$(NodeText, InStr(NodeText, ".") - 1) = Text Then
                            Exit For '>---> Next
                        End If
                    End If
                  Else 'NOT WHOLEWORD...
                    If InStr(NodeText, Text) Then
                        Exit For '>---> Next
                    End If
                End If
                .Nodes(i).Selected = False
            Next i
            If i > .Nodes.Count Then 'found nothing
                LastFoundIndex = 0
              Else 'found something 'NOT I...
                With .Nodes(i)
                    .EnsureVisible
                    .Expanded = True
                    .Selected = True
                End With '.NODES(I)
                LastFoundIndex = i
            End If
            Enabled = True
        End With 'TVWREF
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub mnuMail_Click()

    With App
        SendMeMail hWnd, AppDetails
    End With 'APP
   
End Sub

Private Sub mnuBW_Click()

  'RTF in Black and White

    IsColored = vbBlack
    mnuColor_Click
    IsColored = vbWhite

End Sub

Private Sub mnuColor_Click()
  
  'RTF in Color (also used for B&W - see IsColored)

    With cDlg
        .FileName = "XRef for " & myProjectName & ".rtf"
        .Flags = cdlOFNLongNames Or cdlOFNOverwritePrompt
        .Filter = "Rich Text Format  (*.rtf)|*.rtf"
        .FilterIndex = 1
        .InitDir = SourceDir
        On Error Resume Next
          .ShowSave
          i = Err
        On Error GoTo 0
    End With 'CDLG
    If i = 0 Then
        Select Case MsgBox("Estimated File Size with Icons is " & Format$(tvwRef.Nodes.Count * 1.905, "#,0") & " kB" & vbCrLf & vbCrLf & "Include Icons in .RTF-File?", vbQuestion Or vbYesNoCancel Or vbDefaultButton2, Caption)
          Case vbYes
            MsgType = WM_PASTE
            InsForBMP = "      "
          Case vbNo
            MsgType = 0
            InsForBMP = ""
          Case Else
            MsgType = -1
        End Select
        If MsgType >= 0 Then
            imgUlli.Visible = False
            pgb.Value = 0
            OutCount = 0
            lblDupl.Visible = False
            pgb.Visible = True
            Enabled = False
            lblLoading = "Creating RTF"
            lblLoading.Visible = True
            DoEvents
            Do
            Loop Until ShowCursor(False) < 0
            With rtfXRef
                Set Node1 = tvwRef.Nodes(1)
                .Text = ""
                InsertIcon .hWnd, Node1
                .SelBold = True
                .SelFontSize = 14
                .SelColor = Node1.ForeColor And IsColored
                .SelText = CRL & "for " & String$(Val(Mid$(Node1.Key, 2, 1)), vbTab) & Node1.Text & vbCrLf
                .SelFontSize = 8
                .SelBold = False
                .SelText = InsForBMP & "Created by " & AppDetails & " on " & Format$(Now, "long date") & " at " & Format$(Now, "long time") & vbCrLf & InsForBMP & App.LegalCopyright & vbCrLf
                If CompoFound And IsColored = vbWhite Then
                    .SelBold = True
                    .SelColor = vbBlack
                    .SelFontSize = 9
                    .SelText = vbCrLf & "Color Key"
                    .SelBold = False
                    .SelColor = vbRed
                    .SelText = vbTab & "Component" & vbCrLf
                    For i = 1 To 6
                        If myColorKey(i) <> "" Then
                            .SelColor = QBColor(i)
                            .SelText = vbTab & vbTab & myColorKey(i) & vbCrLf
                        End If
                    Next i
                  Else 'NOT COMPOFOUND...
                    .SelText = vbCrLf
                End If
                If UnrefFound Then
                    .SelColor = &H808080 And IsColored
                    .SelItalic = True
                    .SelText = vbTab & vbTab & "Italic = Unreferenced or Duplicated" & vbCrLf
                    .SelItalic = False
                End If
                If Node1.Children Then
                    Set Node2 = Node1.Child.FirstSibling
                    Do Until Node2 Is Nothing
                        .SelText = vbCrLf & String$(Val(Mid$(Node2.Key, 2, 1)), vbTab)
                        InsertIcon .hWnd, Node2
                        .SelBold = True
                        .SelFontSize = 12
                        .SelColor = Node2.ForeColor And IsColored
                        .SelText = Node2.Text & vbCrLf
                        Inc OutCount, 100
                        pgb.Value = OutCount / tvwRef.Nodes.Count
                        .SelBold = False
                        If Node2.Children Then
                            Set Node3 = Node2.Child.FirstSibling
                            Do Until Node3 Is Nothing
                                .SelText = vbCrLf & String$(Val(Mid$(Node3.Key, 2, 1)), vbTab)
                                InsertIcon .hWnd, Node3
                                .SelBold = True
                                .SelFontSize = 10
                                .SelColor = Node3.ForeColor And IsColored
                                .SelText = Node3.Text & vbCrLf
                                Inc OutCount, 100
                                pgb.Value = OutCount / tvwRef.Nodes.Count
                                If Node3.Children Then
                                    Set Node4 = Node3.Child.FirstSibling
                                    Do Until Node4 Is Nothing
                                        If Len(Node4.Tag) Then
                                            MemberType = " " & myColorKey(Val(Node4.Tag)) & ")"
                                          Else 'NOT LEN(NODE4.TAG)... 'LEN(NODE4.TAG) = 0
                                            MemberType = ")"
                                        End If
                                        .SelText = String$(Val(Mid$(Node4.Key, 2, 1)), vbTab)
                                        InsertIcon .hWnd, Node4
                                        .SelFontSize = 9
                                        .SelColor = Node4.ForeColor And IsColored
                                        .SelItalic = (Node4.BackColor <> vbWhite)
                                        .SelBold = (Node4.Image <> KeyReferenceImg)
                                        .SelText = Replace$(ZeroSuppress(Node4.Text), ")", MemberType) & vbCrLf
                                        Inc OutCount, 100
                                        pgb.Value = OutCount / tvwRef.Nodes.Count
                                        LED(1).Visible = Not LED(1).Visible
                                        DoEvents
                                        If Node4.Children Then
                                            Set Node5 = Node4.Child.FirstSibling
                                            Do Until Node5 Is Nothing
                                                .SelText = String$(Val(Mid$(Node5.Key, 2, 1)), vbTab)
                                                InsertIcon .hWnd, Node5
                                                .SelBold = False
                                                .SelColor = Node5.ForeColor And IsColored
                                                .SelItalic = (Node5.BackColor <> vbWhite)
                                                .SelFontSize = 9
                                                .SelText = ZeroSuppress(Node5.Text) & vbCrLf
                                                Inc OutCount, 100
                                                pgb.Value = OutCount / tvwRef.Nodes.Count
                                                Set Node5 = Node5.Next
                                            Loop
                                        End If
                                        Set Node4 = Node4.Next
                                    Loop
                                End If
                                Set Node3 = Node3.Next
                            Loop
                        End If
                        Set Node2 = Node2.Next
                    Loop
                End If
                Text = Node1.Text
                Set Node1 = Nothing
                .SelBold = True
                .SelColor = vbBlack
                .SelText = vbCrLf & "End of " & CRL & "for " & Left$(Text, InStr(Text, " "))
                LED(1).Visible = True
                lblLoading = "Writing"
                DoEvents
                .SaveFile cDlg.FileName
                lblLoading.Visible = False
                imgUlli.Visible = True
                pgb.Visible = False
                lblDupl.Visible = True
                Enabled = True
                Do
                Loop Until ShowCursor(True) >= 0
                i = FileLen(cDlg.FileName)
                Text = Format$(i, "#,0") & " Bytes"
                MsgBox "File " & UCase$(cDlg.FileTitle) & " saved; " & IIf(i >= 1024, Format$(i / 1024, "#,0.0") & " kB (" & Text & ")", Text) & IIf(MsgType > 0 And ClipCopy, vbCrLf & vbCrLf & "The Clipboard was cleared.", ""), vbInformation, Caption
                If MsgType > 0 Then
                    ClipCopy = False
                    Clipboard.Clear
                End If
                Enabled = False
                Screen.MousePointer = vbHourglass
                On Error Resume Next
                  fTidy.Move Left + (Width - fTidy.Width) / 2, Top + (Height - fTidy.Height) / 2
                On Error GoTo 0
                fTidy.Show
                DoEvents
                .Text = ""
                Unload fTidy
                Enabled = True
                tvwRef.SetFocus
                Screen.MousePointer = vbDefault
            End With 'RTFXREF
        End If
    End If

End Sub

Private Sub InsertIcon(hWnd As Long, Node As MSComctlLib.Node)

    If MsgType Then 'a legal message type
        Clipboard.Clear
        If Node.Image = KeyFolderClosedImg Then
            Clipboard.SetData imgList.ListImages(KeyFolderOpenImg).Picture
          Else 'NOT NODE.IMAGE...
            Clipboard.SetData imgList.ListImages(Node.Image).Picture
        End If
        SendMessage hWnd, MsgType, 0&, 0&
    End If

End Sub

Private Function ZeroSuppress(Txt As String) As String

    ZeroSuppress = Replace$(Replace$(Replace$(Replace$(Replace$(Replace$(Txt, "(L0", "(L"), "(L0", "(L"), "(L0", "(L"), " C0", " C"), " C0", " C"), " C0", " C")

End Function

Public Property Get ProjectName() As String
    
    ProjectName = myProjectName
    
End Property

Public Property Let ProjectName(nuProjectName As String)
    
    myProjectName = nuProjectName
    
End Property

Public Sub ResetColorKey()

    Erase myColorKey

End Sub

Private Sub tvwRef_Collapse(ByVal Node As MSComctlLib.Node)

    If Node.Image = KeyFolderOpenImg Then
        Node.Image = KeyFolderClosedImg
    End If

End Sub

Private Sub tvwRef_DblClick()

    With tvwRef.SelectedItem
        If .Image = "Rel" Then
            Text = Mid$(.Text, 2)
            Text = Left$(Text, InStr(Text, "}") - 1)
            If ShellExecute(hWnd, "open", Text, vbNullString, SourceDir, SW_SHOWNORMAL) < SE_NO_ERROR Then
                MsgBox "Cannot open " & SourceDir & Text, vbCritical, "File not found"
            End If
        End If
    End With 'TVWREF.SELECTEDITEM

End Sub

Private Sub tvwRef_Expand(ByVal Node As MSComctlLib.Node)

    If Node.Image = KeyFolderClosedImg Then
        Node.Image = KeyFolderOpenImg
    End If

End Sub

':) Ulli's VB Code Formatter V2.4.4 (21.10.2001 14:35:06) 26 + 532 = 558 Lines
