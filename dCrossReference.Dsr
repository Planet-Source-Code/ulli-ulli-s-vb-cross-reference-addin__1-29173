VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} dCrossReference 
   ClientHeight    =   10005
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   11670
   _ExtentX        =   20585
   _ExtentY        =   17648
   _Version        =   393216
   Description     =   "VB Cross Reference"
   DisplayName     =   "Ulli's Cross Reference"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "dCrossReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'© 2000     UMGEDV GmbH  (umgedv@aol.com)
'
'Author     UMG (Ulli K. Muehlenweg)
'
'Title      VB6 Cross Reference Add-In
'
'Purpose    This Add-In creates a Cross Reference for your VB Projects. This Cross
'           Reference shows all Public, public, Friend and Private Data- and Code-
'           Member definitions, where they are defined, their scope, and whether they
'           are a Sub or Function, a Property, an Event, a Variable or a Constant.
'           Also shown are all references to these members with ComponentName.MemberName
'           as well as Line and Column numbers.
'           New: References to Controls will now be also be included.
'           Members with only one Reference (which can only be their definition) are
'           backcolored light green so you can check whether the member is indeed
'           unreferenced. Duplicate name definitions are backcolored light red.
'
'           You will need a bit of patience for large projects, the TreeView Cntl
'           and VB's Find Function are a little sluggish.
'
'           Compile the DLL into your VB directory and then use the Add-Ins
'           Manager to load the Cross Reference Add-In into VB.
'
'**********************************************************************************
'Development History
'**********************************************************************************
'24 Nov 2001 Version 1.4.3 - UMG
'
'Added Project References and Component Libraries
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'27 Sep 2001 Version 1.3.5 - UMG
'
'Added References to Controls
'    Unfortunately it is not easy to determine the scope of a Control, so all
'    references to a Control-Name will be listed including Variables with a
'    matching name. Controls will be ignored until the Project is saved.
'FontSizes in exported RTF
'Code Cosmetics
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'22 Sep 2001 Version 1.2.18 - UMG
'
'Finally fixed bugs with Resource Files and Related Documents
'Fixed bug with DupNames detection (Variable versus Constant)
'    Known quirk: If a definition uses continuation lines
'                 then DupNames detection may fail to detect
'Added progress bar to clipboard copy function
'Added Project Type and Component Icons
'Added Initial Directory for Save Function
'Added fTidy
'Clarified Component and Member Types
'Code Cosmetics
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'03 Mar 2001 Version 1.2.7 - UMG
'
'Fixed Bug (thanks to Darlene Fenn who reported it)
'    If the cross-referenced project contains a reference to a
'    Resource File then this Add-In used to crash with Error 91 (trying to
'    create a cross reference for the Resource File which is not possible).
'    Added check for both Searched_Compo present and Searched_For_Name present
'    Added CompoTypes vbext_ct_ResFile and vbext_ct_RelatedDocument
'
'Added Clipboard.Clear request on Closedown
'Changed Find Options First / Next
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'06 Jan 2001 Version 1.2 - UMG
'
'Added Color / Black & White Option to Export Function
'Added Component Type for Clipboard Copy Function
'Added EnsureVisible for Duplicated Items Nodes
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'30 Dec 2000 Version 1.1 - UMG
'
'Changed Export Function for sorted .RTF-File
'    Sorry, the leading zeros in the Treeview are needed to make it sort properly;
'    they will be suppressed in the rtf file and in the clipboard copy
'Added Find function for Treeview
'Changed Clipboard Copy Function for sorted Output
'Extended Export Function for more info in .RTF File
'Killed several bugs
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'01 Oct 2000 Version 1.0 Prototype - UMG
'
'Known Quirk: If you have a member with the same name as a method
'             (Example Function Add - Collection.Add) then these will both be
'             referenced. This is due to how VB has implemented Find Whole Word Only.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
DefLng A-Z 'we're 32 bit

Private Const MenuName          As String = "Add-Ins" 'you may need to localize "Add-Ins"
Private Const VBSettings        As String = "Software\Microsoft\VBA\Microsoft Visual Basic"
Private Const Fontface          As String = "Fontface"
Private Const Fontheight        As String = "Fontheight"
Private Const ResAs             As String = " As "
Private Const ResDim            As String = " Dim "
Private Const ResPrivate        As String = " Private "
Private Const ResConst          As String = " Const "
Private Const RootKey           As String = "X000001"
Private Const Backslash         As String = "\"
Private Const Quote             As String = """"

Private Const SleepTime         As Long = 999
Private Const LightGreen        As Long = &HE0FFE0
Private Const LightRed          As Long = &HE0E0FF
Private Const Gray              As Long = &H707070
Private Const DarkRed           As Long = &H80
Private Const DarkCyan          As Long = &H606000
Private Const DarkGreen         As Long = &H5000
Private Const PropChildIndex    As Long = 35

Private VBInstance              As VBIDE.VBE
Private CommandBarMenu          As CommandBar
Private MenuItem                As CommandBarControl
Private WithEvents MenuEvents   As CommandBarEvents
Attribute MenuEvents.VB_VarHelpID = -1
Private CompoA                  As VBComponent
Private CompoB                  As VBComponent
Private CurrentNode             As MSComctlLib.Node
Private ControlNames            As New Collection
Private ControlName             As Variant

Private VBFontSize
Private i, j, k, l
Private NumDeclLines
Private NumCodeLines
Private StartLine
Private StartColumn
Private EndColumn

Private KeyLevel1
Private KeyLevel2
Private KeyLevel3

Private NodeGroup               As String
Private NodeExt                 As String
Private CreatedNodeGroups       As String
Private Words()                 As String
Private CompoName               As String
Private CompoSaveName           As String
Private ProjectType             As String
Private KeyProjectImg           As String
Private KeyMemberImg            As String
Private CompoType               As String
Private MemberName              As String
Private ProcName                As String
Private CurrProcName            As String
Private VBFontName              As String

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)

  Dim ClipboardText As String

    Set VBInstance = Application
    If ConnectMode = ext_cm_External Then
        CheckProject
      Else 'NOT CONNECTMODE...
        On Error Resume Next
          Set CommandBarMenu = VBInstance.CommandBars(MenuName)
        On Error GoTo 0
        If CommandBarMenu Is Nothing Then
            MsgBox "Cross Referencer was loaded but could not be connected to the " & MenuName & " menu.", vbCritical
          Else 'NOT COMMANDBARMENU...
            fSplash.Show
            DoEvents
            With CommandBarMenu
                Set MenuItem = .Controls.Add(msoControlButton)
                i = .Controls.Count - 1
                If .Controls(i).BeginGroup And Not .Controls(i - 1).BeginGroup Then
                    'menu separator required
                    MenuItem.BeginGroup = True
                End If
            End With 'COMMANDBARMENU
            'set menu caption
            With App
                MenuItem.Caption = "&" & .ProductName & " V" & .Major & "." & .Minor & "." & .Revision & "..."
            End With 'APP
            With Clipboard
                ClipboardText = .GetText
                'set menu picture
                .SetData fSplash.picMenu.Image
                MenuItem.PasteFace
                .Clear
                .SetText ClipboardText
            End With 'CLIPBOARD
            'set event handler
            Set MenuEvents = VBInstance.Events.CommandBarEvents(MenuItem)
            'done connecting
            Sleep SleepTime
            Unload fSplash
        End If
    End If

End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)

    On Error Resume Next
      MenuItem.Delete
    On Error GoTo 0

End Sub

Private Sub MenuEvents_Click(ByVal CommandBarControl As Object, Handled As Boolean, CancelDefault As Boolean)

    CheckProject

End Sub

Public Sub CheckProject()

    Load fXref
    With fXref
        If VBInstance.ActiveVBProject Is Nothing Then
            MsgBox "Cannot see any Project - you must open  a Project first.", vbExclamation, .Caption
            Unload fXref
          Else 'NOT VBINSTANCE.ACTIVEVBPROJECT...
            SourceDir = VBInstance.ActiveVBProject.FileName
            i = InStrRev(SourceDir, Backslash)
            If i Then
                SourceDir = Left$(SourceDir, i)
            End If
            .lblDirty.Visible = VBInstance.ActiveVBProject.IsDirty
            If .ProjectName = VBInstance.ActiveVBProject.Name And Not VBInstance.ActiveVBProject.IsDirty Then
                .WindowState = LastWindowState
                BringWindowToTop .hWnd
              Else 'NOT .PROJECTNAME...
                If MsgBox("Do you want to create a Cross Reference for this Project?", vbQuestion + vbOKCancel, App.ProductName & " - " & VBInstance.ActiveVBProject.Name) = vbOK Then
                    .Enabled = False
                    .ProjectName = VBInstance.ActiveVBProject.Name
                    .WindowState = vbNormal
                    .ResetColorKey
                    BringWindowToTop .hWnd
                    Set VisibleNode = Nothing
                    DuplFound = False
                    .lblDupl = ""
                    .lblDupl.ToolTipText = ""
                    CreateXref
                    If Not VisibleNode Is Nothing Then
                        VisibleNode.EnsureVisible
                    End If
                    If DuplFound Then
                        .lblDupl = "Duplicate Names found"
                        .lblDupl.ToolTipText = "Name Shadowing is bad Coding Style"
                    End If
                    LastSrchFor = ""
                    LastFoundIndex = 0
                    WholeWord = vbChecked
                    .Enabled = True
                  Else 'NOT MSGBOX("DO YOU WANT TO CREATE A CROSS REFERENCE FOR THIS PROJECT?",...
                    Unload fXref
                End If
            End If
        End If
    End With 'FXREF

End Sub

Private Sub CreateXref()

    If RegOpenKeyEx(HKEY_CURRENT_USER, VBSettings, REG_OPTION_RESERVED, KEY_QUERY_VALUE, i) = ERROR_NONE Then
        'get VB editor font properties from registry
        l = Len(VBFontSize)
        If RegQueryValueEx(i, Fontheight, REG_OPTION_RESERVED, j, VBFontSize, l) <> ERROR_NONE Then
            VBFontSize = 9
        End If
        VBFontName = String$(128, 0)
        l = Len(VBFontName)
        If RegQueryValueEx(i, Fontface, REG_OPTION_RESERVED, j, ByVal VBFontName, l) = ERROR_NONE Then
            VBFontName = Left$(VBFontName, l - 1)
          Else 'NOT REGQUERYVALUEEX(I,...
            VBFontName = "Fixedsys"
        End If
        RegCloseKey i
      Else 'NOT REGOPENKEYEX(HKEY_CURRENT_USER,...
        VBFontName = "Fixedsys"
        VBFontSize = 9
    End If
        
    With fXref
        .imgUlli.Visible = False
        .pgb.Value = 0
        .pgb.Visible = True
        .lblLoading.Visible = True
        .Show
        .tvwRef.Font.Name = VBFontName
        .tvwRef.Font.Size = VBFontSize
        .rtfXRef.Font.Name = "Arial"
        .rtfXRef.Font.Size = 9
        DoEvents
        
        'here we go
        Screen.MousePointer = vbHourglass
        UnrefFound = False
        KeyLevel1 = 200000
        KeyLevel2 = 300000
        KeyLevel3 = 400000
        CreatedNodeGroups = ""
        With .tvwRef
            .Nodes.Clear
            j = 0
            Select Case VBInstance.ActiveVBProject.Type
              Case vbext_pt_StandardExe
                ProjectType = "Standard EXE"
                KeyProjectImg = "PrjStd"
              Case vbext_pt_ActiveXExe
                ProjectType = "ActiveX EXE"
                KeyProjectImg = "PrjAEx"
              Case vbext_pt_ActiveXControl
                ProjectType = "ActiveX OCX"
                KeyProjectImg = "PrjCtl"
              Case vbext_pt_ActiveXDll
                ProjectType = "ActiveX DLL"
                KeyProjectImg = "PrjDll"
            End Select
            Set CurrentNode = .Nodes.Add(, , RootKey, VBInstance.ActiveVBProject.Name & " (" & ProjectType & ")", KeyProjectImg)
            With CurrentNode
                .ForeColor = vbBlack
                .BackColor = vbWhite
                .Sorted = True
            End With 'CURRENTNODE
            CompoFound = False
            Open VBInstance.ActiveVBProject.FileName For Input As #1
            k = 0
            l = 0
            ProcName = ""
            CurrProcName = ""
            Do
                Input #1, MemberName 'Project File line, 'MemberName' is used for that
                If Len(Trim$(MemberName)) Then 'not an empty line
                    Words = Split(MemberName, "=")
                    Select Case LCase$(Words(0))
                      Case "reference"
                        Words = Split(Words(1), "#")
                        If k = 0 Then
                            Set CurrentNode = .Nodes.Add(RootKey, tvwChild, "X1r", "(References)", KeyFolderClosedImg)
                            With CurrentNode
                                .ForeColor = DarkRed
                                .BackColor = vbWhite
                            End With 'CURRENTNODE
                        End If
                        Inc k
                        Set CurrentNode = .Nodes.Add("X1r", tvwChild, "X2r" & Format$(k), Words(UBound(Words)), "Rfr")
                        With CurrentNode
                            .ForeColor = DarkCyan
                            .BackColor = vbWhite
                        End With 'CURRENTNODE
                      Case "object"
                        Words = Split(Words(1), " ")
                        If l = 0 Then
                            Set CurrentNode = .Nodes.Add(RootKey, tvwChild, "X1o", "(Libraries)", KeyFolderClosedImg)
                            With CurrentNode
                                .ForeColor = DarkRed
                                .BackColor = vbWhite
                            End With 'CURRENTNODE
                        End If
                        Inc l
                        Set CurrentNode = .Nodes.Add("X1o", tvwChild, "X2o" & Format$(l), Words(UBound(Words)), "Lbr")
                        With CurrentNode
                            .ForeColor = DarkGreen
                            .BackColor = vbWhite
                        End With 'CURRENTNODE
                      Case "versionfiledescription"
                        Words = Split(MemberName, "=")
                        ProcName = Words(UBound(Words))
                      Case "versionlegalcopyright"
                        Words = Split(MemberName, "=")
                        CurrProcName = Words(UBound(Words))
                    End Select
                End If
            Loop Until EOF(1)
            Close #1
            
            If Len(ProcName) + Len(CurrProcName) Then
                If Len(ProcName) Then
                    ProcName = ProcName & IIf(Len(CurrProcName), CurrProcName, "")
                    ProcName = Replace$(ProcName, """""", "; ")
                  Else 'LEN(PROCNAME) = 0
                    ProcName = CurrProcName
                End If
                Set CurrentNode = .Nodes.Add(RootKey, tvwChild, "X1d", ProcName, "Dsc")
                With CurrentNode
                    .ForeColor = vbBlack
                    .BackColor = vbWhite
                End With 'CURRENTNODE
            End If
            
            For Each CompoA In VBInstance.ActiveVBProject.VBComponents
                CompoFound = True
                CompoName = CompoA.Name
                CompoSaveName = CompoA.FileNames(1)
                If CompoName = "" Then
                    i = InStrRev(CompoSaveName, Backslash)
                    If i Then
                        CompoSaveName = Mid$(CompoSaveName, i + 1)
                    End If
                    CompoSaveName = "{" & CompoSaveName & "}"
                    NumDeclLines = -1
                  Else 'NOT COMPONAME...
                    With ControlNames
                        Do While .Count
                            .Remove 1 'remove all items from collection
                        Loop
                        If Len(CompoSaveName) Then
                            'read component file
                            i = FreeFile 'get a free file handle
                            Open CompoSaveName For Input As #i
                            Do
                                Input #i, MemberName 'source line, variable 'MemberName' is used for that
                                If Len(Trim$(MemberName)) Then 'not an empty line
                                    Words = Split(MemberName, " ")
                                    If LCase$(Words(0)) = "begin" Then 'word1 has the control type and word2 has the control name
                                        If UBound(Words) < 2 Then      'unless they are missing
                                            ReDim Preserve Words(2)
                                        End If
                                        On Error Resume Next 'skip duplicate control names
                                          .Add Words(2) & " (" & Words(1) & ")", Words(2)  'ctlname (ctltype), key to prevent duplicates
                                        On Error GoTo 0
                                    End If
                                End If
                            Loop Until LCase$(Words(0)) = "attribute" Or EOF(1)
                            Close #i
                            Erase Words
                            If .Count Then
                                .Remove 1 'don't want the container form, just the controls
                            End If
                        End If
                    End With 'CONTROLNAMES
                    CompoSaveName = CompoName 'name to add to .Nodes below
                    NumDeclLines = CompoA.CodeModule.CountOfDeclarationLines
                    NumCodeLines = CompoA.CodeModule.CountOfLines - NumDeclLines
                End If
                NodeExt = ""
                Select Case CompoA.Type
                  Case vbext_ct_ClassModule
                    CompoType = "Class"
                    NodeGroup = "a"
                  Case vbext_ct_ActiveXDesigner
                    CompoType = "Designer"
                    NodeGroup = "b"
                  Case vbext_ct_PropPage
                    CompoType = "Property Page"
                    NodeGroup = "c"
                  Case vbext_ct_UserControl
                    CompoType = "Control"
                    NodeGroup = "d"
                  Case vbext_ct_DocObject
                    CompoType = "User Document"
                    NodeGroup = "e"
                  Case vbext_ct_VBForm
                    If CompoA.Properties(PropChildIndex) Then
                        CompoType = "Child Form"
                        NodeGroup = "f"
                      Else 'COMPOA.PROPERTIES(PROPCHILDINDEX) = 0
                        CompoType = "Form"
                        NodeGroup = "g"
                    End If
                  Case vbext_ct_VBMDIForm
                    CompoType = "MDI Form"
                    NodeGroup = "h"
                  Case vbext_ct_StdModule
                    CompoType = "Module"
                    NodeGroup = "i"
                  Case vbext_ct_ResFile
                    CompoType = "Resource File"
                    NodeGroup = "j"
                    NodeExt = ")"
                  Case vbext_ct_RelatedDocument
                    CompoType = "Related Document"
                    NodeExt = "; File not found)"
                    On Error Resume Next
                      NodeExt = "; " & Format$(FileLen(CompoA.FileNames(1)), "0,0") & " Bytes)"
                    On Error GoTo 0
                    NodeGroup = "k"
                  Case Else
                    CompoType = "Unknown Component Type"
                    NodeExt = " " & CompoA.Type
                    NodeGroup = "l"
                    NumDeclLines = -1
                End Select
                If InStr(CreatedNodeGroups, NodeGroup) = 0 Then
                    CreatedNodeGroups = CreatedNodeGroups & NodeGroup
                    Set CurrentNode = .Nodes.Add(RootKey, tvwChild, "X1" & CompoType, CompoType & IIf(CompoType = "Class", "es", "s"), KeyFolderClosedImg)
                    With CurrentNode
                        .ForeColor = DarkRed
                        .BackColor = vbWhite
                        '.Sorted = True
                    End With 'CURRENTNODE
                End If
                Inc KeyLevel1
                Set CurrentNode = .Nodes.Add("X1" & CompoType, tvwChild, "X" & Format$(KeyLevel1), CompoSaveName & " (" & CompoType & IIf(Len(NodeExt), NodeExt, "; " & NumDeclLines & " + " & NumCodeLines & " = " & NumDeclLines + NumCodeLines & " Lines)"), Left$(CompoType, 3))
                With CurrentNode
                    .ForeColor = vbRed
                    .BackColor = vbWhite
                    .Sorted = True
                End With 'CURRENTNODE
                fXref.lblLoading = "Loading " & CompoName
                DoEvents
                If NumDeclLines >= 0 Then
                    'search all components for references to a Control
                    For Each ControlName In ControlNames
                        Inc KeyLevel2
                        MemberName = CStr(ControlName)
                        Set CurrentNode = fXref.tvwRef.Nodes.Add("X" & Format$(KeyLevel1), tvwChild, "X" & Format$(KeyLevel2), MemberName, "Tol")
                        With CurrentNode
                            .ForeColor = Gray
                            .BackColor = vbWhite
                            '.Sorted = True
                        End With 'CURRENTNODE
                        fXref.LED(1).Visible = Not fXref.LED(1).Visible
                        DoEvents
                        i = InStr(MemberName, " ")
                        MemberName = Left$(MemberName, i - 1) 'extract name of control
                        For Each CompoB In VBInstance.ActiveVBProject.VBComponents
                            FindReferences CompoB, CompoA, MemberName, 3
                        Next CompoB
                    Next ControlName
                    'search all components for references to current component
                    For Each CompoB In VBInstance.ActiveVBProject.VBComponents
                        FindReferences CompoB, CompoA, CompoName, 2
                    Next CompoB
                    If Not CompoA.CodeModule Is Nothing Then
                        'cycle thru all members of current component
                        For i = 1 To CompoA.CodeModule.Members.Count
                            With CompoA.CodeModule.Members.Item(i)
                                MemberName = .Name
                                Inc KeyLevel2
                                'determine member type
                                Select Case .Type
                                  Case vbext_mt_Method
                                    l = 1
                                    fXref.Colorkey(l) = "Sub or Function"
                                    KeyMemberImg = "Sub"
                                  Case vbext_mt_Property
                                    l = 2
                                    fXref.Colorkey(l) = "Property"
                                    KeyMemberImg = "Prp"
                                  Case vbext_mt_Event
                                    l = 3
                                    fXref.Colorkey(l) = "Event"
                                    KeyMemberImg = "Eve"
                                  Case vbext_mt_Variable
                                    l = 4
                                    fXref.Colorkey(l) = "Variable"
                                    KeyMemberImg = "Var"
                                  Case vbext_mt_Const
                                    l = 5
                                    fXref.Colorkey(l) = "Constant"
                                    KeyMemberImg = "Cst"
                                End Select
                                Set CurrentNode = fXref.tvwRef.Nodes.Add("X" & Format$(KeyLevel1), tvwChild, "X" & Format$(KeyLevel2), MemberName & " (" & Choose(.Scope, "Private", "Public", "Friend") & IIf(.Static And (.Type = vbext_mt_Method Or .Type = vbext_mt_Property), " Static", "") & ")", KeyMemberImg)
                                With CurrentNode
                                    .BackColor = LightGreen 'preliminary backcolor
                                    .ForeColor = QBColor(l)
                                    '.Sorted = True
                                    .Tag = l 'holds type
                                End With 'CURRENTNODE
                                fXref.LED(1).Visible = Not fXref.LED(1).Visible
                                DoEvents
                                'search current component for references to current member
                                FindReferences CompoA, CompoA, MemberName, 3
                                If .Scope <> vbext_Private Then
                                    'search the other components for references to current member
                                    For Each CompoB In VBInstance.ActiveVBProject.VBComponents
                                        If Not (CompoB Is CompoA) Then
                                            FindReferences CompoB, CompoA, MemberName, 3
                                        End If
                                    Next CompoB
                                End If
                            End With 'COMPOA.CODEMODULE.MEMBERS.ITEM(I)
                        Next i
                    End If
                End If
                Inc j, fXref.pgb.Max
                fXref.pgb = j / VBInstance.ActiveVBProject.VBComponents.Count
            Next CompoA
            fXref.lblLoading = "Sorting"
            For i = 1 To .Nodes.Count
                .Nodes(i).Sorted = True
            Next i
            .Nodes(1).Expanded = True
        End With '.TVWREF'FXREF
        .Refresh
        Sleep 333
        .pgb.Visible = False
        .lblLoading.Visible = False
        .imgUlli.Visible = True
        Screen.MousePointer = vbDefault
    End With 'FXREF
        
End Sub

Private Sub FindReferences(CompoSearch As VBComponent, CompoActive As VBComponent, Name As String, TreeLevel As Long)
    
    If CompoSearch.Name <> "" Then
        With CompoSearch.CodeModule
            If Not (CompoSearch.CodeModule Is Nothing Or Name = "") Then
                'Component to search in is present and Name to search for also
                StartLine = 1
                StartColumn = 1
                CurrProcName = "0"
                Do
                    EndColumn = -1
                    If .Find(Name, StartLine, StartColumn, -1, EndColumn, True) Then
                        If AcceptWord(.Lines(StartLine, 1), StartColumn, EndColumn) Then
                            On Error Resume Next
                              ProcName = .ProcOfLine(StartLine, vbext_pk_Get)  'this ProcKind parameter is just plain shit
                              ProcName = .ProcOfLine(StartLine, vbext_pk_Let)  'I know the line number and VB should know the ProcKind
                              ProcName = .ProcOfLine(StartLine, vbext_pk_Proc)
                              ProcName = .ProcOfLine(StartLine, vbext_pk_Set)  'lets hope this Ignore Err approach works
                            On Error GoTo 0
                            Select Case TreeLevel
                              Case 3
                                Inc KeyLevel3
                                Set CurrentNode = fXref.tvwRef.Nodes.Add("X" & Format$(KeyLevel2), tvwChild, "X" & Format$(KeyLevel3), LTrim$(IIf(CompoSearch Is CompoActive, "", CompoSearch.Name & IIf(Len(ProcName), ".", "")) & ProcName & " (L" & Format$(StartLine, "0000") & " C" & Format$(StartColumn, "0000") & ")"), KeyReferenceImg)
                                CurrentNode.BackColor = vbWhite
                                If ProcName = CurrProcName Then
                                    FoundDuplName
                                End If
                                If Not (CompoSearch Is CompoActive) Then
                                    If InStr(" " & .Lines(StartLine, 1), ResDim) Or InStr(" " & .Lines(StartLine, 1), ResPrivate) Or InStr(" " & .Lines(StartLine, 1), ResConst) Then
                                        If Mid$(.Lines(StartLine, 1), StartColumn - 4, 4) <> ResAs Then
                                            FoundDuplName
                                        End If
                                    End If
                                End If
                                If InStr(" " & .Lines(StartLine, 1), ResDim) Or InStr(" " & .Lines(StartLine, 1), ResConst) Then
                                    If Mid$(.Lines(StartLine, 1), StartColumn - 4, 4) <> ResAs Then
                                        If ProcName = "" Then
                                            CurrProcName = "0"
                                          Else 'NOT PROCNAME...
                                            FoundDuplName
                                            CurrProcName = ProcName
                                        End If
                                    End If
                                End If
                              Case 2
                                Inc KeyLevel2
                                Set CurrentNode = fXref.tvwRef.Nodes.Add("X" & Format$(KeyLevel1), tvwChild, "X" & Format$(KeyLevel2), LTrim$(IIf(CompoSearch Is CompoActive, "", CompoSearch.Name & IIf(Len(ProcName), ".", "")) & ProcName & " (L" & Format$(StartLine, "0000") & " C" & Format$(StartColumn, "0000") & ")"), KeyReferenceImg)
                                CurrentNode.BackColor = vbWhite
                            End Select
                            CurrentNode.ForeColor = QBColor(6)
                            fXref.Colorkey(6) = "Reference"
                            If CurrentNode.Parent.Children > 1 Then
                                CurrentNode.Parent.BackColor = vbWhite 'final backcolor
                              Else 'NOT CURRENTNODE.PARENT.CHILDREN...
                                UnrefFound = True
                            End If
                        End If
                        StartColumn = EndColumn
                      Else 'not found or no more found 'NOT .FIND(NAME,...
                        Exit Do '>---> Loop
                    End If
                Loop
            End If
        End With 'COMPOSEARCH.CODEMODULE
    End If

End Sub

Private Sub FoundDuplName()

    CurrentNode.BackColor = LightRed
    CurrentNode.Image = "Dup"
    Set VisibleNode = CurrentNode
    UnrefFound = True
    DuplFound = True

End Sub

Private Function AcceptWord(Line As String, StartColumn As Long, EndColumn As Long) As Boolean

    If StartColumn > 1 Then
        If Mid$(Line, StartColumn - 1, 1) = "_" Then
            Exit Function '>---> Bottom
        End If
    End If
    If EndColumn <= Len(Line) Then
        If Mid$(Line, EndColumn, 1) = "_" Then
            Exit Function '>---> Bottom
        End If
    End If
    For k = 1 To Len(Line)
        If Mid$(Line, k, 1) = Quote Then
            Do
                Inc k
                If k > Len(Line) Then
                    Exit Do '>---> Loop
                  Else 'NOT K...
                    If Mid$(Line, k, 1) = Quote Then
                        Exit Do '>---> Loop
                      Else 'NOT MID$(LINE,...
                        If k = StartColumn Then 'the word is in a literal
                            Exit Function '>---> Bottom
                        End If
                    End If
                End If
            Loop
        End If
    Next k
    k = InStr(" " & Line & " ", " Rem ") 'see if word is in a comment
    l = InStr(Line, "'")
    If k = 0 Then
        k = Len(Line)
    End If
    If l = 0 Then
        l = k
    End If
    If k > l Then
        k = l
    End If
    AcceptWord = (k >= StartColumn)
    
End Function

':) Ulli's VB Code Formatter V2.5.12 (24.11.2001 14:03:30) 154 + 568 = 722 Lines
