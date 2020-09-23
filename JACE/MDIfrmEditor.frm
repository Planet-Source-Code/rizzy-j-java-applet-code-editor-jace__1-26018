VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm mainEditor 
   BackColor       =   &H00808080&
   Caption         =   "JACE (Java Applet Code Editor)"
   ClientHeight    =   8130
   ClientLeft      =   180
   ClientTop       =   1995
   ClientWidth     =   10935
   Icon            =   "MDIfrmEditor.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog docDialog 
      Left            =   9600
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog DisAsmDialog 
      Left            =   9000
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog appletDialog 
      Left            =   10200
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10200
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrmEditor.frx":0ABA
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrmEditor.frx":0C82
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrmEditor.frx":0D82
            Key             =   "save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrmEditor.frx":0E6E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrmEditor.frx":0FB2
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrmEditor.frx":125A
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrmEditor.frx":1396
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrmEditor.frx":14DE
            Key             =   "viewbrowser"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrmEditor.frx":16BA
            Key             =   "find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrmEditor.frx":1896
            Key             =   "findagain"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ne"
            Object.ToolTipText     =   "New"
            ImageKey        =   "new"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "op"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sa"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pr"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "print"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cu"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "cut"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "co"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "copy"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pa"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "paste"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "we"
            Object.ToolTipText     =   "View In Browser"
            ImageKey        =   "viewbrowser"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "fi"
            Object.ToolTipText     =   "Find"
            ImageKey        =   "find"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "fa"
            Object.ToolTipText     =   "Find Again"
            ImageKey        =   "findagain"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog saveJava 
      Left            =   9600
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog openJava 
      Left            =   9000
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu filemenu 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu open 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu close 
         Caption         =   "&Close"
      End
      Begin VB.Menu gap1 
         Caption         =   "-"
      End
      Begin VB.Menu save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu saveas 
         Caption         =   "Save &As..."
         Shortcut        =   {F12}
      End
      Begin VB.Menu gap2 
         Caption         =   "-"
      End
      Begin VB.Menu print 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu gap3 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu editmenu 
      Caption         =   "&Edit"
      Begin VB.Menu cut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu gap4 
         Caption         =   "-"
      End
      Begin VB.Menu selectall 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu gap5 
         Caption         =   "-"
      End
      Begin VB.Menu changecase 
         Caption         =   "Chan&ge Case"
         Begin VB.Menu ucase 
            Caption         =   "&UPPERCASE"
         End
         Begin VB.Menu lcase 
            Caption         =   "&lowercase"
         End
         Begin VB.Menu mix 
            Caption         =   "&Mixed"
         End
      End
   End
   Begin VB.Menu searchmenu 
      Caption         =   "&Search"
      Begin VB.Menu find 
         Caption         =   "&Find..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu findnext 
         Caption         =   "Find &Next"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu viewmenu 
      Caption         =   "&View"
      Begin VB.Menu browser 
         Caption         =   "&Browser"
         Shortcut        =   +^{F7}
      End
      Begin VB.Menu disasm 
         Caption         =   "&Disassembly Window"
         Shortcut        =   +^{F8}
      End
   End
   Begin VB.Menu toolsmenu 
      Caption         =   "&Tools"
      Begin VB.Menu explorer 
         Caption         =   "&Explorer"
         Shortcut        =   +^{F3}
      End
      Begin VB.Menu gap7 
         Caption         =   "-"
      End
      Begin VB.Menu compile 
         Caption         =   "Compile &Java"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu runjava 
         Caption         =   "Run Java &Application"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu runapplet 
         Caption         =   "Run Java App&let"
         Begin VB.Menu appletviewer 
            Caption         =   "&Sun JDK Appletviewer"
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu appbro 
            Caption         =   "Web &Browser Support"
            Shortcut        =   ^{F4}
         End
      End
      Begin VB.Menu gap6 
         Caption         =   "-"
      End
      Begin VB.Menu javadoc 
         Caption         =   "Java &Documentation"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu dos 
         Caption         =   "D&OS Shell"
         Shortcut        =   ^{F6}
      End
   End
   Begin VB.Menu windowmenu 
      Caption         =   "&Window"
      Begin VB.Menu th 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu tv 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu cascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu arricons 
         Caption         =   "Arrange &Icons"
      End
   End
   Begin VB.Menu helpmenu 
      Caption         =   "&Help"
      Begin VB.Menu tip 
         Caption         =   "&Tip of the day..."
      End
      Begin VB.Menu gap9 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "&About JACE..."
      End
   End
End
Attribute VB_Name = "mainEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''
''
'' Name: mainEditor
''
'' Decription: This is the main
'' editor that the user is introduced
'' to. Here they can enter, run Java
'' code for applets, applications
'' and can disassemble code.
''
'' Author: RJ45
''
'' Copyright (C) RJ45
''
'' Send email to rj45software@hotmail.com
'' for comments, suggestions and improvements
''''''''''''''''''''''''''''''''''''''''''''''

Private Sub about_Click()
    frmAbout.Show
End Sub

Private Sub appletviewer_Click()
    Dim fullpath As String
    Dim srcfile As String
    Dim counter As Integer
    Dim myfile As String

    'create and initialise an open file dialog box
    appletDialog.DialogTitle = "Java Applet Location" 'title
    appletDialog.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist 'checks
    appletDialog.FileName = "" 'default name that appears in the filename
    'all possible file extensions
    appletDialog.Filter = "Microsoft HTML Document (*.html)|*.htm*|All Files (*.*)|*.*"
    appletDialog.CancelError = True 'catch errors generated
    appletDialog.InitDir = "C:\WINDOWS\JAVA" 'open up in the root dir initially
    
    'this is executed when an error happens
    On Error Resume Next
    appletDialog.ShowSave
    
    'this is executed when no errors are executed
    If Err = 0 Then
        
        fullpath = appletDialog.FileName 'grab the full path of the filename

        srcfile = fullpath
        counter = Len(srcfile) 'determine length of string

        'loop from end of string working backwords checking each
        'character until a backslash is found
        'this means the filename is found
    
        Do Until mychar = "\"
            mychar = Mid(srcfile, counter, 1)
            newstr = mychar + newstr
            counter = counter - 1
        Loop
    
        'the loop will include the backslash, e.g. c:\s.txt, loop will give us '\s.txt' so
        'get rid of leading backslash using the right function assign string to myfile and
        'store in hidden textfield called txtfval

        myfile = Right(newstr, (Len(newstr) - 1))

        ints = InStr(1, fullpath, myfile)
        
        remainder = Left(fullpath, ints - 1)
        'Debug.Print "remainder: " & remainder
        
        res = Shell("c:\runapplet.bat " & remainder & " " & myfile, vbNormalFocus)
        
        'if cancel is pressed execute this
    
        If Err = cdlCancel Then Exit Sub
    End If
    
End Sub

Private Sub arricons_Click()
    mainEditor.Arrange 3
End Sub

Private Sub browser_Click()
    Dim fullpath As String
    Dim srcfile As String
    Dim counter As Integer
    Dim myfile As String

    'create and initialise an open file dialog box
    docDialog.DialogTitle = "Javadoc Location" 'title
    docDialog.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist 'checks
    docDialog.FileName = "" 'default name that appears in the filename
    'all possible file extensions
    docDialog.Filter = "Microsoft HTML Document (*.html)|*.htm*|All Files (*.*)|*.*"
    docDialog.CancelError = True 'catch errors generated
    docDialog.InitDir = "C:\" 'open up in the root dir initially
    
    'this is executed when an error happens
    On Error Resume Next
    docDialog.ShowSave
    
    'this is executed when no errors are executed
    If Err = 0 Then
        
        fullpath = docDialog.FileName 'grab the full path of the filename

        If Format(Left(fullpath, 3), "<") = "c:\" Then
            strip = Right(fullpath, Len(fullpath) - 3)
            frmJavadoc.docs.Navigate2 "file:///c:/" & strip
        End If
        
        frmJavadoc.Show

        'if cancel is pressed execute this
    
        If Err = cdlCancel Then Exit Sub
    End If
End Sub

Private Sub cascade_Click()
    mainEditor.Arrange 0
End Sub

Public Sub close_Click()
    Dim res As Integer
    Dim txt As String
    
    txt = "Document: " & Left(mainEditor.ActiveForm.Caption, _
            Len(mainEditor.ActiveForm.Caption) - 7) & " is not saved. Would like to save it now?"

    res = MsgBox(txt, 35, "Save Document")

    If res = 6 Then
        Call save_Click
        mainEditor.ActiveForm.Hide
    ElseIf res = 7 Then
        'do nothing
    ElseIf res = 2 Then
        'do nothing
    End If
End Sub

Private Sub compile_Click()
    
    Call save_Click
    
    On Error GoTo Trapper

    srcfile = Left(ActiveForm.Caption, Len(ActiveForm.Caption) - 7)
    Debug.Print "c:\jdk1.4\bin\javac.exe " & srcfile
    res = Shell("c:\jdk1.4\bin\javac.exe " & srcfile, vbNormalFocus)
    
    res = MsgBox("Java compilation completed." & vbCrLf & _
                "If any errors exist, the Java program will not run.", vbInformation, _
                "Finished Code Compilation")

Trapper:
    res = MsgBox("No windows are open." & _
            vbCrLf & "Please open a Java source file and then try again", _
            vbCritical, "No Windows Open")

End Sub

Private Sub copy_Click()
    
    On Error GoTo Trapper
    
    Clipboard.SetText Screen.ActiveControl.SelText
    
Trapper:
    res = MsgBox("No windows are open." & _
            vbCrLf & "Please open a Java source file and then try again", _
            vbCritical, "No Windows Open")

End Sub

Private Sub cut_Click()
    On Error GoTo Trapper
    
    Clipboard.SetText Screen.ActiveControl.SelText
    Screen.ActiveControl.SelText = ""
    
Trapper:
    res = MsgBox("No windows are open." & _
            vbCrLf & "Please open a Java source file and then try again", _
            vbCritical, "No Windows Open")

End Sub

Private Sub disasm_Click()
    Dim fullpath As String
    Dim srcfile As String
    Dim counter As Integer
    Dim myfile As String

    'create and initialise an open file dialog box
    DisAsmDialog.DialogTitle = "Java Class File Location" 'title
    DisAsmDialog.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist 'checks
    DisAsmDialog.FileName = "" 'default name that appears in the filename
    'all possible file extensions
    DisAsmDialog.Filter = "Java Class File (*.class)|*.class|All Files (*.*)|*.*"
    DisAsmDialog.CancelError = True 'catch errors generated
    DisAsmDialog.InitDir = "C:\WINDOWS\JAVA" 'open up in the root dir initially
    
    'this is executed when an error happens
    On Error Resume Next
    DisAsmDialog.ShowSave
    
    'this is executed when no errors are executed
    If Err = 0 Then
        
        fullpath = DisAsmDialog.FileName 'grab the full path of the filename

        srcfile = fullpath
        counter = Len(srcfile) 'determine length of string

        'loop from end of string working backwords checking each
        'character until a backslash is found
        'this means the filename is found
    
        Do Until mychar = "\"
            mychar = Mid(srcfile, counter, 1)
            newstr = mychar + newstr
            counter = counter - 1
        Loop
    
        'the loop will include the backslash, e.g. c:\s.txt, loop will give us '\s.txt' so
        'get rid of leading backslash using the right function assign string to myfile and
        'store in hidden textfield called txtfval

        myfile = Right(newstr, (Len(newstr) - 1))

        ints = InStr(1, fullpath, myfile)
        
        remainder = Left(fullpath, ints - 1)
        
        res = Shell("c:\jdisasm.bat " & remainder & " " & Left(myfile, Len(myfile) - 6), vbNormalFocus)
        
        Open "c:\windows\java\disasm.txt" For Input As #1
            While Not EOF(1)
                Line Input #1, textline
                'set the homepage in the Internet Options form
                frmDisasm.Text1.Text = _
                frmDisasm.Text1.Text & textline & vbCrLf
            Wend
        Close #1
        
        frmDisasm.Show
        
        'if cancel is pressed execute this
        'If Err = cdlCancel Then Exit Sub
        
    End If
End Sub

Private Sub dos_Click()
    res = Shell("C:\WINDOWS\command.com", vbNormalFocus)
End Sub

Private Sub exit_Click()
    Call save_Click
End Sub

Private Sub explorer_Click()
    res = Shell("C:\WINDOWS\Explorer.exe", vbNormalFocus)
End Sub

Private Sub find_Click()
    
    On Error GoTo Trapper
    
    Dim Search, Where, srcText   ' Declare variables.
    ' Get search string from user.
    
    srcText = mainEditor.ActiveForm.ActiveControl
    Search = InputBox("Enter text to be found:", "Search String Required")
    Where = InStr(srcText, Search)   ' Find string in text.
    
    If Where Then   ' If found,
        ActiveForm.ActiveControl.SelStart = Where - 1
        ActiveForm.ActiveControl.SelLength = Len(Search)
    Else
        res = MsgBox("String not found.", vbInformation, "Search Results")
    End If
    
Trapper:
    res = MsgBox("No windows are open." & _
            vbCrLf & "Please open a Java source file and then try again", _
            vbCritical, "No Windows Open")
End Sub

Private Sub javadoc_Click()
    frmJavadoc.docs.Navigate2 "file:///c:/jdk1.4/docs/api/index.html"
End Sub

Private Sub lcase_Click()
    
    On Error GoTo Trapper
    
    Dim src
    
    src = mainEditor.ActiveForm.ActiveControl
    
    If src > 0 Then
        low = Format(src, "<")
        ActiveForm.ActiveControl.SelText = low
    End If

Trapper:
    res = MsgBox("No windows are open." & _
            vbCrLf & "Please open a Java source file and then try again", _
            vbCritical, "No Windows Open")

End Sub


Private Sub mix_Click()
    
    On Error GoTo Trapper
    
    Dim src
    
    src = mainEditor.ActiveForm.ActiveControl
    
    If src > 0 Then
        strChar = Left(src, 1)
        leftOver = Right(src, Len(src) - 1)
        low = Format(src, "<")
        ActiveForm.ActiveControl.SelText = Format(strChar, ">") & Format(leftOver, "<")
    End If

Trapper:
    res = MsgBox("No windows are open." & _
            vbCrLf & "Please open a Java source file and then try again", _
            vbCritical, "No Windows Open")

End Sub

Private Sub new_Click()
    Dim newSrc As frmEditor
    
    Set newSrc = New frmEditor
    Load newSrc
    newSrc.Show
End Sub

Private Sub open_Click()
    
    Dim newSrc As frmEditor
    
    Set newSrc = New frmEditor
    
    'declare variables
    Dim fullpath As String
    
    'create and initialise an open file dialog box
    openJava.DialogTitle = "Open" 'title
    openJava.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist 'checks
    openJava.FileName = "" 'default name that appears in the filename
    'all possible file extensions
    openJava.Filter = "Java Source File (*.java)|*.java|All Files (*.*)|*.*"
    openJava.CancelError = True 'catch errors generated
    openJava.InitDir = "C:\WINDOWS\JAVA" 'open up in the root dir initially
    
    'this is executed when an error happens
    On Error Resume Next
    openJava.ShowOpen
    
    'this is executed when no errors are executed
    If Err = 0 Then
        fullpath = openJava.FileName 'grab the full path of the filename
        Load newSrc
        newSrc.Show
        
        mainEditor.ActiveForm.Caption = fullpath & " - JACE"
        
        Open fullpath For Input As #1
            While Not EOF(1)
                Line Input #1, textline
                'set the homepage in the Internet Options form
                mainEditor.ActiveForm.txtsrc.Text = _
                        mainEditor.ActiveForm.txtsrc.Text & textline & vbCrLf
            Wend
        Close #1

    End If
    'if cancel is pressed execute this
    'If Err = cdlCancel Then Exit Sub

End Sub

Private Sub paste_Click()

    On Error GoTo Trapper
    
    Screen.ActiveControl.SelText = Clipboard.GetText
    
Trapper:
    res = MsgBox("No windows are open." & _
            vbCrLf & "Please open a Java source file and then try again", _
            vbCritical, "No Windows Open")

End Sub

Private Sub appbro_Click()
    Dim fullpath As String
    Dim srcfile As String
    Dim counter As Integer
    Dim myfile As String

    'create and initialise an open file dialog box
    appletDialog.DialogTitle = "Java Applet Location" 'title
    appletDialog.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist 'checks
    appletDialog.FileName = "" 'default name that appears in the filename
    'all possible file extensions
    appletDialog.Filter = "Microsoft HTML Document (*.html)|*.htm*|All Files (*.*)|*.*"
    appletDialog.CancelError = True 'catch errors generated
    appletDialog.InitDir = "C:\WINDOWS\JAVA" 'open up in the root dir initially
    
    'this is executed when an error happens
    On Error Resume Next
    appletDialog.ShowSave
    
    'this is executed when no errors are executed
    If Err = 0 Then
        
        fullpath = appletDialog.FileName 'grab the full path of the filename
        

        srcfile = fullpath
        counter = Len(srcfile) 'determine length of string

        'loop from end of string working backwords checking each
        'character until a backslash is found
        'this means the filename is found
    
        Do Until mychar = "\"
            mychar = Mid(srcfile, counter, 1)
            newstr = mychar + newstr
            counter = counter - 1
        Loop
    
        'the loop will include the backslash, e.g. c:\s.txt, loop will give us '\s.txt' so
        'get rid of leading backslash using the right function assign string to myfile and
        'store in hidden textfield called txtfval

        myfile = Right(newstr, (Len(newstr) - 1))

        If Format(Left(fullpath, 3), "<") = "c:\" Then
            strip = Right(fullpath, Len(fullpath) - 3)
            frmApplet.Text1.Text = "file:///c:/" & strip
            frmApplet.appletWindow.Navigate "file:///c:/" & strip
        End If
        
        frmApplet.Show

        'if cancel is pressed execute this
    
        'If Err = cdlCancel Then Exit Sub
    End If
End Sub

Private Sub runjava_Click()
    Call save_Click
    
    On Error GoTo Trapper

    'Define variables
    
    Dim srcfile As String
    Dim counter As Integer
    Dim myfile As String

    srcfile = Left(ActiveForm.Caption, Len(ActiveForm.Caption) - 12)
    counter = Len(srcfile) 'determine length of string

    'loop from end of string working backwords checking each
    'character until a backslash is found
    'this means the filename is found
    
    Do Until mychar = "\"
        mychar = Mid(srcfile, counter, 1)
        newstr = mychar + newstr
        counter = counter - 1
    Loop
    
    'the loop will include the backslash, e.g. c:\s.txt, loop will give us '\s.txt' so
    'get rid of leading backslash using the right function assign string to myfile and
    'store in hidden textfield called txtfval

    myfile = Right(newstr, (Len(newstr) - 1))

    Debug.Print "c:\jdk1.4\bin\java.exe " & myfile
    res = Shell("c:\jdk1.4\bin\java.exe " & myfile, vbNormalFocus)
    
    res = MsgBox("Java program execution completed." & vbCrLf & _
                "Check to see if your Java program requires any special parameters to be submitted." & vbCrLf & _
                "If so please startup DOS and manually run the program.", vbInformation, _
                "Finished Code Compilation")

Trapper:
    res = MsgBox("No windows are open." & _
            vbCrLf & "Please open a Java source file and then try again", _
            vbCritical, "No Windows Open")

End Sub

Private Sub save_Click()
    
    On Error Resume Next
    
    Dim id As Integer
    
    id = InStr(1, mainEditor.ActiveForm.Caption, "Untitled")
    
    If id > 0 Then

        'declare variables
        Dim fullpath As String
    
        'create and initialise an open file dialog box
        saveJava.DialogTitle = "Save" 'title
        saveJava.Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist 'checks
        saveJava.FileName = "" 'default name that appears in the filename
        'all possible file extensions
        saveJava.Filter = "Java Source File (*.java)|*.java|All Files (*.*)|*.*"
        saveJava.CancelError = True 'catch errors generated
        saveJava.InitDir = "C:\WINDOWS\JAVA" 'open up in the root dir initially
    
        'this is executed when an error happens
        On Error Resume Next
        saveJava.ShowSave
    
        'this is executed when no errors are executed
        If Err = 0 Then
        
            fullpath = saveJava.FileName 'grab the full path of the filename
        
        Open fullpath For Output As #1
             Print #1, Trim(mainEditor.ActiveForm.txtsrc.Text)
        Close #1
               
            mainEditor.ActiveForm.Caption = fullpath & " - JACE"

        End If
        
        'if cancel is pressed execute this
        'If Err = cdlCancel Then Exit Sub
    Else
        
        Open Left(mainEditor.ActiveForm.Caption, _
                        Len(mainEditor.ActiveForm.Caption) - 7) For Output As #2
             Print #2, Trim(mainEditor.ActiveForm.txtsrc.Text)
        Close #2
        
        mainEditor.ActiveForm.Caption = Trim(Left(mainEditor.ActiveForm.Caption, _
                        Len(mainEditor.ActiveForm.Caption) - 7)) & " - JACE"

    End If
    
End Sub

Private Sub saveas_Click()

    On Error GoTo Trapper
    
    'declare variables
    Dim fullpath As String
    
    'create and initialise an open file dialog box
    saveJava.DialogTitle = "Save" 'title
    saveJava.Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist 'checks
    saveJava.FileName = "" 'default name that appears in the filename
    'all possible file extensions
    saveJava.Filter = "Java Source File (*.java)|*.java|All Files (*.*)|*.*"
    saveJava.CancelError = True 'catch errors generated
    saveJava.InitDir = "C:\WINDOWS\JAVA" 'open up in the root dir initially
    
    'this is executed when an error happens
    On Error Resume Next
    saveJava.ShowSave
    
    'this is executed when no errors are executed
    If Err = 0 Then
        
        fullpath = saveJava.FileName 'grab the full path of the filename
        
        Open fullpath For Output As #1
            Print #1, Trim(mainEditor.ActiveForm.txtsrc.Text)
        Close #1
               
        mainEditor.ActiveForm.Caption = fullpath & " - JACE"

    End If
        
    'if cancel is pressed execute this
    'If Err = cdlCancel Then Exit Sub

Trapper:
    res = MsgBox("No windows are open." & _
            vbCrLf & "Please open a Java source file and then try again", _
            vbCritical, "No Windows Open")

End Sub

Private Sub selectall_Click()
    
    On Error GoTo Trapper
    
    act = ActiveForm.ActiveControl
    ActiveForm.ActiveControl.SelLength = Len(act)
    
Trapper:
    res = MsgBox("No windows are open." & _
            vbCrLf & "Please open a Java source file and then try again", _
            vbCritical, "No Windows Open")

End Sub

Private Sub th_Click()
    mainEditor.Arrange 1
End Sub

Private Sub tip_Click()
    frmTip.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   'this code handles the toolbar buttons control.
    
    'base the result on what key is obtained when the button
    'is clicked
    
    Select Case Button.Key
    
    'new button is pressed
    Case Is = "ne"
        On Error Resume Next
        Call new_Click
    
    'open button is pressed
    Case Is = "op"
        Call open_Click
    
    'save button is pressed
    Case Is = "sa"
        Call save_Click
    
    'print button is pressed
    Case Is = "pr"
        
    
    'cut button is pressed
    Case Is = "cu"
        Call cut_Click
    
    'copy button is pressed
    Case Is = "co"
        Call copy_Click
    
    'paste button is pressed
    Case Is = "pa"
        Call paste_Click
    
    'web browser button is pressed
    Case Is = "we"
        Call browser_Click
    
    'find button is pressed
    Case Is = "fi"
        Call find_Click
    
    'find again button is pressed
    Case Is = "fa"
        
        
    End Select
End Sub

Private Sub tv_Click()
    mainEditor.Arrange 2
End Sub

Private Sub ucase_Click()
        
    On Error GoTo Trapper
    
    Dim src
    
    src = mainEditor.ActiveForm.ActiveControl
    
    If src > 0 Then
        upper = Format(src, ">")
        ActiveForm.ActiveControl.SelText = upper
    End If

Trapper:
    res = MsgBox("No windows are open." & _
            vbCrLf & "Please open a Java source file and then try again", _
            vbCritical, "No Windows Open")

End Sub
