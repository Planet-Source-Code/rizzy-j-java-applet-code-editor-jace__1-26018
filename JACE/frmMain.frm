VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Untitled - JACE"
   ClientHeight    =   9510
   ClientLeft      =   2400
   ClientTop       =   885
   ClientWidth     =   10695
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog saveJava 
      Left            =   8760
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog openJava 
      Left            =   9360
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8760
      Top             =   1800
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
            Picture         =   "frmMain.frx":0ABA
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C82
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D82
            Key             =   "save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E6E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FB2
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":125A
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1396
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14DE
            Key             =   "viewbrowser"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16BA
            Key             =   "find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1896
            Key             =   "findagain"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   9255
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
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
         Caption         =   "Chan&ge Case..."
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
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu gap6 
         Caption         =   "-"
      End
      Begin VB.Menu javadoc 
         Caption         =   "Java &Documentation"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu dos 
         Caption         =   "D&OS Shell"
         Shortcut        =   ^{F5}
      End
   End
   Begin VB.Menu helpmenu 
      Caption         =   "&Help"
      Begin VB.Menu about 
         Caption         =   "&About JACE..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
    'scales the dimensions of the components of the text editor
    'when the form size changes
    
End Sub

Private Sub open_Click()
    
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
        
        Open fullpath For Input As #1
            While Not EOF(1)
                Line Input #1, textline
                'set the homepage in the Internet Options form
                txtsrc.Text = txtsrc.Text & textline & vbCrLf
            Wend
        Close #1
               
        frmMain.Caption = fullpath & " - JACE"

    End If
    'if cancel is pressed execute this
    If Err = cdlCancel Then Exit Sub
End Sub

Private Sub save_Click()
    
    If (InStr(0, frmMain.Caption, "Untitled")) > 0 Then
            
        'declare variables
        Dim fullpath As String
    
        'create and initialise an open file dialog box
        saveJava.DialogTitle = "Save" 'title
        saveJava.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist 'checks
        saveJava.FileName = "" 'default name that appears in the filename
        'all possible file extensions
        saveJava.Filter = "Java Source File (*.java)|*.java|All Files (*.*)|*.*"
        saveJava.CancelError = True 'catch errors generated
        saveJava.InitDir = "C:\WINDOWS\JAVA" 'open up in the root dir initially
    
        'this is executed when an error happens
        On Error Resume Next
        saveJava.ShowOpen
    
        'this is executed when no errors are executed
        If Err = 0 Then
        
            fullpath = saveJava.FileName 'grab the full path of the filename
        
            Open fullpath For Input As #1
                While Not EOF(1)
                    Line Input #1, textline
                    'set the homepage in the Internet Options form
                    txtsrc.Text = txtsrc.Text & textline & vbCrLf
                Wend
            Close #1
               
            frmMain.Caption = fullpath & " - JACE"

        End If
        'if cancel is pressed execute this
        If Err = cdlCancel Then Exit Sub
    Else
        
End Sub

Private Sub txtsrc_Change()
    If InStr(1, frmMain.Caption, "*") = 0 Then
        frmMain.Caption = frmMain.Caption & "*"
    End If
End Sub
