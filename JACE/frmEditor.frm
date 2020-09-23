VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEditor 
   Caption         =   "Untitled - JACE"
   ClientHeight    =   4335
   ClientLeft      =   2760
   ClientTop       =   1215
   ClientWidth     =   6495
   Icon            =   "frmEditor.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4335
   ScaleWidth      =   6495
   Begin MSComDlg.CommonDialog saveJava 
      Left            =   5400
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtsrc 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''
''
'' Name: frmEditor
''
'' Decription: This window allows the user to
'' type up their Java programs and applets.
''
'' Author: RJ45
''
'' Copyright (C) RJ45
''
'' Send email to rj45software@hotmail.com
'' for comments, suggestions and improvements
''''''''''''''''''''''''''''''''''''''''''''''


Private Sub Form_Resize()
    txtsrc.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim title, prompt As String
    title = mainEditor.ActiveForm.Caption
    prompt$ = "Document: " & Left(title, Len(title) - 7) & " has not been saved. Would like to save it now?"
    res = MsgBox(prompt$, 35, "Save Document")
    
    If res = 6 Then
        Call saveDoc
        mainEditor.ActiveForm.Hide
    ElseIf res = 7 Then
        'do nothing
    ElseIf res = 2 Then
        mainEditor.ActiveForm.Caption = mainEditor.ActiveForm.Caption
    End If
End Sub

Private Sub txtsrc_Change()
    If InStr(1, mainEditor.ActiveForm.Caption, "*") = 0 Then
        mainEditor.ActiveForm.Caption = mainEditor.ActiveForm.Caption & "*"
    End If
End Sub

Private Sub saveDoc()
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
        If Err = cdlCancel Then Exit Sub
    Else
        
        Open Left(mainEditor.ActiveForm.Caption, _
                        Len(mainEditor.ActiveForm.Caption) - 7) For Output As #2
             Print #2, Trim(mainEditor.ActiveForm.txtsrc.Text)
        Close #2
        
        mainEditor.ActiveForm.Caption = Trim(Left(mainEditor.ActiveForm.Caption, _
                        Len(mainEditor.ActiveForm.Caption) - 7)) & " - JACE"

    End If
End Sub

Public Sub selall()
    txtsrc.SelLength = Len(txtsrc.Text)
End Sub
