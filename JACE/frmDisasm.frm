VERSION 5.00
Begin VB.Form frmDisasm 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Disassembly Window"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   6600
      Top             =   600
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "frmDisasm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''
''
'' Name: frmDisasm
''
'' Decription: This window shows the class file
'' code disassembled into JVM instructions.
''
'' Author: RJ45
''
'' Copyright (C) RJ45
''
'' Send email to rj45software@hotmail.com
'' for comments, suggestions and improvements
''''''''''''''''''''''''''''''''''''''''''''''

Dim i As Integer

Private Sub Form_Load()
    i = 0
End Sub

Private Sub Form_Resize()
    Text1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Form_Terminate()
    i = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    i = 0
End Sub

Private Sub Timer1_Timer()
    
    If i = 0 Then
        Open "c:\windows\java\disasm.txt" For Input As #1
            While Not EOF(1)
                Line Input #1, textline
                'set the homepage in the Internet Options form
                frmDisasm.Text1.Text = _
                frmDisasm.Text1.Text & textline & vbCrLf
            Wend
        Close #1
        i = i + 1
    End If
End Sub
