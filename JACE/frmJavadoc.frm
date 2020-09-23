VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmJavadoc 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Javadoc Window"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser docs 
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      ExtentX         =   19076
      ExtentY         =   15055
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "res://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/dnserror.htm#http:///"
   End
End
Attribute VB_Name = "frmJavadoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''
''
'' Name: frmJavadoc
''
'' Decription: This form shows a browser object
'' so that users can view their Javadoc
'' documentation
''
'' Author: RJ45
''
'' Copyright (C) RJ45
''
'' Send email to rj45software@hotmail.com
'' for comments, suggestions and improvements
''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Form_Resize()
    docs.Width = ScaleWidth
    docs.Height = ScaleHeight
End Sub
