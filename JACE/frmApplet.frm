VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmApplet 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "View Applets Window"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   9975
   End
   Begin SHDocVwCtl.WebBrowser appletWindow 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   10455
      ExtentX         =   18441
      ExtentY         =   14631
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
   Begin VB.Label Label1 
      Caption         =   "URL:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmApplet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''
''
'' Name: frmApplet
''
'' Decription: This window shows a browser so
'' that the user can view applets
''
'' Author: RJ45
''
'' Copyright (C) RJ45
''
'' Send email to rj45software@hotmail.com
'' for comments, suggestions and improvements
''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Form_Resize()
    appletWindow.Width = ScaleWidth
    appletWindow.Height = ScaleHeight - 360
    Text1.Width = ScaleWidth - 960
End Sub
