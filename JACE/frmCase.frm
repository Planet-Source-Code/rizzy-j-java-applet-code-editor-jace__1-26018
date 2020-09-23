VERSION 5.00
Begin VB.Form frmCase 
   Caption         =   "Change Case"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.OptionButton optLCase 
      Caption         =   "lowercase"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Caption         =   "First character uppercase, rest lowercase"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3255
   End
   Begin VB.OptionButton optUCase 
      Caption         =   "UPPERCASE"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Dim startChar, leftOver As String
    
    If optUCase.Value = True Then
        frmCase.Hide
        Debug.Print ActiveForm.ActiveControl.SelText
        ucase (ActiveForm.ActiveControl.SelText)
    ElseIf optLCase.Value = True Then
        frmCase.Hide
        lcase (ActiveForm.ActiveControl.SelText)
    Else
        frmCase.Hide
        strchar = Left(ActiveForm.ActiveControl.SelText, 1)
        leftOver = Right(ActiveForm.ActiveControl.SelText, Len(ActiveForm.ActiveControl.SelText) - 1)
        ActiveForm.ActiveControl.SelText = ucase(strchar) & lcase(leftOver)
    End If
End Sub
