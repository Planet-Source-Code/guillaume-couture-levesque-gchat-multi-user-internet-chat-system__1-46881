VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message To Send"
   ClientHeight    =   1560
   ClientLeft      =   7500
   ClientTop       =   7440
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtMessage 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   3975
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter message to send:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public sMessage As String

Private Sub CancelButton_Click()
    sMessage = ""
    frmMessage.Visible = False
    txtMessage.text = ""
End Sub

Private Sub OKButton_Click()
    sMessage = txtMessage.text
    frmMessage.Visible = False
    txtMessage.text = ""
End Sub
