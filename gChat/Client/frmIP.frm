VERSION 5.00
Begin VB.Form frmIP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connection"
   ClientHeight    =   930
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lbl1 
      Caption         =   "Connect to what IP?"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()
    frmIP.Visible = False
End Sub

Private Sub txtIP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        OKButton_Click
    End If
End Sub
