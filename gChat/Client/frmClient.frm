VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmClient 
   Caption         =   "gChat Client"
   ClientHeight    =   3975
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   2760
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   735
      Left            =   3480
      TabIndex        =   2
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtMsg 
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   3240
      Width           =   3495
   End
   Begin RichTextLib.RichTextBox rtText 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4048
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmClient.frx":0000
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuPref 
         Caption         =   "Preferences..."
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuCnction 
      Caption         =   "Connection"
      Begin VB.Menu mnuCnct 
         Caption         =   "Connect..."
      End
      Begin VB.Menu mnuDis 
         Caption         =   "Disconnect"
      End
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sName, sOldName As String

Private Sub cmdSend_Click()
    'send the message to the server
    tcpClient.SendData (sName & ": " & txtMsg.text & vbNewLine)
    txtMsg.text = ""
    txtMsg.SetFocus
End Sub

Private Sub Form_Load()

    'Set size
    Form_Resize
    
    'make sure you have to connect to send a message
    cmdSend.Enabled = False
    
    'make client window visible
    frmClient.Visible = True
    
    'Get Name
    frmPrefs.cmdCancel.Visible = False
    frmPrefs.cmdOK.Left = 960
    mnuPref_Click
    frmPrefs.cmdCancel.Visible = True
    frmPrefs.cmdOK.Left = 360
    
    'Connect to server
    mnuCnct_Click
End Sub

Private Sub Form_Resize()
    'check if minimized
    If frmClient.WindowState = 1 Then
        Exit Sub
    End If
    
    'check for minimum size
    If frmClient.Width < 4000 Then
        frmClient.Width = 4000
    End If
    If frmClient.Height < 4000 Then
        frmClient.Height = 4000
    End If
    
    'Set size for objects
    txtMsg.Top = frmClient.ScaleHeight - txtMsg.Height
    txtMsg.Width = frmClient.ScaleWidth - cmdSend.Width
    rtText.Width = frmClient.ScaleWidth
    rtText.Height = frmClient.ScaleHeight - txtMsg.Height
    cmdSend.Top = txtMsg.Top
    cmdSend.Height = txtMsg.Height
    cmdSend.Left = txtMsg.Width
    cmdSend.Width = frmClient.ScaleWidth - txtMsg.Width
End Sub

Private Sub Status(text As String)
    'print out a status line
    rtText.SelColor = vbBlue
    rtText.SelItalic = True
    rtText.SelStart = Len(rtText.text)
    rtText.SelText = text & vbNewLine
    rtText.SelItalic = False
    rtText.SelColor = vbBlack
End Sub

Private Sub WriteMsg(text As String)
    'print out a message
    rtText.SelStart = Len(rtText.text)
    rtText.SelText = text
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frm As Form
    
    'Unload all of the other forms
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next frm
End Sub

Private Sub mnuCnct_Click()
    Dim IP As String
    
    'make sure you're not already connected
    If tcpClient.State <> sckClosed Then
        MsgBox "Already connected.", vbOKOnly, "Connection Error"
        Exit Sub
    End If
    
    'Get connection IP
    frmIP.Visible = True
    frmIP.txtIP.SetFocus
    Do
        DoEvents
    Loop Until Not frmIP.Visible
    IP = frmIP.txtIP.text
    
    'check for valid IP
    If IP = "" Then
        Exit Sub
    End If
    
    'Connect to server
    Status ("Connecting to server...")
    Do Until tcpClient.State <> sckClosed
        DoEvents
        tcpClient.Close
        tcpClient.Connect IP, 1337
    Loop
    
    'set focus to message box
    txtMsg.SetFocus
End Sub

Private Sub mnuDis_Click()
    'make sure the connection's not already closed
    If tcpClient.State = sckClosed Then
        Exit Sub
    End If
    
    'notify of disconnect
    tcpClient.SendData (sName & " has disconnected." & vbNewLine)
    DoEvents
    tcpClient.SendData ("disband")
    DoEvents
    
    'close client
    tcpClient.Close
    Status ("Disconnected")
    DoEvents
    cmdSend.Enabled = False
End Sub

Private Sub mnuExit_Click()
    'disconnect
    mnuDis_Click
    'end program
    End
End Sub

Private Sub mnuPref_Click()
    'save old name
    sOldName = sName
    
    'Get prefs
    frmPrefs.Visible = True
    frmPrefs.txtName.SetFocus
    Do
        DoEvents
    Loop Until Not frmPrefs.Visible
    
    'check if name is too long
    If Len(sName) > 32 Then
        sName = Left(sName, 32)
    End If
    
    'check if name has changed
    If (sName <> sOldName) And (tcpClient.State <> sckClosed) Then
        tcpClient.SendData (sOldName & " is now called " & sName & vbNewLine)
    End If
    
    'set focus to message box
    txtMsg.SetFocus
End Sub

Private Sub tcpClient_Close()
    'notify client of disconnect
    tcpClient.Close
    DoEvents
    Status ("Disconnected")
    cmdSend.Enabled = False
End Sub

Private Sub tcpClient_Connect()
    'notify client of connect
    Status ("Connected")
    DoEvents
    'notify server of new client arrival
    tcpClient.SendData (sName & " has connected." & vbNewLine)
    cmdSend.Enabled = True
End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
    Dim text As String
    
    tcpClient.GetData text
    WriteMsg (text)
End Sub

Private Sub tcpClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
End Sub

Private Sub txtMsg_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) And (cmdSend.Enabled) Then
        cmdSend_Click
    End If
End Sub
