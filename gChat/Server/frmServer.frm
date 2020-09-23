VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmServer 
   Caption         =   "gChat Server"
   ClientHeight    =   5580
   ClientLeft      =   5025
   ClientTop       =   5295
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSendAll 
      Caption         =   "Send To All"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSndMsg 
      Caption         =   "Send Message"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdKick 
      Caption         =   "Kick"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
   End
   Begin VB.ComboBox cmbClients 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   4560
      Width           =   5775
   End
   Begin MSWinsockLib.Winsock tcpAccept 
      Index           =   0
      Left            =   6840
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock tcpListen 
      Left            =   6360
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtLog 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7223
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmServer.frx":0000
   End
   Begin VB.Label Label2 
      Caption         =   "Admin Functions"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Log"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A multi-connection server for my new chat client.
'Created by Guillaume Couture-Levesque on July 13th 2003.

Option Explicit

Private Sub cmdExit_Click()
    Dim i As Integer
    
    For i = 1 To maxCon
        If bSocketUsed(i) Then
            'send shutdown message
            tcpAccept(i).SendData ("Server shutting down." & vbNewLine)
            DoEvents
            'close connection
            tcpAccept(i).Close
            DoEvents
        End If
    Next i
    
    'quit program
    End
End Sub

Private Sub cmdKick_Click()
    Dim clientID As Integer
    
    'If there's no client selected
    If cmbClients.text = "" Then
        MsgBox "Please choose a client to kick.", vbCritical, "Error"
        Exit Sub
    End If
    
    'find the client number
    clientID = Split(cmbClients.text, " ")(1)
    
    'notify the client
    tcpAccept(clientID).SendData ("You've been kicked." & vbNewLine)
    DoEvents
    
    'Log the information
    Log ("Client " & clientID & " (" & tcpAccept(clientID).RemoteHostIP & ") got kicked.")
    
    'close the connection
    tcpAccept(clientID).Close
    
    'set port as unused
    bSocketUsed(clientID) = False
    
    'update combo box
    updateCombo
End Sub

Private Sub cmdSendAll_Click()
    Dim text As String
    Dim i As Integer
    
    'get text to send
    text = getMessage
    
    'check for cancel
    If text = "" Then
        Exit Sub
    End If
    
    'send message
    For i = 1 To maxCon
        If bSocketUsed(i) Then
            tcpAccept(i).SendData (text & vbNewLine)
            DoEvents
        End If
    Next i
    
    'log info
    Log ("Message sent to all clients.")
End Sub

Private Sub cmdSndMsg_Click()
    Dim clientID As Integer
    Dim text As String
    
    'If there's no client selected
    If cmbClients.text = "" Then
        MsgBox "Please choose a client to send a message to.", vbCritical, "Error"
        Exit Sub
    End If
    
    'find the client number
    clientID = Split(cmbClients.text, " ")(1)
    
    'get text to send
    text = getMessage
    
    'check for cancel
    If text = "" Then
        Exit Sub
    End If
    
    'send message
    tcpAccept(clientID).SendData (text & vbNewLine)
    
    'log info
    Log ("Message sent to client " & clientID & ".")
End Sub

Private Sub Form_Load()
    'start listening for connections
    tcpListen.LocalPort = listenPortNum
    tcpListen.Listen
    
    'create the slots for the accept sockets
    If Not InitAccept Then
        MsgBox "Error creating sockets.", vbCritical, "Socket Error!"
        End
    End If
    
    'Set Combo box
    cmbClients.Clear
    
    'Tell User that the log is started
    Log ("Application Started")
    Log ("Awaiting Connections")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frm As Form
    
    'Unload all of the other forms
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next frm
End Sub

Private Sub tcpAccept_Close(Index As Integer)
    'close connection
    tcpAccept(Index).Close
    'Log this fact
    Log ("Connection closed for client " & Index & ".")
    'set the port to open
    bSocketUsed(Index) = False
    'update the combo box to reflect changes
    updateCombo
End Sub

Private Sub tcpAccept_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim text As String
    Dim i As Integer
    
    'get data
    tcpAccept(Index).GetData text
    
    'check for disconnects
    If text = ("disband") Then
        tcpAccept_Close (Index)
        Exit Sub
    End If
    
    'log it
    Log (Left("Client " & Index & ": " & text, Len("Client " & Index & ": " & text) - 1))
    
    'send it out
    For i = 1 To maxCon
        If bSocketUsed(i) Then
            tcpAccept(i).SendData (text)
            DoEvents
        End If
    Next i
End Sub

Private Sub tcpListen_ConnectionRequest(ByVal requestID As Long)
    Dim portNum As Integer
    
    'find out if there's a free socket
    portNum = GetFreeSocket
    
    If portNum = 0 Then
        'tell the client that the server is full
        tcpAccept(0).Accept (requestID)
        DoEvents
        tcpAccept(0).SendData "Sorry, maximum number of connections reached!"
        DoEvents
        'log the fact that the server is full
        Log ("Maximum number of connections reached.")
        Log ("Client at IP: " & tcpAccept(0).RemoteHostIP & " denied.")
        'close the client's connection
        tcpAccept(0).Close
    Else
        'accept the client using the accepting winsock
        tcpAccept(portNum).Accept (requestID)
        DoEvents
        'log the fact that there's a new client
        Log ("New client ID " & portNum & " at IP " & tcpAccept(portNum).RemoteHostIP & ".")
        'send welcome message
        tcpAccept(portNum).SendData ("Connection accepted, welcome." & vbNewLine)
        'set port to used
        bSocketUsed(portNum) = True
        'update combo box to reflect changes
        updateCombo
    End If
End Sub

Private Sub updateCombo()
    Dim i As Integer
    
    'clear the list
    cmbClients.Clear
    
    'set list
    For i = 1 To maxCon
        If bSocketUsed(i) Then
            cmbClients.AddItem ("Client " & i & " (" & tcpAccept(i).RemoteHostIP & ")")
        End If
    Next i
End Sub
