Attribute VB_Name = "modServer"
'All of the general server functions go here.
Option Explicit

Public Const maxCon = 20
Public Const listenPortNum = 1337

Public bSocketUsed(maxCon) As Boolean


'Create the sockets for accepting the connections
Public Function InitAccept() As Boolean
    Dim i As Integer
    
    For i = 1 To maxCon
        'load the sockets
        Load frmServer.tcpAccept(i)
        'set as unused
        bSocketUsed(i) = False
    Next i
    
    'all is well, nothing is ruined
    InitAccept = True
    Exit Function
    
err:
    'flagrant system error
    InitAccept = False
End Function

'find unused socket
Public Function GetFreeSocket() As Integer
    Dim i As Integer
    
    'Find unused socket
    For i = 1 To maxCon
        If Not bSocketUsed(i) Then
            GetFreeSocket = i
            Exit Function
        End If
    Next i
    
    'no socket found
    GetFreeSocket = 0
End Function

'output some log text
Public Sub Log(text As String)
    frmServer.rtLog.SelStart = Len(frmServer.rtLog.text)
    frmServer.rtLog.SelText = text & vbNewLine
End Sub

'get message
Public Function getMessage() As String
    'show dialog box
    frmMessage.Visible = True
    frmMessage.txtMessage.SetFocus
    
    'wait for dialog to close
    Do
        DoEvents
    Loop Until frmMessage.Visible = False
    
    'get text
    getMessage = frmMessage.sMessage
End Function
