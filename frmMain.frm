VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "frmMain"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   ""
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sockListener 
      Left            =   240
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sockClient 
      Index           =   0
      Left            =   840
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Binary

'TODO
' Global block array

'Port to listen on
Private Const PORT As Long = 1503

'Path to html files (must not end in a \)
Private Const HTDOCS As String = "N:"

'Default html file
Private Const DEFAULT_FILE As String = "index.html"

'Size of blocks to send
Private Const BLOCK_SIZE As Long = 16384	'16 KB

'Information about a connection
Private Type ConnectionInfo
    FreeChain As Long       'Pointer to next free connection or -1 if this is in use
    FileNbr As Integer      'File number being read (-1 = request not received, -2 = nothing left to send)
End Type

'Connection list
Private conns() As ConnectionInfo

'Pointer to first free connection
Private freePtr As Long

Private Sub GrowConnections()
    'Double connection count
    Dim firstNew As Long
    firstNew = UBound(conns) + 1
    
    'Resize array
    ReDim Preserve conns((UBound(conns) + 1) * 2 - 1)
    
    'Initialize array
    Dim i As Long
    For i = firstNew To UBound(conns)
        'Add to free chain
        conns(i).FreeChain = freePtr
        freePtr = i
        
        'Create socket
        Load sockClient(i)
    Next
End Sub

Private Sub SendError(ByVal id As Integer, ByVal str As String)
    'Send error response
    sockClient(id).SendData "HTTP/1.0 " & str & vbCrLf & vbCrLf
    conns(id).FileNbr = -2
End Sub

Private Sub DoGetRequest(ByVal id As Integer, ByRef filename As String)
    Dim newFilename As String
    Dim FileNbr As Long
    Dim i As Long
    
    'Force \ at beginning
    newFilename = "\"
    
    'Do filename transform
    For i = 1 To Len(filename)
        'What character?
        Select Case Mid$(filename, i, 1)
        Case "/", "\"
            'Add path separator if there isn't one and we're not at beginning
            If i <> 0 And Right$(newFilename, 1) <> "\" Then
                newFilename = newFilename & "\"
            End If
            
        Case "#", ";", "?"
            'End translation
            Exit For
            
        Case "%"
            'Next 2 chars are hex and need to be converted
            If Len(filename) - i - 1 < 2 Then
                'Not enough characters
                SendError id, "400 Bad Request"
                Exit Sub
            End If
            
            'Get chars and convert
            newFilename = newFilename & Chr$(Val("&H" & Mid$(filename, i + 1, 2)))
            
            'Advance 2 extra characters
            i = i + 2
            
        Case Else
            'Copy verbatim
            newFilename = newFilename & Mid$(filename, i, 1)
        End Select
    Next
    
    'No going up directories
    If InStr(newFilename, "\..\") <> 0 Then
        SendError id, "400 Bad Request"
        Exit Sub
    End If
    
    'Prepend htdocs path
    newFilename = HTDOCS & newFilename
    
    On Local Error Resume Next
    
    'Directory?
    If (GetAttr(newFilename) And vbDirectory) <> 0 Then
        'Only continue if there is no error
        If Err.Number = 0 Then
            'Redirect if path does not end in a \
            If Right$(newFilename, 1) <> "\" Then
                SendError id, "301 Moved Permanently" & vbCrLf & "Location: " & filename & "/"
                Exit Sub
            End If
        
            'Add index.html
            newFilename = newFilename & DEFAULT_FILE
        End If
    End If
    
    'Exists?
    GetAttr newFilename
    
    If Err.Number = 5 Then
        'Access Denied
        SendError id, "403 Forbidden"
        Exit Sub
    ElseIf Err.Number <> 0 Then
        'File not found
        SendError id, "404 File Not Found"
        Exit Sub
    End If
    
    'Get free file number
    FileNbr = FreeFile
    
    If Err.Number <> 0 Then
        'Too many open files
        SendError id, "503 Service Unavaliable"
        Exit Sub
    End If
    
    'Attempt to open file
    Open newFilename For Binary Access Read As #FileNbr
    
    'Error opening file?
    If Err.Number <> 0 Then
        'Other generic error
        SendError id, "500 Internal Server Error"
        Exit Sub
    End If
    
    On Local Error GoTo 0
    
    'Send OK response
    sockClient(id).SendData "HTTP/1.0 200 OK" & vbCrLf & vbCrLf
    
    'Any data?
    If LOF(fileNbr) = 0 Then
    	'Free file
    	Close #fileNbr
    	conns(id).FileNbr = -2
    Else
    	'Assign file to connection
    	conns(id).FileNbr = FileNbr
    End If
End Sub

Private Sub Form_Load()
    'Start listening
    sockListener.Bind PORT
    sockListener.Listen
    
    'Create array
    ReDim conns(0)
    conns(0).FreeChain = -1
    freePtr = 0
    
End Sub

Private Sub sockClient_Close(Index As Integer)
    'Close connection
    sockClient(Index).Close
    
    'Close file
    If conns(Index).FileNbr >= 0 Then
        Close conns(Index).FileNbr
    End If
    
    'Add to free chain
    conns(Index).FreeChain = freePtr
    freePtr = Index
    
End Sub

Private Sub sockClient_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim data As String
    Dim pos As Long
    Dim i As Long
    Dim requestParts As Variant
    
    Dim method As String
    Dim filename As String
    Dim httpVersion As String
    
    'Only accept data from connections not already processed
    If conns(Index).FileNbr <> -1 Then Exit Sub
    
    'Request received
    sockClient(Index).GetData data, vbString
    data = LTrim$(data)
    
    'Get request line
    pos = InStr(data, vbLf)
    
    If pos > 0 Then
        'Trim to get only the request line
        data = Left$(data, pos)
    End If
    
    'Parse line
    requestParts = Split(Trim$(data), " ")
    
    For i = 0 To UBound(requestParts)
        'Empty?
        If Len(requestParts(i)) > 0 Then
            'Fill method or filename
            If Len(method) = 0 Then
                method = requestParts(i)
            ElseIf Len(filename) = 0 Then
                filename = requestParts(i)
            Else
                httpVersion = requestParts(i)
                Exit For
            End If
        End If
    Next
    
    'Valid version?
    If Len(httpVersion) > 0 Then
        If Len(httpVersion) < 8 Then
            'Too short
            SendError Index, "505 HTTP Version Not Supported"
            Exit Sub
        ElseIf Left$(httpVersion, 7) <> "HTTP/1." Then
            'Invalid major version
            SendError Index, "505 HTTP Version Not Supported"
            Exit Sub
        End If
    End If
    
    'Valid method?
    If method <> "GET" And method <> "HEAD" Then
        'Not implemented
        SendError Index, "501 Not Implemented"
        Exit Sub
    End If
    
    'Do request
    On Error GoTo serverError
    conns(Index).FileNbr = -2
    DoGetRequest Index, filename
    On Error GoTo 0
    
    'If HEAD, clear file to be sent back
    If method = "HEAD" Then conns(Index).FileNbr = -2
    
    Exit Sub
    
serverError:
    'Error processing request
    ' Close file if needed
    If conns(Index).FileNbr <> -2 Then
        Close conns(Index).FileNbr
    End If
    
    ' Send error
    SendError Index, "500 Internal Server Error"
    
End Sub

Private Sub sockClient_SendComplete(Index As Integer)
    'Connection completed send
    
    'Get more data to send
    If conns(Index).FileNbr >= 0 Then
        'How many bytes left?
        Dim bytesToRead As Long
        bytesToRead = LOF(conns(Index).FileNbr) - Loc(conns(Index).FileNbr)
        
        'Don't continue if there's nothing left
        If bytesToRead > 0 Then
            'Cap at block size
            If bytesToRead > BLOCK_SIZE Then
                bytesToRead = BLOCK_SIZE
            End If
            
            'Read data
            Dim data() As Byte
            ReDim data(bytesToRead)
            Get #conns(Index).FileNbr, , data
            
            'Send data on
            sockClient(Index).SendData data
            
            Exit Sub
        End If
    End If
    
    'Close connection
    sockClient_Close Index
End Sub

Private Sub sockListener_ConnectionRequest(ByVal requestID As Long)
    Dim sockID As Long
    
    'Get free socket
    sockID = freePtr
    
    If sockID = -1 Then
        'Grow array first
        GrowConnections
        sockID = freePtr
    End If
    
    'Update free pointer
    freePtr = conns(sockID).FreeChain
    conns(sockID).FreeChain = -1
    
    'Accept connection and wait for request
    sockClient(sockID).Close
    sockClient(sockID).Accept requestID
    conns(sockID).FileNbr = -1
End Sub
