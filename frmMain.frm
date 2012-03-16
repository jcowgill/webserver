VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   Caption         =   "Web Server"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6375
   ScaleHeight     =   2895
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timStatsTimer 
      Interval        =   1000
      Left            =   5760
      Top             =   1680
   End
   Begin MSWinsockLib.Winsock sockListener 
      Left            =   5160
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sockClient 
      Index           =   0
      Left            =   5760
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label labKBytesDelta 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label labKBytes 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label labRequestsDelta 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label labRequests 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "KBytes / sec:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "KBytes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Requests / sec:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Requests:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Statistics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label labAddress 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6135
   End
   Begin VB.Label labRunning 
      Caption         =   "Web server running. Goto this address in a web browser."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Binary

'Port to listen on
Private Const PORT As Long = 1503

'Path to html files (must not end in a \)
Private Const HTDOCS As String = "N:\Computing\x\git\doc\git\html"

'Default html file
Private Const DEFAULT_FILE As String = "index.html"

'Server information (must not contain spaces)
Private Const SERVER_LINE As String = "WebCow/1.0"

'Size of blocks to send
Private Const BLOCK_SIZE As Long = 16384        '16 KB

'Information about a connection
Private Type ConnectionInfo
    FreeChain As Long       'Pointer to next free connection or -1 if this is in use
    FileNbr As Integer      'File number being read (-1 = request not received, -2 = nothing left to send)
End Type

'Connection list
Private conns() As ConnectionInfo

'Global buffer
Private buffer() As Byte

'Content type mappings
Private contentTypes As Collection

'Pointer to first free connection
Private freePtr As Long

'STATISTICS
Private requests As Long
Private requestsBefore As Long
Private bytes As Long
Private bytesBefore As Long

Private Sub SetupContentTypes()
    'Add extension -> content type mappings
    Set contentTypes = New Collection
    contentTypes.Add "application/java-archive", "jar"
    contentTypes.Add "application/java-vm", "class"
    contentTypes.Add "application/javascript", "js"
    contentTypes.Add "audio/midi", "mid"
    contentTypes.Add "audio/midi", "midi"
    contentTypes.Add "audio/mpeg", "mp3"
    contentTypes.Add "audio/ogg", "ogg"
    contentTypes.Add "audio/x-ms-wma", "wma"
    contentTypes.Add "audio/x-wav", "wav"
    contentTypes.Add "image/bmp", "bmp"
    contentTypes.Add "image/gif", "gif"
    contentTypes.Add "image/jpeg", "jpg"
    contentTypes.Add "image/jpeg", "jpeg"
    contentTypes.Add "image/png", "png"
    contentTypes.Add "image/tiff", "tif"
    contentTypes.Add "image/tiff", "tiff"
    contentTypes.Add "image/vnd.adobe.photoshop", "psd"
    contentTypes.Add "image/x-icon", "ico"
    contentTypes.Add "message/rfc822", "eml"
    contentTypes.Add "text/css", "css"
    contentTypes.Add "text/csv", "csv"
    contentTypes.Add "text/html", "htm"
    contentTypes.Add "text/html", "html"
    contentTypes.Add "text/plain", "txt"
    contentTypes.Add "text/plain", "text"
    contentTypes.Add "text/plain", "conf"
    contentTypes.Add "text/plain", "def"
    contentTypes.Add "text/plain", "list"
    contentTypes.Add "text/plain", "log"
    contentTypes.Add "text/x-asm", "asm"
    contentTypes.Add "text/x-asm", "s"
    contentTypes.Add "text/x-c", "c"
    contentTypes.Add "text/x-c", "cc"
    contentTypes.Add "text/x-c", "cpp"
    contentTypes.Add "text/x-c", "h"
    contentTypes.Add "text/x-c", "hh"
    contentTypes.Add "text/x-java-source", "java"
    contentTypes.Add "video/mp4", "mp4"
    contentTypes.Add "video/mpeg", "mpeg"
    contentTypes.Add "video/mpeg", "mpg"
    contentTypes.Add "video/ogg", "ogv"
    contentTypes.Add "video/x-msvideo", "avi"
End Sub

Private Function GetContentType(ByVal filename As String) As String
    Dim dotPos As Long
    Dim slashPos As Long
    
    'Get positions of special characters
    dotPos = InStrRev(filename, ".")
    slashPos = InStrRev(filename, "\")
    
    'Is there a valid extension?
    If dotPos <> 0 And dotPos > slashPos And Len(filename) > dotPos Then
        'Trim string
        filename = Trim$(Mid$(filename, dotPos + 1))
    Else
        'Use blank extension
        filename = ""
    End If
    
    'Lookup content type
    On Local Error GoTo notFound
    GetContentType = contentTypes.Item(filename)
    
    Exit Function
    
notFound:
    'No content type found, use octet stream (raw binary data)
    GetContentType = "application/octet-stream"
    
End Function

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

Private Function HttpFormatDate(ByVal dateVal As Date) As String
    'Formats a date for http
    HttpFormatDate = Format$(dateVal, "ddd, dd mmm yyyy Hh:Nn:Ss") & " GMT"
End Function

Private Sub SendData(ByVal id As Integer, ByVal data As String)
    'Send data and update byte count
    bytes = bytes + Len(data)
    sockClient(id).SendData data
End Sub

Private Sub SendHeader(ByVal id As Integer, ByVal header As String, ByVal value As String)
    'Send header to client
    SendData id, header & ": " & value & vbCrLf
End Sub

Private Sub SendResponse(ByVal id As Integer, ByVal status As String)
    'Sends the response line and main headers
    SendData id, "HTTP/1.0 " & status & vbCrLf
    SendData id, "Date: " & HttpFormatDate(Now) & vbCrLf
    SendData id, "Server: " & SERVER_LINE & vbCrLf
End Sub

Private Sub SendError(ByVal id As Integer, ByVal status As String)
    'Send error response
    SendResponse id, status
    
    'End of headers
    SendData id, vbCrLf
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
    SendResponse id, "HTTP/1.0 200 OK"
    SendHeader id, "Content-Length", LOF(FileNbr)
    SendHeader id, "Content-Type", GetContentType(newFilename)
    SendHeader id, "Last-Modified", HttpFormatDate(FileDateTime(newFilename))
    SendData id, vbCrLf
    
    'Any data?
    If LOF(FileNbr) = 0 Then
        'Free file
        Close #FileNbr
        conns(id).FileNbr = -2
    Else
        'Assign file to connection
        conns(id).FileNbr = FileNbr
    End If
End Sub

Private Sub Form_Load()
    'Setup types
    SetupContentTypes

    'Start listening
    sockListener.Bind PORT
    sockListener.Listen
    
    'Create array
    ReDim conns(0)
    conns(0).FreeChain = -1
    freePtr = 0
    
    'Set running address
    labAddress = "http://" & sockListener.LocalHostName & ":" & PORT & "/"
    
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
            ReDim buffer(bytesToRead)
            Get #conns(Index).FileNbr, , buffer
            
            'Send data on
            bytes = bytes + bytesToRead
            sockClient(Index).SendData buffer
            
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
    
    'Increment request count
    requests = requests + 1
End Sub

Private Sub timStatsTimer_Timer()
    'Update requests
    labRequests = requests
    labRequestsDelta = requests - requestsBefore
    requestsBefore = requests
    
    'Update KBytes
    labKBytes = Int(bytes / 1024)
    labKBytesDelta = Int((bytes - bytesBefore) / 1024)
    bytesBefore = bytes
End Sub
