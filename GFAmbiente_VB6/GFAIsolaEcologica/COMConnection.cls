VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COMConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"message"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarDSRHolding As Boolean 'local copy
Private mvarCDHolding As Boolean 'local copy
Private mvarComID As String 'local copy
Private mvarBaudeRate As String 'local copy
Private mvarHandshaking As String 'local copy
Private mvarParityBit As String 'local copy
Private mvarDataBits As String 'local copy
Private mvarStopBit As String 'local copy
Private mvarStartChar As String 'local copy
Private mvarEndChar As String
Private mvarInBufferSize As Long 'local copy
Private mvarkey As String 'local copy
'local variable(s) to hold property value(s)
Private mvarCOMControl As MSComm   'local copy
'local variable(s) to hold property value(s)
Private mvarIndex As Long 'local copy
Private mvarOldState As ConnectionState  'local copy
Public Errore As Long
Public DesErrore As String

Private mvarSTX As String
Private mvarETX As String
Private CR As String
Private LF As String
 
Public Enum StatoLetturaCOM
    InLettura
    ErroreRicevuto
    LetturaTerminata
End Enum


Public Enum ConnectionState
    NotConnect = 0
    Connect = 1
    OnError = 2
    NotChange = 3
End Enum

Public Enum ConnectionProtocol
    STXETX = 0
    NOProtocol = 1
End Enum

Private mvarMessage As message

Public Sub CloseCOM()
    If Not COMControl Is Nothing Then
        If COMControl.PortOpen = True Then COMControl.PortOpen = False
        If COMControl.Index > 0 Then Unload COMControl
        Set COMControl = Nothing
    End If
End Sub

Public Sub ClearBuffer()
    Dim a As String
    Do While COMControl.InBufferCount > 0
        a = COMControl.Input
        DoEvents
    Loop
End Sub

Friend Sub ClearError()
    Errore = 0
    DesErrore = vbNullString
End Sub

Public Property Get message() As message
    Set message = mvarMessage
End Property

Public Property Set message(vData As message)
    Set mvarMessage = vData
End Property



Public Function GetNextChar(Optional InitString As String = "") As String
    Static i As Long
    Static s As String
    
    If InitString <> "" Then
        i = 1
        s = InitString
    Else
        If i <= Len(s) Then
            GetNextChar = Mid$(s, i, 1)
            i = i + 1
        Else
            GetNextChar = ""
        End If
    End If
    

    
End Function


Public Function ReadMessage(message As String) As StatoLetturaCOM
    'Lettura dal buffer del MSG
    Dim Esci As Boolean
    Dim sChar As String
    Dim bErr As Boolean
    Dim s As String
    Dim i As Integer
    ReadMessage = InLettura
    Esci = False
    
    If Me.message.message = vbNullString Then
        bErr = True
    Else
        bErr = False
    End If
    
    i = 1
    With COMControl
        
        Do While ((.InBufferCount > 0 Or mDEBUG) And Not Esci And IsConnect)
            
            If bErr Then
                gTimerLettura = Timer
                Me.message.message = vbNullString
                Esci = False: Me.message.STX = False: Me.message.ETX = False
                bErr = False: Me.message.LF = False: Me.message.CheckSum = 0
                Me.message.START_CHAR = False
                Me.message.chk1 = 0: Me.message.chk2 = 0
            End If
            
            If .InBufferCount > 0 Or mDEBUG Then
                If mDEBUG Then
                    sChar = GetNextChar
                    If sChar = vbNullString Then Exit Do
                Else
                    sChar = .Input
                End If
                
                
                'termine lista
                If sChar = Chr$(&H4&) Then
                    ReadMessage = LetturaTerminata
                    message = sChar
                    Me.message.message = vbNullString
                    Exit Do
                End If
                
                'sChar = Mid$("$1507/01/200815:39:35    74    200101       CARTA CARTONE   1010003400010                                                                       11.60 kg       0 kg        0 kg     0.00 kg        0 kg        0 kg    11.60 kg        0 kg        0 kg24" & vbCr, i, 1)
                'sChar = Mid$("$1110/12/200709:52:44     21111111111aaaaaaaaaaaaaaaaaaaa2222222222bbbbbbbbbbbbbbbbbbbb3333333333ssssssssssssssssssss         0                                                                                                                                                                                       0          10110 kg      237 kg     9873 kg31" & vbCr, i, 1)
                'i = i + 1
                If sChar = StartChar Then
                    Me.message.START_CHAR = True
                    Me.message.message = sChar
                    If Me.message.CheckSumON Then
                        Me.message.CheckSum = Me.message.CheckSum Xor Me.message.chk1
                        Me.message.chk1 = Me.message.chk2: Me.message.chk2 = Asc(sChar)
                    End If
                    
                ElseIf sChar = STX Then
                    Me.message.STX = True
                ElseIf sChar = ETX Then
                    If Me.message.STX And Not Me.message.ETX Then
                        Me.message.ETX = True
                        If Me.message.CarriageReturnON = False Then
                            Esci = True
                        End If
                        
                    Else
                        bErr = True
                    End If
                ElseIf sChar = CR Then
                    If Me.message.STX And Me.message.ETX And Me.message.START_CHAR Then
                        Me.message.CR = True
                        If Me.message.LineFeedON = False Then
                            Esci = True
                        End If
                    Else
                        bErr = True
                    End If
                ElseIf sChar = LF Then
                    If Me.message.STX And Me.message.ETX And Me.message.CR And Not Me.message.LF Then
                        Me.message.LF = True
                        Esci = True
                    Else
                        bErr = True
                    End If
                
                Else
                    If Not Me.message.STX And Not Me.message.START_CHAR Then
                        bErr = True
                    ElseIf (Not Me.message.ETX Or Not Me.message.STXETX_ON) And (Asc(sChar) >= Asc(" ") And Asc(sChar) <= Asc("~")) Then
                            Me.message.message = Me.message.message & sChar
                            
                            If Me.message.CheckSumON Then
                                Me.message.CheckSum = Me.message.CheckSum Xor Me.message.chk1
                                Me.message.chk1 = Me.message.chk2: Me.message.chk2 = Asc(sChar)
                            End If
                        Else
                            bErr = True
                    End If
                End If
                
                If (Esci And Me.message.CheckSumON And Not Right$("0" & Hex(Me.message.CheckSum), 2) = Chr(Me.message.chk1) & Chr(Me.message.chk2)) Or bErr = True Then
                    Esci = True: bErr = True
                    ReadMessage = ErroreRicevuto
                    If Me.message.message <> vbNullString Then
                        ScriviLog TPL_ERRORE, App.Title, Me.message.message, "ReadMessage"
                    End If
                    Me.message.message = vbNullString
                ElseIf Esci Then
                        If Me.message.CheckSumON Then Me.message.message = left$(Me.message.message, Len(Me.message.message) - 2)
                        
                        ReadMessage = LetturaTerminata
                        message = Me.message.message
                        Me.message.message = vbNullString
'                       Debug.Print "Lettura Pistola: " & Format$(Timer - gTimerLettura, "0.001")
                End If
            End If
        Loop
    End With
    


End Function



Public Function SendMessage(ByVal message As String, Optional Protocol As ConnectionProtocol = NOProtocol, Optional ByVal AddCRC As Boolean = False) As Boolean
    Dim crc As Long, i As Long
    Dim num As Long
    
    On Error GoTo SendERR
    SendMessage = False
    
    'Add CRC char
    If AddCRC Then
        crc = 0
        For i = 1 To Len(message)
            crc = crc Xor Asc(Mid$(message, i, 1))
        Next
        
        message = message & Right$("0" & Hex(crc), 2)
    End If
    
    'test type of protocol
    Select Case Protocol
        Case STXETX
            'STX ETX protocol
            message = STX & message & ETX & CR
    End Select
    
    With COMControl
        .Output = message
    End With
    SendMessage = True
    
    Exit Function
SendERR:
End Function


Friend Property Let Index(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Index = 5
    mvarIndex = vData
End Property


Public Property Get Index() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Index
    Index = mvarIndex
End Property


Public Property Set COMControl(ByVal vData As MSComm)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.COMControl = Form1
    Set mvarCOMControl = vData
End Property


Public Property Get COMControl() As MSComm
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.COMControl
    Set COMControl = mvarCOMControl
End Property


Friend Property Let StartChar(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Key = 5
    mvarStartChar = vData
End Property


Public Property Get StartChar() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Key
    StartChar = mvarStartChar
End Property



Friend Property Let key(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Key = 5
    mvarkey = vData
End Property


Public Property Get key() As String
Attribute key.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Key
    key = mvarkey
End Property


Public Function Init(MSCCom As Object) As Boolean
    
    ClearError
    On Error GoTo InitERR
    Init = False
    'test if all parameters are fill
    If BaudeRate = vbNullString Or ComID = vbNullString Or _
       DataBits = vbNullString Or InBufferSize = 0 Or _
       ParityBit = vbNullString Or StopBit = vbNullString Then
           Exit Function
    End If
        
    If Index > 0 Then
        Load MSCCom(MSCCom.UBound + 1)
    End If
        
    Set COMControl = MSCCom(MSCCom.UBound)
        
    With COMControl
        'Setting Port
        .CommPort = ComID
        .Settings = BaudeRate & "," & left$(ParityBit, 1) & "," & DataBits & "," & StopBit
        .InBufferSize = 10240
        .HandShaking = HandShaking
        If .PortOpen = False Then
            .PortOpen = True
            .InputLen = 1
            .InputMode = comInputModeText
        End If
    End With
    
    Init = True
    Exit Function
InitERR:
    Errore = Err
    DesErrore = Error$
End Function

Public Function IsConnect(Optional st As Variant) As Boolean
    'Test if the cable is connect to the PC
    IsConnect = False
    
    If Me.COMControl.PortOpen = False Then
        If Not IsMissing(st) Then
            st = NotConnect
            If mvarOldState = st Then
                st = NotChange
            Else
                mvarOldState = st
            End If
        End If
        Exit Function
    End If
    
    If Me.CDHolding = True And Me.DSRHolding = True Then
        With COMControl
           If .DSRHolding And .CDHolding Then
                IsConnect = True
                If Not IsMissing(st) Then
                    st = Connect
                    If mvarOldState = st Then
                        st = NotChange
                    Else
                        mvarOldState = st
                    End If
                End If
           Else
                If Not IsMissing(st) Then
                    st = NotConnect
                    If mvarOldState = st Then
                        st = NotChange
                    Else
                        mvarOldState = st
                    End If
                End If
           End If
        End With
    Else
        IsConnect = True
        If Not IsMissing(st) Then
            st = Connect
            If mvarOldState = st Then
                st = NotChange
            Else
                mvarOldState = st
            End If
        End If
    End If


End Function

Public Property Let InBufferSize(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.InBufferSize = 5
    mvarInBufferSize = vData
End Property


Public Property Get InBufferSize() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.InBufferSize
    InBufferSize = mvarInBufferSize
End Property


Public Property Let HandShaking(ByVal vHandShaking As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StopBit = 5
    mvarHandshaking = vHandShaking
End Property


Public Property Get HandShaking() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StopBit
    HandShaking = mvarHandshaking
End Property



Public Property Let StopBit(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StopBit = 5
    mvarStopBit = vData
End Property


Public Property Get STX() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StopBit
    STX = mvarSTX
End Property

Public Property Get ETX() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StopBit
    ETX = mvarETX
End Property


Public Property Get StopBit() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StopBit
    StopBit = mvarStopBit
End Property


Public Property Let STX(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DataBits = 5


        Select Case vData
        Case "N"
            mvarSTX = vbNullString
        Case "STX"
            mvarSTX = Chr(&H2&)
        Case "ETX"
            mvarSTX = Chr(&H3&)
        Case "ESC"
            mvarSTX = Chr(27)
        Case Else
            mvarSTX = vData
    End Select
End Property

Public Property Let ETX(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DataBits = 5
        Select Case vData
        Case "N"
            mvarETX = vbNullString
        Case "STX"
            mvarETX = Chr(&H2&)
        Case "ETX"
            mvarETX = Chr(&H3&)
        Case "ESC"
            mvarETX = Chr(27)
        Case Else
            mvarETX = vData
    End Select
End Property

Public Property Let DataBits(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DataBits = 5
    mvarDataBits = vData
End Property


Public Property Get DataBits() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DataBits
    DataBits = mvarDataBits
End Property



Public Property Let ParityBit(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ParityBit = 5
    mvarParityBit = vData
End Property


Public Property Get ParityBit() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ParityBit
    ParityBit = mvarParityBit
End Property



Public Property Let BaudeRate(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BaudeRate = 5
    mvarBaudeRate = vData
End Property


Public Property Get BaudeRate() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BaudeRate
    BaudeRate = mvarBaudeRate
End Property



Public Property Let ComID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ComID = 5
    mvarComID = vData
End Property


Public Property Get ComID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ComID
    ComID = mvarComID
End Property



Public Property Let CDHolding(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CDHolding = 5
    mvarCDHolding = vData
End Property


Public Property Get CDHolding() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CDHolding
    CDHolding = mvarCDHolding
End Property



Public Property Let DSRHolding(ByVal vData As Boolean)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DSRHolding = Form1
    mvarDSRHolding = vData
End Property


Public Property Get DSRHolding() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DSRHolding
    DSRHolding = mvarDSRHolding
End Property

Private Sub Class_Initialize()
    'Variables Inizilize
    BaudeRate = vbNullString
    CDHolding = False
    Set COMControl = Nothing
    ComID = vbNullString
    DataBits = vbNullString
    DSRHolding = False
    InBufferSize = 0
    ParityBit = vbNullString
    StopBit = vbNullString
    mvarOldState = NotChange
    
    STX = "STX"
    ETX = "ETX"
    CR = vbCr
    LF = vbLf
    
    'create the mMessage object when the COMConnection class is created
    Set mvarMessage = New message
End Sub

Private Sub Class_Terminate()
  Set mvarMessage = Nothing
End Sub
