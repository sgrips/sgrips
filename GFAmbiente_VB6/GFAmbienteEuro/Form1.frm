VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form FrmTestFSC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test Collegamento con FSC"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkProtocollo 
      Caption         =   "STX/ETX"
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "&Stop"
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CheckBox ChkRicezione 
      Caption         =   "Loop Ricezione"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   240
      Width           =   1455
   End
   Begin VB.CheckBox ChkInvio 
      Caption         =   "Loop Invio"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCOM 
      Caption         =   "Sort Message"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton CmdChiudi 
      Caption         =   "&Chiudi"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ricezione"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1440
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Invio"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   3375
   End
   Begin MSCommLib.MSComm MSCPortaSORT 
      Left            =   210
      Top             =   1890
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327681
      DTREnable       =   -1  'True
      BaudRate        =   19200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   4935
   End
End
Attribute VB_Name = "FrmTestFSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bLoopInvio As Boolean
Dim bLoopRicezione As Boolean
Dim bProtocollo As Boolean


Private Sub ChkInvio_Click()
    bLoopInvio = ChkInvio.Value = 1
End Sub

Private Sub chkProtocollo_Click()
    bProtocollo = chkProtocollo.Value = 1
End Sub

Private Sub ChkRicezione_Click()
    bLoopRicezione = ChkRicezione.Value = 1
End Sub

Private Sub CmdChiudi_Click()

If MSCPortaSORT.PortOpen = True Then MSCPortaSORT.PortOpen = False
Unload Me

End Sub

Private Sub cmdCOM_Click()
    cmdCOM.Caption = IIf(cmdCOM.Caption = "Sort Message", "Request Message", "Sort Message")
    Label1.Caption = "Testing COM" & IIf(cmdCOM.Caption = "Sort Message", FrmMonitoraggio.MSCPortaCom.CommPort, MSCPortaSORT.CommPort)
End Sub

Private Sub CmdStop_Click()
    bLoopInvio = False
    bLoopRicezione = False
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim i As Long
    Dim sChar As String
    Dim smsg As String
    
    If bLoopInvio Or bLoopRicezione Then
        
        ChkInvio.Enabled = False
        ChkRicezione.Enabled = False
        chkProtocollo.Enabled = False
        Command1.Item(0).Enabled = False
        Command1.Item(1).Enabled = False
        cmdCOM.Enabled = False
        cmdChiudi.Enabled = False
        CmdStop.Enabled = True
    End If
    
    Select Case Index
    
        Case 0
            Do
                smsg = Text1(0).Text
                If bProtocollo Then
                    InvioMsgSTX_ETX smsg
                End If
                If cmdCOM.Caption = "Sort Message" Then
                    FrmMonitoraggio.MSCPortaCom.Output = smsg
                Else
                    MSCPortaSORT.Output = smsg
                End If
                DoEvents
            Loop While bLoopInvio
        Case 1
            Text1(1) = vbNullString
            sChar = vbNullString
            Do While bLoopRicezione
                If cmdCOM.Caption = "Sort Message" Then
                    For i = 1 To FrmMonitoraggio.MSCPortaCom.InBufferCount
                        sChar = sChar & FrmMonitoraggio.MSCPortaCom.Input
                    Next
                Else
                    For i = 1 To MSCPortaSORT.InBufferCount
                        sChar = sChar & MSCPortaSORT.Input
                    Next
                End If
                Text1(1) = sChar
                DoEvents
            Loop
    End Select
    
    Command1.Item(0).Enabled = True
    Command1.Item(1).Enabled = True
    CmdStop.Enabled = True
    cmdCOM.Enabled = True
    cmdChiudi.Enabled = True
    bLoopInvio = ChkInvio.Value = 1
    bLoopRicezione = ChkRicezione.Value = 1
    ChkInvio.Enabled = True
    ChkRicezione.Enabled = True
    chkProtocollo.Enabled = True
    
End Sub

Function InitCom(objCOM As MSComm) As Boolean
    'Inizializzazione porta
    
    On Error GoTo InitComERR
    InitCom = False
    With objCOM
        If .PortOpen = False Then
            .PortOpen = True
            .InputLen = 1
            .InputMode = comInputModeText
            InitCom = True
        End If
    End With
    Exit Function
InitComERR:
    If err = comPortAlreadyOpen Then
        MsgBox "Attenzione il test delle porte può essere effettuato solo se" + Chr$(13) + "non si è precedentemente avviato il PosiSorter." + Chr$(13) + "Uscire dal programma e provare a riavviare il test", vbInformation + vbOKOnly, App.Title
    Else
        MsgBox Error$
    End If
End Function

Sub SettingCom(objCOM As MSComm)
    'Settaggio porta
    With objCOM
        .CommPort = gCOM(1)
        .Settings = gBaudeRateCOM(1) & "," & Left$(gBitParita(1), 1) & "," & gDataBits(1) & "," & gBitDiStop(1)
        .InBufferSize = 10240
    End With
End Sub

Private Sub Form_Activate()
    Label1.Caption = "Testing COM" & IIf(cmdCOM.Caption = "Sort Message", FrmMonitoraggio.MSCPortaCom.CommPort, MSCPortaSORT.CommPort)

End Sub

Private Sub Form_Load()

    SettingCom MSCPortaSORT
    If Not InitCom(MSCPortaSORT) Then

        Unload Me
    Else
        Label1.ForeColor = vbRed
        bLoopInvio = ChkInvio.Value = 1
        bLoopRicezione = ChkRicezione.Value = 1
        CmdStop.Enabled = bLoopRicezione Or bLoopInvio
        bProtocollo = chkProtocollo.Value = 1
    End If
End Sub

Public Static Sub InvioMsgSTX_ETX(smsg As String)
    
    Dim crc As Long, i As Long
    Dim num As Long
   
    crc = 0
    For i = 1 To Len(smsg)
        crc = crc Xor Asc(Mid$(smsg, i, 1))
    Next
    
    smsg = STX & smsg & Right$("0" & Hex(crc), 2) & ETX & CR
End Sub

