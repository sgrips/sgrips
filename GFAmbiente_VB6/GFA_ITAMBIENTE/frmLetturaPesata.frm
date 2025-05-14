VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form frmLetturaPesata 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2160
      Top             =   2400
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   0
      Left            =   2520
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   4
      DTREnable       =   0   'False
      InputLen        =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Conferma"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1680
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Annulla"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   3120
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Text            =   "0"
      Top             =   1680
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Text            =   " "
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Peso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Badge"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "frmLetturaPesata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private pPeso As Double
Private pCodeBadge As String
Private pEsci As Boolean
Private pAnnulla As Boolean
Private comCon As New COMConnections
Private comPesa As COMConnection
Private messaggio As String
Private inlettura As Boolean

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            pAnnulla = True
        Case 1
            pCodeBadge = Trim$(Text1(0).Text)
            pPeso = Text1(1).Text
    End Select
    
    Timer1.enabled = False
    pEsci = True
    

   If Not inlettura Then Unload Me
   
End Sub

Private Sub Form_Load()

START_CHAR = "&"

    Set comPesa = comCon.Add("COM4")
    comPesa.BaudeRate = "9600"
    comPesa.CDHolding = False
    comPesa.ComID = 4
    comPesa.DataBits = 8
    comPesa.DSRHolding = False
    comPesa.HandShaking = 0
    comPesa.InBufferSize = 1024
    comPesa.ParityBit = "N"
    comPesa.StopBit = 1
    
    comPesa.message.CheckSumON = False
    comPesa.message.LineFeedON = False
    comPesa.message.STXETX_ON = False
    Label3.Caption = "PRONTO"
    If (comPesa.Init(MSComm1) = False) Then
        Label3.Caption = "Errore COM"
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
comPesa.ClearBuffer
comPesa.CloseCOM
comCon.Remove 1
START_CHAR = "$"

Set comPesa = Nothing
Set comCon = Nothing
End Sub

Private Sub Text1_Change(Index As Integer)
    If Index = 0 And Len(Trim$(Text1(0).Text)) = 12 Then
        Text1(1).enabled = True
        Timer1.enabled = True
         
    End If
End Sub

Public Sub showMe(mPeso As Double, mCodBadge As String)

    Me.Show vbModal
    If Not pAnnulla And pPeso > 0 And pCodeBadge <> vbNullString Then
        mPeso = pPeso
        mCodBadge = Left$(pCodeBadge, 12)
    End If
    
End Sub
 


Private Function letturaPesoCOM() As Double
    
    letturaPesoCOM = 0
    Do While letturaPesoCOM = 0 And Not pEsci
        inlettura = True
        Select Case comPesa.ReadMessage(messaggio)
            Case ErroreRicevuto
             Label3.Caption = "errore lettura"
            Case inlettura
             Label3.Caption = "in lettura"
            Case LetturaTerminata
            
                letturaPesoCOM = Mid$(messaggio, 3, 4) & "," & Mid$(messaggio, 7, 2)
                Label3.Caption = "peso rilevato"
    
         End Select
         DoEvents
    Loop
    
    If pEsci Then
        letturaPesoCOM = 0
         
    End If
inlettura = False
End Function


Private Sub Timer1_Timer()
    On Error GoTo erro1
    
    Timer1.enabled = False
    pPeso = letturaPesoCOM
    If pPeso > 0 Then Text1(1).Text = pPeso
    If Not pEsci Then
        Timer1.enabled = True
        Else
        Unload Me
        End If
    Exit Sub
erro1:
    
End Sub
