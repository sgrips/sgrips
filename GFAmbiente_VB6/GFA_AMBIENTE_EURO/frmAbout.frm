VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "       Informazioni su MiaApplicazione"
   ClientHeight    =   4695
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6030
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240.572
   ScaleMode       =   0  'User
   ScaleWidth      =   5662.482
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegister 
      Cancel          =   -1  'True
      Caption         =   "Register"
      Default         =   -1  'True
      Height          =   345
      Left            =   4560
      TabIndex        =   9
      Top             =   3720
      Width           =   1260
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   4920
      ScaleHeight     =   1575
      ScaleWidth      =   735
      TabIndex        =   7
      Top             =   1440
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1452
      Left            =   120
      ScaleHeight     =   1455
      ScaleWidth      =   4215
      TabIndex        =   6
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4800
      Top             =   480
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   4560
      TabIndex        =   0
      Top             =   3240
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4560
      TabIndex        =   1
      Top             =   4200
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      TabIndex        =   14
      Top             =   720
      Width           =   540
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CopyRight @ 1998-2008 Antonio Sgrignoli"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1455
      TabIndex        =   2
      Top             =   1245
      Width           =   3585
   End
   Begin VB.Label lblMail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3840
      TabIndex        =   13
      Top             =   720
      Width           =   555
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   12
      Top             =   720
      Width           =   555
   End
   Begin VB.Label lblcompany 
      BackStyle       =   0  'Transparent
      Caption         =   "Versione"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label lblDayCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "La versione demo scadrà tra"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   468
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   3732
   End
   Begin VB.Label lblEuro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "€"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   5520
      TabIndex        =   8
      Top             =   0
      Width           =   285
   End
   Begin VB.Image picIcon 
      Height          =   1005
      Left            =   105
      Picture         =   "frmAbout.frx":000C
      Stretch         =   -1  'True
      Top             =   150
      Width           =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5070.879
      Y1              =   2184.539
      Y2              =   2184.539
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Titolo applicazione"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   1836
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   5070.879
      Y1              =   2194.892
      Y2              =   2194.892
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versione"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   5
      Top             =   1043
      Width           =   750
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":67B0
      ForeColor       =   &H00000000&
      Height          =   1065
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   4215
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Opzioni di protezione per la chiave del registro di configurazione
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Chiavi principali del registro di configurazione
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Stringa Unicode che termina con un carattere Null
Const REG_DWORD = 4                      ' Numero a 32 bit

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private p1 As StdPicture
Private p2 As StdPicture
Public Function MustRegister() As Boolean
    
    
    
'    If Not objActive.RegisteredUser And gGiorniScandenzaDemo = 0 Then
'        'Registrazione Automatica
'        objActive.Register frmGFAKey.getKeyRegisterGFA(objActive.SoftwareCode)
'        Unload frmGFAKey
'    End If
 '   MustRegister = Not objActive.RegisteredUser And gGiorniScandenzaDemo - objActive.UsedDays <= 0
    Me.lblDayCaption = "Versione scaduta contattare il 051-726291 per Attivare il Programma"
    
    
    MustRegister = False
End Function

Private Sub cmdRegister_Click()
  '  frmRegister.showME
End Sub

Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub




Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblURL.FontUnderline = False
    lblURL.ForeColor = vbBlack
    
    lblMail.FontUnderline = False
    lblMail.ForeColor = vbBlack

End Sub


Private Sub lblMail_Click()
    Call ShellExecute(0&, vbNullString, "mailto:" & lblMail.Caption, vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub lblMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With lblMail
        .ForeColor = vbBlue
        .Font.Underline = True
        .MousePointer = vbIconPointer
    End With
End Sub

Private Sub lblURL_Click()
    Call ShellExecute(0&, vbNullString, lblURL.Caption, vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub lblURL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With lblURL
        .ForeColor = vbBlue
        .Font.Underline = True
        .MousePointer = vbIconPointer
    End With
End Sub
Private Sub Form_Load()
    Timer1.Enabled = False
    Me.Caption = "   Informazioni su " & App.Title
    lblVersion.Caption = "Versione " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDescription = App.LegalCopyright
    lblcompany = App.CompanyName & " " & App.Comments
    Set p1 = LoadPicture(App.Path & "\GFAmbienti.gif")
    Set p2 = LoadPicture(App.Path & "\image9.gif")
    Timer1.Enabled = True
    lblURL.Caption = "www.gfambiente.it"
    lblMail.Caption = "info@gfambiente.it"
'    If Not objActive.RegisteredUser Then
'        lblDayCaption = lblDayCaption & " " & gGiorniScandenzaDemo - objActive.UsedDays & " Giorni"
'    Else
        lblDayCaption.Visible = False
'    End If
    cmdRegister.Enabled = lblDayCaption.Visible = True
    'objActive.Reset
    
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Tenta di recuperare dal registro di configurazione il percorso e il nome
    ' del programma che consente di visualizzare le informazioni sul sistema
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Tenta di recuperare dal registro di configurazione solo il percorso
    ' del programma che consente di visualizzare le informazioni sul sistema
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Convalida l'esistenza di una versione a 32 bit del file conosciuta
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Errore. Il file non è stato trovato.
        Else
            GoTo SysInfoErr
        End If
    ' Errore. La chiave del registro di configurazione non è stata trovata.
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "Le informazioni sul sistema non sono attualmente disponibili.", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Contatore per il ciclo
    Dim rc As Long                                          ' Codice restituito
    Dim hKey As Long                                        ' Handle a una chiave del registro di configurazione aperta
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Tipo di dati di una chiave del registro di configurazione
    Dim tmpVal As String                                    ' Posizione di memorizzazione temporanea del valore di una chiave del registro di configurazione
    Dim KeyValSize As Long                                  ' Dimensioni della variabile della chiave del registro di configurazione
    '------------------------------------------------------------
    ' Apre una chiave del registro di configurazione in una chiave principale {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Apre la chiave del registro di configurazione
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Gestione degli errori
    
    tmpVal = String$(1024, 0)                             ' Assegna spazio alla variabile
    KeyValSize = 1024                                       ' Specifica le dimensioni della variabile
    
    '------------------------------------------------------------
    ' Recupera il valore della chiave del registro di configurazione
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Recupera/crea il valore della chiave
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Gestione degli errori
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' In Win95 viene aggiunta una stringa che termina con un carattere Null
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' È stato trovato un carattere Null, che viene estratto dalla stringa
    Else                                                    ' In WinNT non viene aggiunto un carattere Null al termine della stringa
        tmpVal = Left(tmpVal, KeyValSize)                   ' Non è stato trovato nessun carattere Null, pertanto estrae solo la stringa
    End If
    '------------------------------------------------------------
    ' Determina il tipo del valore della chiave per la conversione
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Cerca i tipi di dati
    Case REG_SZ                                             ' Tipo di dati String per la chiave del registro di configurazione
        KeyVal = tmpVal                                     ' Copia il valore String
    Case REG_DWORD                                          ' Tipo di dati Double Word per la chiave del registro di configurazione
        For i = Len(tmpVal) To 1 Step -1                    ' Converte ogni bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Crea il valore carattere per carattere
        Next
        KeyVal = Format$(SEPARATORE & "h" + KeyVal)                     ' Converte Double Word in String
    End Select
    
    GetKeyValue = True                                      ' Restituisce un valore che indica che l'operazione è riuscita
    rc = RegCloseKey(hKey)                                  ' Chiude la chiave del registro di configurazione
    Exit Function                                           ' Esce dalla routine
    
GetKeyError:      ' Reimposta i dati se viene generato un errore
    KeyVal = ""                                             ' Imposta su una stringa vuota il valore restituito
    GetKeyValue = False                                     ' Restituisce un valore che indica che l'operazione non è riuscita
    rc = RegCloseKey(hKey)                                  ' Chiude la chiave del registro di configurazione
End Function


Private Sub Timer1_Timer()
    
    Picture1 = LoadPicture()
    Picture2 = LoadPicture()
    
    
    If Picture1.Width - 50 < 0 Then
        Picture1.Width = Me.ScaleWidth
        Picture2.Left = Picture1.Width + Picture1.Left
        Picture2.Width = Me.ScaleWidth - Picture1.Width + 5
        Sleep 2000
    Else
        Picture1.Width = Picture1.Width - 50
        Picture2.Left = Picture1.Left + Picture1.Width
        Picture2.Width = Me.ScaleWidth - Picture1.Width + 5
    End If

    Picture1.PaintPicture p1, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, , , , , vbSrcAnd
    Picture2.PaintPicture p2, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, , , , , vbSrcAnd
End Sub
