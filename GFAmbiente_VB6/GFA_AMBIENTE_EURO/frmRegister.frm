VERSION 5.00
Object = "{BE28590A-189B-4760-8B3D-9B626E8A790C}#1.0#0"; "activelock1884.ocx"
Begin VB.Form frmRegister 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrazione Prodotto"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSoftwareCode 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   195
      Width           =   855
   End
   Begin VB.TextBox txtRegisteredUser 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Conferma"
      Height          =   1005
      Left            =   3720
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtUsedDays 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtCounter 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtLiberationKey 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   1800
      Width           =   1815
   End
   Begin activelock1884.ActiveLock objActive 
      Left            =   4080
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   820
      SoftwareName    =   "Bollettazione"
      SoftwarePassword=   "GFAmbiente"
      LiberationKeyLength=   16
      SoftwareCodeLength=   6
      LockToHardDrive =   0   'False
      LockToWindowsSerial=   -1  'True
      LockToRandomNumber=   -1  'True
      LockToComputerName=   -1  'True
      LockToMACAddress=   0   'False
      UseDataLock     =   0   'False
      RegistryPath    =   "ActiveLock"
      RegistryKey     =   "GFAmbiente"
      RegistryName    =   "ProductReg"
      RegistryHive    =   "HKLM"
      LockToCustomString=   ""
      HashAlgorithm   =   0
      RegCounterKey   =   "Counter"
      RegLiberationKey=   "LiberationKey"
      RegLastRunDateKey=   "LastRunDate"
      RegInitialRunDateKey=   "InitialRunDate"
      RegRandomKey    =   "RandomKey"
      EncKey          =   "Default"
      RegEncKey       =   -1  'True
   End
   Begin VB.Label lbldi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "di"
      Height          =   195
      Left            =   2760
      TabIndex        =   11
      Top             =   1125
      Width           =   120
   End
   Begin VB.Image picIcon 
      Height          =   1005
      Left            =   0
      Picture         =   "frmRegister.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1200
   End
   Begin VB.Label LabelSoftwareCode 
      AutoSize        =   -1  'True
      Caption         =   "Codice Software"
      Height          =   195
      Left            =   1320
      TabIndex        =   10
      Top             =   240
      Width           =   1170
   End
   Begin VB.Label LabelRegisteredUser 
      AutoSize        =   -1  'True
      Caption         =   "Utente Registrato"
      Height          =   195
      Left            =   1320
      TabIndex        =   9
      Top             =   645
      Width           =   1245
   End
   Begin VB.Label LabelUsedDays 
      AutoSize        =   -1  'True
      Caption         =   "Giorni di utilizzo"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1125
      Width           =   1080
   End
   Begin VB.Label LabelCounter 
      AutoSize        =   -1  'True
      Caption         =   "Tentativi"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label LabelLiberationKey 
      AutoSize        =   -1  'True
      Caption         =   "Chiave GFAMBIENTE"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1845
      Width           =   1575
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pCounter As Integer
Private pCounterINI As Integer
Private Sub cmdOK_Click()
    ' Set the liberation keys
    ' (This method will raise the ActiveLock1_Registration event)
    pCounter = objActive.Counter - pCounterINI
    objActive.Register txtLiberationKey

End Sub

Private Sub Form_Load()
    pCounter = 0
    UpdateControls
    pCounter = objActive.Counter
    pCounterINI = pCounter
End Sub

Private Sub objActive_Registration(WasSuccessful As Boolean)
    If WasSuccessful Then
        MsgBox "Grazie di avere registrato ""Calcolo Tariffa by GFAmbiente"" !", vbInformation, "Registrazione - GF Ambiente"
        Unload Me
        Unload frmAbout
    Else
        If pCounter > 3 Then
            MsgBox "Mi spiace ma è stato superato il numero massimo di tentativi riprovare più tardi", vbExclamation, "Registrazione - GF Ambiente"
            
            Unload Me
            Unload frmAbout
            If Not gUnloadMdi Then
                Unload MDIMain
            End If
            
        Else
            MsgBox "Mi spiace ma la chiave inserita non è valida!", vbExclamation, "Registrazione - GF Ambiente"
            UpdateControls
        End If
    End If
End Sub

Private Sub UpdateControls()
    
    txtSoftwareCode = objActive.SoftwareCode
    txtRegisteredUser = objActive.RegisteredUser
    txtUsedDays = objActive.UsedDays
    lbldi.Caption = "di " & gGiorniScandenzaDemo
    txtCounter = pCounter
End Sub

Public Sub showME()
    Me.Show vbModal
End Sub

