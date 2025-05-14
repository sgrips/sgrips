VERSION 5.00
Begin VB.Form frmImpEsp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impostazioni Esportazione"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   9435
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "HTTP"
      Height          =   1095
      Left            =   120
      TabIndex        =   58
      Top             =   6840
      Width           =   9255
      Begin VB.TextBox URLWebServiceTxt 
         Height          =   285
         Left            =   120
         TabIndex        =   60
         Top             =   480
         Width           =   5175
      End
      Begin VB.TextBox URLTokenSicurezzaTxt 
         Height          =   285
         Left            =   5400
         TabIndex        =   59
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "URL"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Token Sicurezza"
         Height          =   255
         Index           =   4
         Left            =   5400
         TabIndex        =   61
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame FTPFrame 
      Caption         =   "FTP"
      Height          =   1095
      Left            =   120
      TabIndex        =   48
      Top             =   5640
      Width           =   9255
      Begin VB.CheckBox FTPCheck 
         Caption         =   "Attiva"
         Height          =   255
         Left            =   8400
         TabIndex        =   55
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox FTPPasswordTxT 
         Height          =   285
         Left            =   5520
         TabIndex        =   54
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox FTPUserTxt 
         Height          =   285
         Left            =   3480
         TabIndex        =   52
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox FTPURLTxT 
         Height          =   285
         Left            =   120
         TabIndex        =   50
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   255
         Index           =   2
         Left            =   5520
         TabIndex        =   53
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "User"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   51
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "URL"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Accessi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   5
      Left            =   120
      TabIndex        =   40
      Top             =   2760
      Width           =   6255
      Begin VB.ComboBox TipoInvioAccessiCMB 
         Height          =   315
         ItemData        =   "FrmImportazioneEsportazione.frx":0000
         Left            =   4320
         List            =   "FrmImportazioneEsportazione.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   600
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   "Formato File / Separatore"
         Height          =   615
         Index           =   5
         Left            =   1440
         TabIndex        =   42
         Top             =   360
         Width           =   2775
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   1920
            MaxLength       =   1
            TabIndex        =   46
            Text            =   "Text1"
            Top             =   240
            Width           =   615
         End
         Begin VB.PictureBox Picture6 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   1815
            TabIndex        =   43
            Top             =   240
            Width           =   1815
            Begin VB.OptionButton Option6 
               Caption         =   "TXT"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   45
               Top             =   0
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.OptionButton Option6 
               Caption         =   "CSV"
               Height          =   255
               Index           =   1
               Left            =   720
               TabIndex        =   44
               Top             =   0
               Width           =   615
            End
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   240
         MaxLength       =   4
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Modalità Trasmissione"
         Height          =   255
         Left            =   4320
         TabIndex        =   57
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Prefisso File"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   47
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Utenti"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   4
      Left            =   120
      TabIndex        =   23
      Top             =   1440
      Width           =   4575
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   240
         MaxLength       =   4
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   600
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Caption         =   "Formato File / Separatore"
         Height          =   615
         Index           =   4
         Left            =   1440
         TabIndex        =   24
         Top             =   360
         Width           =   2775
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   1815
            TabIndex        =   31
            Top             =   240
            Width           =   1815
            Begin VB.OptionButton Option5 
               Caption         =   "CSV"
               Height          =   255
               Index           =   1
               Left            =   720
               TabIndex        =   33
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton Option5 
               Caption         =   "TXT"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   32
               Top             =   0
               Value           =   -1  'True
               Width           =   615
            End
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   1920
            MaxLength       =   1
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Prefisso File"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Codici CER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   3
      Left            =   4800
      TabIndex        =   18
      Top             =   1440
      Width           =   4575
      Begin VB.Frame Frame1 
         Caption         =   "Formato File / Separatore"
         Height          =   615
         Index           =   3
         Left            =   1440
         TabIndex        =   20
         Top             =   360
         Width           =   2775
         Begin VB.PictureBox Picture5 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   1695
            TabIndex        =   37
            Top             =   240
            Width           =   1695
            Begin VB.OptionButton Option4 
               Caption         =   "TXT"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   39
               Top             =   0
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.OptionButton Option4 
               Caption         =   "CSV"
               Height          =   255
               Index           =   1
               Left            =   720
               TabIndex        =   38
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   1920
            MaxLength       =   1
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   240
         MaxLength       =   4
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Prefisso File"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salva"
      Height          =   735
      Index           =   1
      Left            =   6480
      TabIndex        =   17
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Movimenti"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   4575
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   240
         MaxLength       =   4
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   600
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Caption         =   "Formato File / Separatore"
         Height          =   615
         Index           =   2
         Left            =   1440
         TabIndex        =   13
         Top             =   360
         Width           =   2775
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   1575
            TabIndex        =   34
            Top             =   240
            Width           =   1575
            Begin VB.OptionButton Option3 
               Caption         =   "CSV"
               Height          =   255
               Index           =   1
               Left            =   720
               TabIndex        =   36
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton Option3 
               Caption         =   "TXT"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   35
               Top             =   0
               Value           =   -1  'True
               Width           =   615
            End
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   1920
            MaxLength       =   1
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Prefisso File"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fatture"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   4800
      TabIndex        =   7
      Top             =   120
      Width           =   4575
      Begin VB.Frame Frame1 
         Caption         =   "Formato File / Separatore"
         Height          =   615
         Index           =   1
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   2775
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   1695
            TabIndex        =   28
            Top             =   240
            Width           =   1695
            Begin VB.OptionButton Option2 
               Caption         =   "TXT"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   30
               Top             =   0
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.OptionButton Option2 
               Caption         =   "CSV"
               Height          =   255
               Index           =   1
               Left            =   720
               TabIndex        =   29
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   1920
            MaxLength       =   1
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   240
         MaxLength       =   4
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Prefisso File"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Chiudi"
      Height          =   735
      Index           =   0
      Left            =   7920
      TabIndex        =   6
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   495
      Index           =   1
      Left            =   8640
      TabIndex        =   5
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   240
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   5040
      Width           =   8295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   495
      Index           =   0
      Left            =   8640
      TabIndex        =   2
      Top             =   4275
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   4440
      Width           =   8295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Percorso Importazione"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   240
      TabIndex        =   3
      Top             =   4800
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Percorso Esportazione"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   4200
      Width           =   1875
   End
End
Attribute VB_Name = "frmImpEsp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TXT_BONUS = 0
Private Const TXT_BONUS_SEPARATORE = 6

Private Const TXT_FATTURE = 1
Private Const TXT_FATTURE_SEPARATORE = 3

Private Const TXT_MOVIMENTI = 7
Private Const TXT_MOVIMENTI_SEPARATORE = 4

Private Const TXT_CODICICER = 8
Private Const TXT_CODICICER_SEPARATORE = 9

Private Const TXT_UTENTI = 11
Private Const TXT_UTENTI_SEPARATORE = 10

Private Const TXT_ACCESSI = 12
Private Const TXT_ACCESSI_SEPARATORE = 13


Private Const TXT_PERCORSO_IMP = 5
Private Const TXT_PERCORSO_ESP = 2


Private Sub Command1_Click(Index As Integer)
    Dim sFile As String
    
    Select Case Index
        Case 0
            DoCommonDialog DT_Browse, sFile, "DDD", "Indicare Percorso", Me.hWnd, , , Text1(TXT_PERCORSO_ESP).Text, "*.*"

            If sFile <> vbNullString Then
                Text1(TXT_PERCORSO_ESP).Text = sFile
            End If
        Case 1
            DoCommonDialog DT_Browse, sFile, "DDD", "Indicare Percorso", Me.hWnd, , , Text1(TXT_PERCORSO_IMP).Text, "*.*"

            If sFile <> vbNullString Then
                Text1(TXT_PERCORSO_IMP).Text = sFile
            End If
        
    End Select
    
End Sub

Private Sub Command2_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            salva
            Unload Me
    End Select
End Sub

Private Sub salva()
    Dim f As New FileSystemObject
    
    If Not f.FolderExists(Text1(TXT_PERCORSO_ESP).Text) Then
        MsgBox "Attenzione percorso Esportazione " & Text1(TXT_PERCORSO_ESP).Text & " non trovato", vbCritical + vbOKOnly, App.Title
        Exit Sub
    End If
    
    If Not f.FolderExists(Text1(TXT_PERCORSO_IMP).Text) Then
        MsgBox "Attenzione percorso Importazione " & Text1(TXT_PERCORSO_IMP).Text & " non trovato", vbCritical + vbOKOnly, App.Title
        Exit Sub
    End If
    
    gFattureFile = Text1(TXT_FATTURE)
    gFattureFileType = IIf(Option2(0).value = True, 0, 1)
    gFattureSeparatore = Text1(TXT_FATTURE_SEPARATORE)

    gMovimentiFile = Text1(TXT_MOVIMENTI)
    gMovimentiFileType = IIf(Option3(0).value = True, 0, 1)
    gMovimentiSeparatore = Text1(TXT_MOVIMENTI_SEPARATORE)

    gCodiciCERFile = Text1(TXT_CODICICER)
    gCodiciCERFileType = IIf(Option4(0).value = True, 0, 1)
    gCodiciCERSeparatore = Text1(TXT_CODICICER_SEPARATORE)

    gUtentiFile = Text1(TXT_UTENTI)
    gUtentiFileType = IIf(Option5(0).value = True, 0, 1)
    gUtentiSeparatore = Text1(TXT_UTENTI_SEPARATORE)


    gAccessiFile = Text1(TXT_ACCESSI).Text
    gAccessiFileType = IIf(Option6(0).value = True, 0, 1)
    gAccessiSeparatore = Text1(TXT_ACCESSI_SEPARATORE).Text


    gPercorsoImp = Text1(TXT_PERCORSO_IMP).Text
    gPercorsoEsp = Text1(TXT_PERCORSO_ESP).Text
    gFTP_URL = FTPURLTxT.Text
    gFTP_UID = FTPUserTxt.Text
    gFTP_PWD = FTPPasswordTxT.Text
    gFTP_Active = FTPCheck.value

    'HTTP
    gTipoInvioAccessi = TipoInvioAccessiCMB.ListIndex
    gURLTokenSicurezza = URLTokenSicurezzaTxt.Text
    gURLWebService = URLWebServiceTxt.Text


    SaveSetting App.Title, "IMPOSTAZIONI", "FATTUREFILE", gFattureFile
    SaveSetting App.Title, "IMPOSTAZIONI", "FATTUREFILETYPE", gFattureFileType
    SaveSetting App.Title, "IMPOSTAZIONI", "FATTUREFILESEPARATORE", gFattureSeparatore

    SaveSetting App.Title, "IMPOSTAZIONI", "MOVIMENTIFILE", gMovimentiFile
    SaveSetting App.Title, "IMPOSTAZIONI", "MOVIMENTIFILETYPE", gMovimentiFileType
    SaveSetting App.Title, "IMPOSTAZIONI", "MOVIMENTIFILESEPARATORE", gMovimentiSeparatore
    
    SaveSetting App.Title, "IMPOSTAZIONI", "CODICICERFILE", gCodiciCERFile
    SaveSetting App.Title, "IMPOSTAZIONI", "CODICICERFILETYPE", gCodiciCERFileType
    SaveSetting App.Title, "IMPOSTAZIONI", "CODICICERFILESEPARATORE", gCodiciCERSeparatore

    SaveSetting App.Title, "IMPOSTAZIONI", "UTENTIFILE", gUtentiFile
    SaveSetting App.Title, "IMPOSTAZIONI", "UTENTIFILETYPE", gUtentiFileType
    SaveSetting App.Title, "IMPOSTAZIONI", "UTENTIFILESEPARATORE", gUtentiSeparatore

    SaveSetting App.Title, "IMPOSTAZIONI", "ACCESSIFILE", gAccessiFile
    SaveSetting App.Title, "IMPOSTAZIONI", "ACCESSIFILETYPE", gAccessiFileType
    SaveSetting App.Title, "IMPOSTAZIONI", "ACCESSIFILESEPARATORE", gAccessiSeparatore


    SaveSetting App.Title, "IMPOSTAZIONI", "PERCORSOIMP", gPercorsoImp
    SaveSetting App.Title, "IMPOSTAZIONI", "PERCORSOESP", gPercorsoEsp
    
    'FTP
    SaveSetting App.Title, "IMPOSTAZIONI", "FTPURL", gFTP_URL
    SaveSetting App.Title, "IMPOSTAZIONI", "FTPUID", gFTP_UID
    SaveSetting App.Title, "IMPOSTAZIONI", "FTPPWD", gFTP_PWD
    SaveSetting App.Title, "IMPOSTAZIONI", "FTPACTIVE", gFTP_Active
    
    'HTTP
    SaveSetting App.Title, "IMPOSTAZIONI", "TIPOINVIOACCESSI", gTipoInvioAccessi
    SaveSetting App.Title, "IMPOSTAZIONI", "URLTOKENSICUREZZA", gURLTokenSicurezza
    SaveSetting App.Title, "IMPOSTAZIONI", "URLWEBSERVICE", gURLWebService
    
End Sub
Private Sub Form_Load()
    
    Text1(TXT_FATTURE).Text = gFattureFile
    Option2(gFattureFileType).value = True
    Text1(TXT_FATTURE_SEPARATORE).Text = gFattureSeparatore
    
    Text1(TXT_MOVIMENTI).Text = gMovimentiFile
    Option3(gMovimentiFileType).value = True
    Text1(TXT_MOVIMENTI_SEPARATORE).Text = gMovimentiSeparatore
    
    Text1(TXT_CODICICER).Text = gCodiciCERFile
    Option4(gCodiciCERFileType).value = True
    Text1(TXT_CODICICER_SEPARATORE).Text = gCodiciCERSeparatore
    
    Text1(TXT_UTENTI).Text = gUtentiFile
    Option5(gUtentiFileType).value = True
    Text1(TXT_UTENTI_SEPARATORE).Text = gUtentiSeparatore
    
    Text1(TXT_ACCESSI).Text = gAccessiFile
    Option6(gAccessiFileType).value = True
    Text1(TXT_ACCESSI_SEPARATORE).Text = gAccessiSeparatore
    
    
    Text1(TXT_PERCORSO_IMP).Text = gPercorsoImp
    Text1(TXT_PERCORSO_ESP).Text = gPercorsoEsp
    
    'HTTP
    TipoInvioAccessiCMB.ListIndex = gTipoInvioAccessi
    URLTokenSicurezzaTxt.Text = gURLTokenSicurezza
    URLWebServiceTxt.Text = gURLWebService
    
    
    FTPURLTxT.Text = gFTP_URL
    FTPUserTxt.Text = gFTP_UID
    FTPPasswordTxT.Text = gFTP_PWD
    FTPCheck.value = IIf(gFTP_Active, 1, 0)
    applicaRestrizioniUtente
End Sub

Public Sub applicaRestrizioniUtente()
    Command2(1).enabled = frmLogin.Livello <> LVL_USER
End Sub

