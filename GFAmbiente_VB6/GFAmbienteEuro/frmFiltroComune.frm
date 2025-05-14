VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFiltroComune 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filtro"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc ADO1 
      Height          =   345
      Left            =   2280
      Top             =   2490
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   609
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "ADO1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comune"
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   4335
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmFiltroComune.frx":0000
         DataSource      =   "ADO1"
         Height          =   315
         Left            =   600
         TabIndex        =   10
         Top             =   345
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Comune"
         BoundColumn     =   "CODISTAT"
         Text            =   ""
      End
      Begin VB.OptionButton Option2 
         Caption         =   "  Tutti"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      Begin MSComCtl2.DTPicker DTPDataDA 
         Height          =   315
         Left            =   600
         TabIndex        =   3
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   19726337
         UpDown          =   -1  'True
         CurrentDate     =   36315
      End
      Begin MSComCtl2.DTPicker DTPDataA 
         Height          =   315
         Left            =   2760
         TabIndex        =   4
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   19726337
         UpDown          =   -1  'True
         CurrentDate     =   36315
      End
      Begin VB.Label Label2 
         Caption         =   "A"
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Da"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Annulla"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Conferma"
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   3240
      Width           =   1095
   End
End
Attribute VB_Name = "frmFiltroComune"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sComune As String
Dim sData As String

Private mDataDa As Date
Private mDataA As Date
Private mNOSave As Boolean
Private mCodComune As String


Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            
            
            If DTPDataDA.Value > DTPDataDA.Value Then MsgBox "Attenzione la DataDA deve essere minere della DataA", vbInformation + vbOKOnly, App.Title: Exit Sub
            
            
            If Option1.Value = True Then
                 sComune = " CodComune='" & DataCombo1.BoundText & "' and "
            Else
                sComune = vbNullString
            End If
            
            
            sData = "AnnoRiferimento between year(" & Chr$(34) & DTPDataDA.Value & Chr$(34) & ") and Year(" & Chr$(34) & DTPDataA.Value & Chr$(34) & ")"
            If gWhereSuComune = vbNullString Then
                gWhereFiltroSuComune = "Where "
            Else
                gWhereFiltroSuComune = " And "
            End If
            gWhereFiltroSuComune = gWhereFiltroSuComune & sComune & sData
            gcmdAnnulla = False
            
            mDataDa = DTPDataDA
            mDataA = DTPDataA
            mCodComune = IIf(Option2.Value, ALL_COMUNI, DataCombo1.Text)
            
            If mNOSave = False Then
                SaveSetting App.Title, SEZ_FILTRI, KEY_FILTRI_DATADA, mDataDa
                SaveSetting App.Title, SEZ_FILTRI, KEY_FILTRI_DATAA, mDataA
                SaveSetting App.Title, SEZ_FILTRI, KEY_FILTRI_COMUNE, mCodComune
            End If

            mNOSave = False

            
            Unload Me
        Case 1
            gcmdAnnulla = True
            Unload Me
    End Select
        
End Sub

Public Sub SetFiltro(DataDa As Date, DataA As Date, CodComune As String)
    
    'set filtro
    mDataDa = DataDa
    mDataA = DataA
    mCodComune = CodComune
    mNOSave = True
    
    Me.Show vbModal
    
    DataDa = mDataDa
    DataA = mDataA
    CodComune = mCodComune
    
End Sub




Private Sub Form_Load()
    
    Set DataCombo1.DataSource = Nothing
    ADO1.ConnectionString = gDBC
    ADO1.CommandType = adCmdText
    ADO1.RecordSource = "Select comune,CODISTAT from comuni " & gWhereSuComune & " order by comune"
    ADO1.Refresh
    Set DataCombo1.DataSource = ADO1
    
    
    'Leggo le impostazioni per il filtro da file registro
    If mNOSave = False Then
        DTPDataDA = GetSetting(App.Title, SEZ_FILTRI, KEY_FILTRI_DATADA, DateAdd("YYYY", -1, Now))
        DTPDataA = GetSetting(App.Title, SEZ_FILTRI, KEY_FILTRI_DATAA, Now)
        DataCombo1.Text = GetSetting(App.Title, SEZ_FILTRI, KEY_FILTRI_COMUNE, ALL_COMUNI)
    Else
        DTPDataDA = mDataDa
        DTPDataA = mDataA
        DataCombo1.Text = mCodComune
    End If
    
    If DataCombo1.Text = ALL_COMUNI Then
        Option2_Click
    Else
        Option1.Value = True
    End If

End Sub

Private Sub Option1_Click()
    DataCombo1.Enabled = True
End Sub

Private Sub Option2_Click()
    DataCombo1.Enabled = False
End Sub
