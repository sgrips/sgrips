VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmComuneDaLavorare 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleziona il comune"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataListLib.DataCombo DBCmbComuni 
      Bindings        =   "frmComuneDaLavorare.frx":0000
      DataSource      =   "ADO1"
      Height          =   315
      Left            =   1125
      TabIndex        =   2
      ToolTipText     =   "Selezionare un comune"
      Top             =   1935
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Comune"
      BoundColumn     =   "CODISTAT"
      Text            =   "DataCombo1"
   End
   Begin MSAdodcLib.Adodc ADO1 
      Height          =   360
      Left            =   3360
      Top             =   2505
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   635
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
      ConnectStringType=   3
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
      Caption         =   "Adodc1"
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
   Begin VB.CommandButton BubbleButton1 
      Caption         =   "&Conferma"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2490
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   $"frmComuneDaLavorare.frx":002A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmComuneDaLavorare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub BubbleButton1_Click()
    Dim sCodComune As String
    SaveSetting App.Title, SEZ_ULTIMI_VALORI, KEY_ULTIMI_VALORI_COMUNESEL, DBCmbComuni.BoundText
    Unload Me
End Sub

Private Sub Form_Activate()
    'Reperisco dal file registro l'ultimo comune selezionato
    DBCmbComuni.BoundText = GetSetting(App.Title, SEZ_ULTIMI_VALORI, KEY_ULTIMI_VALORI_COMUNESEL, "")
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    DoEvents
    On Error Resume Next
    ADO1.Recordset.Close
    ADO1.CommandType = adCmdText
    ADO1.RecordSource = "Select * from comuni " & gWhereSuComune & " order by Comune "
    
    ADO1.ConnectionString = gDBC
    ADO1.Refresh
    
    Screen.MousePointer = vbDefault
End Sub
