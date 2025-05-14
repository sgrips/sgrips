VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmParametriSistema 
   Caption         =   "Parametri Sistema"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd 
      Caption         =   "&Chiudi"
      Height          =   390
      Left            =   8010
      TabIndex        =   1
      Top             =   6120
      Width           =   1365
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmParametriSistema.frx":0000
      Height          =   5880
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   10372
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   19
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "CodIstat"
         Caption         =   "CodIstat"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Parametro"
         Caption         =   "Parametro"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Descrizione"
         Caption         =   "Descrizione"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Valore"
         Caption         =   "Valore"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "ValoreMin"
         Caption         =   "ValoreMin"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "ValoreMax"
         Caption         =   "ValoreMax"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   4649,953
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1395,213
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1379,906
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1395,213
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ado1 
      Height          =   330
      Left            =   90
      Top             =   6060
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
   Begin MSComCtl2.DTPicker DTPAnno 
      Height          =   405
      Left            =   6870
      TabIndex        =   2
      Top             =   6105
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy"
      Format          =   60620803
      UpDown          =   -1  'True
      CurrentDate     =   36967
   End
End
Attribute VB_Name = "frmParametriSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mAnnoRiferimento As String
Public mCodComune As String
Public mAbitanti As String

Public Sub Carica(Optional sCodComune As String = vbNullString)
    Dim qry As New ADODB.Command

   
    mAnnoRiferimento = Year(Now)
    'se la modalità è tutti i comuni chiedo di quale comune si vanno ad inserire i dati
    If sCodComune = vbNullString Then
        If gComuneAttivo = ALL_COMUNI Then
            frmComuneDaLavorare.Show vbModal
            sCodComune = GetSetting(App.Title, SEZ_ULTIMI_VALORI, KEY_ULTIMI_VALORI_COMUNESEL, "")
        Else
            sCodComune = gComuneAttivo
        End If
    End If
    
    mCodComune = sCodComune
    Me.Caption = "Parametri Applicativi:" & getDesComune(sCodComune)
    
    Me.Show vbModal
End Sub

Private Function CaricaDati(Optional SuperUtente As Boolean = False) As Boolean
    Dim rs As Recordset
    
    CaricaDati = False
    'Reperisco i dati comunali e anagrafici
    Set rs = DatiAnagraficiComunali(mCodComune, mAnnoRiferimento, mAbitanti)
    If rs.EOF Then
        
        If ADO1.Recordset Is Nothing Then
            ADO1.RecordSource = "SELECT ParametriSistema.*  FROM ParametriSistema where   ParametriSistema.CodIstat='-1'"
        Else
            If ADO1.Recordset.State = adStateOpen Then
                ADO1.Recordset.Close
                ADO1.RecordSource = "SELECT ParametriSistema.*  FROM ParametriSistema where   ParametriSistema.CodIstat='-1'"
            End If
        End If
    Else
        If Not SuperUtente Then
            ADO1.RecordSource = "SELECT ParametriSistema.*  FROM ParametriSistema where   (ParametriSistema.AnnoRiferimento='" & _
                            mAnnoRiferimento & "' and ParametriSistema.CodIstat='" & mCodComune & "' or (ParametriSistema.CodIstat='ALL'" & _
                            " and  not   ParametriSistema.Parametro in  (Select  ParametriSistema.Parametro from ParametriSistema  where ParametriSistema.AnnoRiferimento='" & _
                            mAnnoRiferimento & "' and  ParametriSistema.CodIstat='" & mCodComune & "'))) and NOT privata"
        Else
            ADO1.RecordSource = "SELECT ParametriSistema.*  FROM ParametriSistema where   (ParametriSistema.AnnoRiferimento='" & mAnnoRiferimento & "' and ParametriSistema.CodIstat='" & _
                            mCodComune & "' or (ParametriSistema.CodIstat='ALL' and  not   ParametriSistema.Parametro in  (Select  ParametriSistema.Parametro from ParametriSistema  where  ParametriSistema.AnnoRiferimento='1900'" & _
                            " and ParametriSistema.CodIstat='" & mCodComune & "')))"
        End If
    End If
    
    
    On Error Resume Next
    ADO1.Recordset.Close
    On Error GoTo 0
    
    ADO1.Refresh

    CaricaDati = Not ADO1.Recordset.EOF
    
    If Not CaricaDati Then
        DataGrid1.RowHeight = 30
    Else
        DataGrid1.RowHeight = 270
        DataGrid1.Refresh
    End If
    
End Function



Private Sub Cmd_Click()
    
    Unload Me
    
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        SendKeys "{ENTER}", True
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Static KeyCode1 As Integer
    Static KeyCode2 As Integer
    If WindowState = vbMinimized Then Exit Sub
    'Se si preme in sequenza G F A tenendo premuto il CTRL
    If (Shift And vbCtrlMask) > 0 Then
        If KeyCode = Asc("G") Then
            KeyCode1 = Asc("F")
        ElseIf KeyCode1 = KeyCode Then
                KeyCode2 = Asc("A")
            ElseIf KeyCode2 = KeyCode Then
                    KeyCode2 = 0
                    KeyCode1 = 0
                    
                    MsgBox "Benvenuto Super Utente", vbOKOnly, App.Title
                    
                    CaricaDati True
            End If
    End If
End Sub

Private Sub Form_Load()
    Dim bCaricaDati  As Boolean

    MDIMain.mnuComprimi.Enabled = False
    ADO1.ConnectionString = gDBC
    
    DataGrid1.Left = 0
    DataGrid1.Top = 0
    DataGrid1.Height = cmd.Top - 15 * 4
    DataGrid1.Width = Me.ScaleWidth
    
   
    ADO1.ConnectionString = gDBC
    ADO1.CommandType = adCmdText
    
    DTPAnno.Value = Format$(Now, "yyyy-mm-dd")
    
    bCaricaDati = CaricaDati
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.mnuComprimi.Enabled = True
End Sub

Private Sub DTPAnno_Change()
    mAnnoRiferimento = Year(DTPAnno.Value)
    CaricaDati
End Sub
