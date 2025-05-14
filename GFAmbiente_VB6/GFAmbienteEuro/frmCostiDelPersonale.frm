VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCostiDelPersonale 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Costi - Del Personale"
   ClientHeight    =   8892
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   14028
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8892
   ScaleWidth      =   14028
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc ADO1 
      Height          =   345
      Left            =   9120
      Top             =   6885
      Visible         =   0   'False
      Width           =   1410
      _ExtentX        =   2477
      _ExtentY        =   614
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Modifica Retribuzione"
      Height          =   540
      Index           =   2
      Left            =   9120
      TabIndex        =   9
      Top             =   8235
      Width           =   1470
   End
   Begin VB.TextBox txtValoreCampo 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   1
      Left            =   11640
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   5610
      Width           =   2250
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aggiungi"
      Height          =   540
      Index           =   0
      Left            =   10770
      TabIndex        =   6
      Top             =   8235
      Width           =   1470
   End
   Begin VB.TextBox txtValoreCampo 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   0
      Left            =   11640
      TabIndex        =   4
      Text            =   "0"
      Top             =   5190
      Width           =   2235
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Totalizza"
      Height          =   540
      Index           =   1
      Left            =   12405
      TabIndex        =   1
      Top             =   8235
      Width           =   1500
   End
   Begin VB.TextBox txtValoreCampo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   360
      Index           =   2
      Left            =   11625
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   6045
      Width           =   2265
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3615
      Left            =   90
      TabIndex        =   3
      Top             =   5160
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   6371
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCostiDelPersonale.frx":0000
      Height          =   4965
      Left            =   105
      TabIndex        =   10
      Top             =   0
      Width           =   13875
      _ExtentX        =   24469
      _ExtentY        =   8763
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   2
      RowHeight       =   19
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
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
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "CodComune"
         Caption         =   "CodComune"
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
         DataField       =   "DataInserimento"
         Caption         =   "DataInserimento"
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
         DataField       =   "CodTipoPersonale"
         Caption         =   "CodTipoPersonale"
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
         DataField       =   "CGG"
         Caption         =   "CGG"
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
         DataField       =   "Descrizione"
         Caption         =   "Personale"
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
         DataField       =   "Liv2"
         Caption         =   "2° Livello"
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
      BeginProperty Column06 
         DataField       =   "Liv3"
         Caption         =   "3° Livello"
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
      BeginProperty Column07 
         DataField       =   "Liv4"
         Caption         =   "4° Livello"
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
      BeginProperty Column08 
         DataField       =   "Liv5"
         Caption         =   "5° Livello"
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
      BeginProperty Column09 
         DataField       =   "Liv6"
         Caption         =   "6° Livello"
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
      BeginProperty Column10 
         DataField       =   "Liv7"
         Caption         =   "7° Livello"
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
      BeginProperty Column11 
         DataField       =   "Liv7q"
         Caption         =   "7° Liv. Quadro"
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
      BeginProperty Column12 
         DataField       =   "Liv8"
         Caption         =   "8° Livello"
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
      BeginProperty Column13 
         DataField       =   "Liv8q"
         Caption         =   "8° Liv. Quadro"
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
      BeginProperty Column14 
         DataField       =   "Amministrativo"
         Caption         =   "Amministrativo"
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
      BeginProperty Column15 
         DataField       =   "Manager"
         Caption         =   "Manager"
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
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   1319,811
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            WrapText        =   -1  'True
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   3288,189
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            WrapText        =   -1  'True
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            WrapText        =   -1  'True
            ColumnWidth     =   780,095
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            WrapText        =   -1  'True
            ColumnWidth     =   731,906
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            WrapText        =   -1  'True
            ColumnWidth     =   768,189
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            WrapText        =   -1  'True
            ColumnWidth     =   768,189
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            WrapText        =   -1  'True
            ColumnWidth     =   768,189
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            WrapText        =   -1  'True
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column12 
            Alignment       =   1
            WrapText        =   -1  'True
            ColumnWidth     =   815,811
         EndProperty
         BeginProperty Column13 
            Alignment       =   1
            WrapText        =   -1  'True
            ColumnWidth     =   912,189
         EndProperty
         BeginProperty Column14 
            Alignment       =   1
            ColumnWidth     =   1620,284
         EndProperty
         BeginProperty Column15 
            Alignment       =   1
            ColumnWidth     =   1128,189
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCampi 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "50% del Costo Personale:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   9240
      TabIndex        =   8
      Top             =   5625
      Width           =   2310
   End
   Begin VB.Label lblCampi 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Altri Costi Di Gestione:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   23
      Left            =   9630
      TabIndex        =   5
      Top             =   5205
      Width           =   1965
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Totale Costi Gen.Gestione:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   9030
      TabIndex        =   2
      Top             =   6165
      Width           =   2535
   End
End
Attribute VB_Name = "frmCostiDelPersonale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mComune As String
Private mAnnoRif As String
Private mTotale As Double
Private mCostiGeneraliGestione As Double
Private OldTxtValue As String

Private Const cmd_Retribuzione = 2
Private Const cmd_Totalizza = 1
Private Const cmd_Aggiungi = 0

Function Carica(Comune As String, AnnoRif As String, CostiGeneraliGestione As Double) As Double

    mComune = Comune
    mAnnoRif = AnnoRif
    mCostiGeneraliGestione = CostiGeneraliGestione
    
    Me.Show vbModal

    Carica = mTotale

End Function



Private Sub Cmd_Click(Index As Integer)
    Dim i As Integer
    Dim lItem As ListItem
    
    
    Select Case Index
        Case cmd_Totalizza
            Totalizza
            Unload Me

        Case cmd_Aggiungi
            With lvw.ListItems
            For i = 1 To .Count
                If .Item(i).Selected Then
                    Set lItem = .Item(i)
                    Aggiungi lItem
                End If
            Next
            
            CaricaDati
            End With
        Case cmd_Retribuzione
            DoEvents
            On Error Resume Next
            frmRetribuzionePersonale.Carica mComune, mAnnoRif
            Totalizza
    End Select
End Sub

Private Function EsisteGiaPersonale(Codice As Long) As Boolean
    Dim rs As Recordset
    
    Set rs = gDBC.Execute("Select * from PersonaleImpiegatoNelComune WHERE CodComune='" & mComune & "' and AnnoRiferimento='" & mAnnoRif & "' and CodTipoPersonale=" & Codice)
    EsisteGiaPersonale = Not rs.EOF
End Function

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    txtValoreCampo(2) = Format$(Totalizza, FORMATVALUTA)
End Sub

Function Totalizza(Optional CalcoloCostoPersonale As Boolean = True) As Double
    
    If CalcoloCostoPersonale Then CostoTotPersonale
    
    Totalizza = CDbl(txtValoreCampo(0)) + CDbl(txtValoreCampo(1))
    
    mTotale = Totalizza
    
End Function

Function CostoTotPersonale() As Double
    Dim i As Integer
    Dim ii As Integer
    Dim Valore As Double
    Dim Retribuzione As Double
    Static InCorso As Boolean
    
    mTotale = 0
    
    If InCorso Then Exit Function
    
    
    InCorso = True
    
    If Not ADO1.Recordset.EOF Then
        ADO1.Recordset.MoveFirst
        Do While Not ADO1.Recordset.EOF
            For ii = 5 To DataGrid1.Columns.Count - 1
                Valore = 0
                If DataGrid1.Columns(3) = True Then
                    Retribuzione = GetRetribuzione(mComune, mAnnoRif, DataGrid1.Columns(2), COD_LIVELLO & CStr(ii - 5 + 1))
                      
                    DataGrid1.Columns(ii) = IIf(Retribuzione = 0, 0, ADO1.Recordset(ii))
                    Valore = (0 & DataGrid1.Columns(ii)) * Retribuzione
                    
                End If
                mTotale = mTotale + Valore
            Next
            ADO1.Recordset.MoveNext
        Loop
        ADO1.Recordset.MoveFirst
    End If
    

    InCorso = False
    Valore = mTotale * (GetParamValue(mAnnoRif, mComune, "CPPERSON") / 100)
    
    txtValoreCampo(1) = Format$(Valore, FORMATVALUTA)
    
    CostoTotPersonale = Valore
    
    mTotale = CostoTotPersonale

End Function




Private Sub DataGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    'I campi devono essere numerici
    If Not IsNumeric(DataGrid1.Columns(ColIndex).Text) Then
        MsgBox "Il valore non è numerico", vbExclamation + vbOKOnly, App.Title
        Cancel = True
    End If
    
    'Verifica se il livello appartiene al tipo personale
    If GetRetribuzione(mComune, mAnnoRif, DataGrid1.Columns(2), COD_LIVELLO & CStr(ColIndex - 5 + 1)) = 0 Then
        MsgBox "Per il livello selezionato non è previsto personale di tipo: ( " & DataGrid1.Columns(4) & " )"
        Cancel = True
    End If
End Sub

Private Sub DataGrid1_BeforeDelete(Cancel As Integer)

    If MsgBox("Cancellare tipo personale:(" & DataGrid1.Columns(4) & ")", vbQuestion + vbYesNo, App.Title) = vbYes Then
            gDBC.BeginTrans
            gDBC.Execute ("Delete from PersonaleImpiegatoNelComune  WHERE PersonaleImpiegatoNelComune.CodComune='" & mComune & "' and PersonaleImpiegatoNelComune.AnnoRiferimento='" & mAnnoRif & "' and PersonaleImpiegatoNelComune.CodTipoPersonale=" & DataGrid1.Columns(2))
            gDBC.CommitTrans
            CaricaDati
            
            txtValoreCampo(2) = Format$(Totalizza, FORMATVALUTA)
    End If
    
    Cancel = 1
End Sub

Private Sub DataGrid1_Error(ByVal DataError As Integer, Response As Integer)
    If DataError = 7011 Then Response = 0
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        SendKeys "{ENTER}", True
    End If

End Sub

Private Sub Form_Activate()
    
    DataGrid1.SetFocus
    
End Sub


Private Sub Form_Load()
    

    ADO1.ConnectionString = gDBC
    
    ADO1.CommandType = adCmdText
    ADO1.RecordSource = "SELECT PersonaleImpiegatoNelComune.CodComune, PersonaleImpiegatoNelComune.AnnoRiferimento, PersonaleImpiegatoNelComune.CodTipoPersonale,TipiPersonaleImpiegato.CGG, " _
    & " TipiPersonaleImpiegato.Descrizione, PersonaleImpiegatoNelComune.LIV2, PersonaleImpiegatoNelComune.LIV3,PersonaleImpiegatoNelComune.LIV4,PersonaleImpiegatoNelComune.LIV5, " _
    & " PersonaleImpiegatoNelComune.LIV6,PersonaleImpiegatoNelComune.LIV7,PersonaleImpiegatoNelComune.LIV7Q,PersonaleImpiegatoNelComune.LIV8,PersonaleImpiegatoNelComune.LIV8Q,PersonaleImpiegatoNelComune.Amministrativo, PersonaleImpiegatoNelComune.Manager " _
    & " FROM PersonaleImpiegatoNelComune INNER JOIN TipiPersonaleImpiegato ON PersonaleImpiegatoNelComune.CodTipoPersonale = TipiPersonaleImpiegato.CodTipoPersonale" _
    & " WHERE PersonaleImpiegatoNelComune.CodComune='" & mComune & "' and PersonaleImpiegatoNelComune.AnnoRiferimento='" & mAnnoRif & "'"
    ADO1.ConnectionString = gDBC
    
    lvw.SmallIcons = MDIMain.imglist
    
    txtValoreCampo(1).BackColor = BACKCOLOR_CAMPI_CALC
    txtValoreCampo(2).BackColor = BACKCOLOR_CAMPI_CALC
    CaricaInListView lvw, ICON_PERSONALE, Lvw_Personale
    
    lblCampi(0) = GetParamValue(mAnnoRif, mComune, "CPPERSON") & " % del Costo Personale:"
    
    
    
    If Not CaricaDati Then
    End If

    txtValoreCampo(0) = mCostiGeneraliGestione
    CostoTotPersonale
    txtValoreCampo(0) = Format$(mCostiGeneraliGestione - CDbl(txtValoreCampo(1)), FORMATVALUTA)
        
    RicalcolaTot
End Sub


Private Function CaricaDati() As Boolean
    Dim rs As ADODB.Recordset
    CaricaDati = False
    On Error Resume Next
    ADO1.Recordset.Close
    On Error GoTo 0
    
    ADO1.Refresh
    
    CaricaDati = Not ADO1.Recordset.EOF
    
    If Not CaricaDati Then
        DataGrid1.RowHeight = 30
    Else
        DataGrid1.RowHeight = 270
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    ADO1.Recordset.Close
End Sub


Private Sub Aggiungi(Item As ListItem)
    If Item.Selected Then
        If EsisteGiaPersonale(Item.Text) Then
            MsgBox "Personale:" & Item.SubItems(1) & vbCrLf & "già presente", vbExclamation + vbOKOnly
        Else
            gDBC.BeginTrans
            gDBC.Execute "Insert into PersonaleImpiegatoNelComune (CodComune,AnnoRiferimento,CodTipoPersonale) Values('" & mComune & "','" & mAnnoRif & "'," & CDbl(Item.Text) & ")"
            gDBC.CommitTrans
        End If
    End If
End Sub


Private Sub Lvw_DblClick()
    If lvw.SelectedItem Is Nothing Then Exit Sub
    Aggiungi lvw.SelectedItem
    CaricaDati
End Sub




Private Sub txtValoreCampo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
End Sub

Private Sub txtValoreCampo_LostFocus(Index As Integer)
    
    txtValoreCampo(Index) = Format$(CDbl(txtValoreCampo(Index)), FORMATVALUTA)
    txtValoreCampo(2) = Format$(CDbl(txtValoreCampo(0)) + CDbl(txtValoreCampo(1)), FORMATVALUTA)
    
End Sub

Private Sub txtValoreCampo_GotFocus(Index As Integer)
    
    txtValoreCampo(Index).SelStart = 0
    txtValoreCampo(Index).SelLength = Len(txtValoreCampo(Index))
    OldTxtValue = txtValoreCampo(Index)
End Sub

Private Sub txtValoreCampo_Validate(Index As Integer, Cancel As Boolean)
        'Il valore deve essere numerico
        If Not IsNumeric(txtValoreCampo(Index)) Then MsgBox "Attenzione il valore deve essere numerico", vbOKOnly + vbExclamation, App.Title: Cancel = True: txtValoreCampo(Index) = OldTxtValue: Exit Sub

End Sub

Sub RicalcolaTot()
    If mCostiGeneraliGestione > 0 Then
        CostoTotPersonale
    End If
    
    txtValoreCampo(2) = Format$(Totalizza(False), FORMATVALUTA)
    If Not ADO1.Recordset.EOF Then
        DataGrid1.Row = 0
        DataGrid1.Col = 4
        DataGrid1.SelStart = 1
        DataGrid1.SelLength = Len(DataGrid1.Columns(4))
    End If
End Sub
