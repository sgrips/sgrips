VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEsportazoine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Esportazione Dati"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   2235
      Index           =   0
      Left            =   120
      ScaleHeight     =   2235
      ScaleWidth      =   6915
      TabIndex        =   10
      Top             =   4440
      Width           =   6915
      Begin VB.Frame Frame2 
         Height          =   1515
         Index           =   0
         Left            =   0
         TabIndex        =   18
         Top             =   660
         Width           =   6855
         Begin VB.ComboBox cmbTipoCampo 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   420
            Width           =   1515
         End
         Begin VB.TextBox txtRiempimento 
            Height          =   315
            Left            =   1740
            TabIndex        =   24
            Top             =   1080
            Width           =   915
         End
         Begin VB.ComboBox cmbAllineamento 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1080
            Width           =   1515
         End
         Begin VB.TextBox txtSeparatoreDecimale 
            Height          =   315
            Left            =   1740
            MaxLength       =   1
            TabIndex        =   22
            Top             =   420
            Width           =   735
         End
         Begin VB.TextBox txtSeparatoreData 
            Height          =   315
            Left            =   2580
            MaxLength       =   1
            TabIndex        =   21
            Top             =   420
            Width           =   735
         End
         Begin VB.TextBox txtSeparatoreOra 
            Height          =   315
            Left            =   3420
            MaxLength       =   1
            TabIndex        =   20
            Top             =   420
            Width           =   735
         End
         Begin VB.CheckBox chkTitoli 
            Caption         =   "Titoli colonne"
            Height          =   315
            Left            =   2820
            TabIndex        =   19
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Campo"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   180
            Width           =   1395
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Allineamento "
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   840
            Width           =   945
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Riempimento"
            Height          =   195
            Left            =   1740
            TabIndex        =   29
            Top             =   840
            Width           =   915
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Sep. Dec."
            Height          =   195
            Left            =   1740
            TabIndex        =   28
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Sep. Data"
            Height          =   195
            Left            =   2580
            TabIndex        =   27
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Sep. Ora"
            Height          =   195
            Left            =   3420
            TabIndex        =   26
            Top             =   180
            Width           =   630
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Separatore"
         Height          =   675
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   6855
         Begin VB.OptionButton optSeparatore 
            Caption         =   "Larghezza Fissa"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   17
            Top             =   300
            Width           =   1455
         End
         Begin VB.TextBox txtSeparatore 
            Height          =   255
            Left            =   6060
            MaxLength       =   1
            TabIndex        =   16
            Top             =   300
            Width           =   555
         End
         Begin VB.OptionButton optSeparatore 
            Caption         =   "Altro"
            Height          =   255
            Index           =   4
            Left            =   5340
            TabIndex        =   15
            Top             =   300
            Width           =   675
         End
         Begin VB.OptionButton optSeparatore 
            Caption         =   "virgola"
            Height          =   255
            Index           =   3
            Left            =   4440
            TabIndex        =   14
            Top             =   300
            Width           =   795
         End
         Begin VB.OptionButton optSeparatore 
            Caption         =   "punto virgola"
            Height          =   255
            Index           =   2
            Left            =   3060
            TabIndex        =   13
            Top             =   300
            Width           =   1215
         End
         Begin VB.OptionButton optSeparatore 
            Caption         =   "Tabulazione"
            Height          =   255
            Index           =   1
            Left            =   1740
            TabIndex        =   12
            Top             =   300
            Width           =   1215
         End
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   2235
      Index           =   1
      Left            =   120
      ScaleHeight     =   2235
      ScaleWidth      =   6915
      TabIndex        =   32
      Top             =   4440
      Width           =   6915
      Begin VB.CheckBox chkXML 
         Caption         =   "XML"
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   675
      End
      Begin VB.Label lblinfoXML 
         Height          =   795
         Left            =   60
         TabIndex        =   33
         Top             =   180
         Width           =   6795
      End
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Chiudi"
      Default         =   -1  'True
      Height          =   315
      Index           =   0
      Left            =   6000
      TabIndex        =   9
      Top             =   4020
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Percorso Esp."
      Height          =   315
      Left            =   4680
      TabIndex        =   8
      Top             =   4020
      Width           =   1275
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Esporta"
      Height          =   315
      Index           =   1
      Left            =   3540
      TabIndex        =   7
      Top             =   4020
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2655
      Left            =   60
      TabIndex        =   6
      Top             =   4080
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   4683
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Formato Testo"
            Key             =   "txt"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Formato XML"
            Key             =   "xml"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Esempio "
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtInfo 
      Height          =   675
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   4215
   End
   Begin VB.TextBox txtEsempio 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   1080
      Width           =   6975
   End
   Begin VB.ComboBox cmbTabelle 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   2715
   End
   Begin MSComctlLib.ProgressBar prbEsportazione 
      Height          =   375
      Left            =   1740
      TabIndex        =   4
      Top             =   720
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Tabella"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   2715
   End
End
Attribute VB_Name = "frmEsportazoine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mAnnoRiferimento As Long
Private mCodComune As String
Private mTitoliColonna As Boolean
Private mTipoCampo As TipoDato
Private mTipoSeparatoreCampo  As TipoSeparatoreCampo
Private rsTable As ADODB.Recordset 'mi rappresenta il recordset delle tabelle che si possono esportare
                                   'presenti in Esportazione
Private rsRowExample As ADODB.Recordset
Private mArrayTipoSeparatoreCampo(0 To 4) As String
Private mPath_to_export As String
Private mExportXML As Boolean
Private Const SEPARATORE = "ç"

Private Enum TipoSeparatoreCampo
    TIPO_SEP_CAMP_LARGHEZZA_FISSA = 0
    TIPO_SEP_CAMP_TAB = 1
    TIPO_SEP_CAMP_PUNTO_VIRGOLA = 2
    TIPO_SEP_CAMP_VIRGOLA = 3
    TIPO_SEP_CAMP_ALTRO = 4
End Enum

Private Enum TipoDato
    TIPO_NUMERICO = 0
    TIPO_TESTO = 1
    TIPO_DATA = 2
End Enum

Private Enum TipoAllineamento
    ASINISTRA = 0
    ADETRA = 1
End Enum

Private Type TipiDati
    Tipo As TipoDato
    Allineamento As TipoAllineamento
    SeparatoreData As String
    SeparatoreOra As String
    SeparatoreDecimale As String
    NumeroDecimali As Integer
    Riempimento As String
End Type

Private ArrayTipiDati(TIPO_NUMERICO To TIPO_DATA) As TipiDati

Public Sub Carica(Optional sCodComune As String = vbNullString, Optional AnnoRiferimento As String)
    On Error GoTo CaricaERR
    'se la modalità è tutti i comuni chiedo di quale comune si vanno ad inserire i dati
    If sCodComune = vbNullString Then
        'DataIns = DateAdd("yyyy", -1, Now)
        AnnoRiferimento = IIf(AnnoRiferimento = vbNullString, Year(Now), AnnoRiferimento)

        If gComuneAttivo = ALL_COMUNI Then
            frmComuneDaLavorare.Show vbModal
            sCodComune = GetSetting(App.Title, SEZ_ULTIMI_VALORI, KEY_ULTIMI_VALORI_COMUNESEL, "")
        Else
            sCodComune = gComuneAttivo
        End If
    End If
    
    mAnnoRiferimento = AnnoRiferimento
    mCodComune = sCodComune
    Me.Caption = "Esportazione dati del comune:" & getDesComune(sCodComune)
    caricoTabelleDaEsportare
    Exit Sub
CaricaERR:
    MsgBox Error$, vbCritical + vbOKOnly, App.Title
End Sub

Private Sub caricoTabelleDaEsportare()
    On Error GoTo CaricoTabelleDaEsportareERR
    'carico tabelle da TBL Esportazione
    Dim sql As String
    
    
    sql = "select * from esportazione where codComune ='ALL' or codComune ='" & mCodComune & "'"
    Set rsTable = CreateObject("ADODB.Recordset")
    With rsTable
        .Source = sql
        .LockType = adLockBatchOptimistic
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        Set .ActiveConnection = gDBC
        .Open
        Set .ActiveConnection = Nothing
    End With
   
    If Not rsTable.EOF Then
            
        Do While Not rsTable.EOF
            'carico combo
            cmbTabelle.AddItem rsTable("NomeTabella")
            cmbTabelle.ItemData(cmbTabelle.ListCount - 1) = cmbTabelle.ListCount - 1
            rsTable.MoveNext
        Loop
        
    End If
    
    cmbTabelle.ListIndex = -1
    cmbTabelle.ListIndex = 0
    Me.Show vbModal
    Exit Sub
CaricoTabelleDaEsportareERR:
    MsgBox Error$, vbCritical + vbOKOnly, App.Title
End Sub


Private Sub recuperoInfoTable()
    On Error GoTo RecuperoInfoTableERR
    
    'recupero info sulla tabella
    Dim eleSel As Integer
    Dim idxEle As Long
    Dim Info As String
    Dim sql As String
    
    eleSel = cmbTabelle.ListIndex
    idxEle = cmbTabelle.ItemData(eleSel)
    sql = "select * from " & cmbTabelle.List(eleSel) & " where codComune='" & mCodComune & "'"
    
    Set rsRowExample = CreateObject("ADODB.Recordset")
    
    If rsRowExample.State = adStateOpen Then
        rsRowExample.Close
    End If
    
    rsRowExample.MaxRecords = 10
    rsRowExample.Source = sql
    rsRowExample.CursorType = adOpenStatic
    rsRowExample.LockType = adLockBatchOptimistic
    rsRowExample.CursorLocation = adUseClient
    Set rsRowExample.ActiveConnection = gDBC
    rsRowExample.Open
    Set rsRowExample.ActiveConnection = Nothing
        
    Info = "numero record . . . ." & rsRowExample.RecordCount & vbCrLf
    Info = Info & "numero colonne . . ." & rsRowExample.fields.Count & vbCrLf
    Info = Info & "path esportazione. ." & mPath_to_export
    
    txtInfo = Info
    
    
    
    
    Exit Sub
RecuperoInfoTableERR:
    MsgBox Error$, vbCritical + vbOKOnly, App.Title
End Sub



Private Sub chkTitoli_Click()
 mTitoliColonna = chkTitoli.Value
End Sub

Private Sub chkXML_Click()
    mExportXML = chkXML.Value
End Sub

Private Sub cmbAllineamento_Click()
    Dim eleSel As Long
    
    eleSel = cmbAllineamento.ListIndex
    'allineamento
    ArrayTipiDati(mTipoCampo).Allineamento = cmbAllineamento.ItemData(eleSel)
End Sub

Private Sub cmbTipoCampo_Click()
    
    
    mTipoCampo = cmbTipoCampo.ListIndex
    
    'disable separatori non necessari
    txtSeparatoreData.Enabled = False
    txtSeparatoreData.BackColor = BACKCOLOR_CAMPI_CALC
    txtSeparatoreDecimale.Enabled = False
    txtSeparatoreDecimale.BackColor = BACKCOLOR_CAMPI_CALC
    txtSeparatoreOra.Enabled = False
    txtSeparatoreOra.BackColor = BACKCOLOR_CAMPI_CALC
    'disable allineamenti
    cmbAllineamento.Enabled = False
    cmbAllineamento.BackColor = BACKCOLOR_CAMPI_CALC
    'disable riempimento
    txtRiempimento.Enabled = False
    txtRiempimento.BackColor = BACKCOLOR_CAMPI_CALC
    If mTipoCampo = -1 Then Exit Sub
    Select Case cmbTipoCampo.ItemData(mTipoCampo)
        Case TIPO_TESTO
            cmbAllineamento.Enabled = True
            cmbAllineamento.BackColor = vbWhite
        Case TIPO_NUMERICO
            txtSeparatoreDecimale.Enabled = True
            txtSeparatoreDecimale.BackColor = vbWhite
            cmbAllineamento.Enabled = True
            cmbAllineamento.BackColor = vbWhite
            txtRiempimento.Enabled = True
            txtRiempimento.BackColor = vbWhite
        Case TIPO_DATA
            txtSeparatoreData.Enabled = True
            txtSeparatoreData.BackColor = vbWhite
            txtSeparatoreOra.Enabled = True
            txtSeparatoreOra.BackColor = vbWhite
    End Select
    
    'impostazione valori firmattazione
    cmbAllineamento.ListIndex = ArrayTipiDati(mTipoCampo).Allineamento
    txtRiempimento.Text = ArrayTipiDati(mTipoCampo).Riempimento
    txtSeparatoreData.Text = ArrayTipiDati(mTipoCampo).SeparatoreData
    txtSeparatoreDecimale.Text = ArrayTipiDati(mTipoCampo).SeparatoreDecimale
    txtSeparatoreOra.Text = ArrayTipiDati(mTipoCampo).SeparatoreOra
    
    
End Sub


Private Sub Cmd_Click(Index As Integer)
    Select Case Index
        Case 0
            saveSettingExport
            Unload Me
        Case 1
            esportazione
    End Select
        
End Sub

Private Sub cmbTabelle_Click()
    'al click carico 10 record esempio esportazione
    recuperoInfoTable
    loadSettingExport
End Sub

Private Sub Command1_Click()
    Dim XMLDom As DOMDocument
    If Not mExportXML Then
        applyFormat
    Else
        Set XMLDom = CreateObject("MSXML2.DOMDocument")
        XMLDom.preserveWhiteSpace = True
        rsRowExample.save XMLDom, adPersistXML
        
        txtEsempio = XMLDom.xml
    End If
End Sub

Private Sub Command2_Click()
     
    DoCommonDialog DT_Browse, mPath_to_export, "DDD", "Indicare Percorso dove esportare", Me.hwnd, "*.txt", , App.Path, "txt"
    
    If mPath_to_export = "" Then mPath_to_export = App.Path
    
    SaveSetting App.Title, "ESPORTAZIONE", "PATH", mPath_to_export
    
    recuperoInfoTable
End Sub

Private Sub TipiDatiDefault()
    'Tipo Testo
    ArrayTipiDati(TIPO_TESTO).Tipo = TIPO_TESTO
    ArrayTipiDati(TIPO_TESTO).Allineamento = ASINISTRA
    'Tipo Numerico
    ArrayTipiDati(TIPO_NUMERICO).Tipo = TIPO_NUMERICO
    ArrayTipiDati(TIPO_NUMERICO).Allineamento = ADETRA
    ArrayTipiDati(TIPO_NUMERICO).SeparatoreDecimale = ","
    ArrayTipiDati(TIPO_NUMERICO).Riempimento = "0"
    'Tipo Data
    ArrayTipiDati(TIPO_DATA).Tipo = TIPO_DATA
    ArrayTipiDati(TIPO_DATA).SeparatoreData = "/"
    ArrayTipiDati(TIPO_DATA).SeparatoreOra = ":"
End Sub


Private Sub Form_Load()
    
    lblinfoXML = "Il formato XML è un formato standard utilizzato per lo scambio " & _
                "dati attraverso internet. Se si vuole utilizzare questo formato " & _
                "selezionare il ceck-box seguente. Tutte le impostazioni relative al formato " & _
                "testo verranno comunque salvate, ma non utilizzate."
    TipiDatiDefault
    'carico combi statici
    cmbAllineamento.AddItem "A Sinistra"
    cmbAllineamento.ItemData(cmbAllineamento.ListCount - 1) = TipoAllineamento.ASINISTRA
    cmbAllineamento.AddItem "A Destra"
    cmbAllineamento.ItemData(cmbAllineamento.ListCount - 1) = TipoAllineamento.ADETRA
    
    cmbTipoCampo.AddItem "Numerico"
    cmbTipoCampo.ItemData(cmbTipoCampo.ListCount - 1) = TipoDato.TIPO_NUMERICO
    cmbTipoCampo.AddItem "Testo"
    cmbTipoCampo.ItemData(cmbTipoCampo.ListCount - 1) = TipoDato.TIPO_TESTO
    cmbTipoCampo.AddItem "Data"
    cmbTipoCampo.ItemData(cmbTipoCampo.ListCount - 1) = TipoDato.TIPO_DATA
    
    optSeparatore(mTipoSeparatoreCampo).Value = True
    mArrayTipoSeparatoreCampo(TipoSeparatoreCampo.TIPO_SEP_CAMP_LARGHEZZA_FISSA) = vbNullString
    mArrayTipoSeparatoreCampo(TipoSeparatoreCampo.TIPO_SEP_CAMP_TAB) = vbTab
    mArrayTipoSeparatoreCampo(TipoSeparatoreCampo.TIPO_SEP_CAMP_PUNTO_VIRGOLA) = ";"
    mArrayTipoSeparatoreCampo(TipoSeparatoreCampo.TIPO_SEP_CAMP_VIRGOLA) = ","
    mArrayTipoSeparatoreCampo(TipoSeparatoreCampo.TIPO_SEP_CAMP_ALTRO) = vbNullString
    
    mPath_to_export = GetSetting(App.Title, "ESPORTAZIONE", "PATH", App.Path)
    
    pic(0).Visible = True
    pic(1).Visible = False
    
    
End Sub




Private Sub saveSettingExport()
    'salvo le impostazioni
    Dim i As Integer
    Dim s As String
    s = ""
    On Error GoTo saveSettingExportERR
    For i = LBound(ArrayTipiDati) To UBound(ArrayTipiDati)
        s = s & ArrayTipiDati(i).Allineamento & SEPARATORE & ArrayTipiDati(i).NumeroDecimali & SEPARATORE _
              & ArrayTipiDati(i).Riempimento & SEPARATORE & ArrayTipiDati(i).SeparatoreData & SEPARATORE _
              & ArrayTipiDati(i).SeparatoreDecimale & SEPARATORE & ArrayTipiDati(i).SeparatoreOra & SEPARATORE _
              & ArrayTipiDati(i).Tipo & SEPARATORE
              
    Next
    
    'altri parametri -
    s = s & mTitoliColonna & SEPARATORE
    s = s & mExportXML & SEPARATORE
    s = s & mTipoSeparatoreCampo & SEPARATORE
    s = s & txtSeparatore.Text & SEPARATORE
    
    gDBC.Execute ("update Esportazione set UltimaEsportazione =now, parametri='" & s & "' where nometabella='" & cmbTabelle.List(cmbTabelle.ListIndex) & "'")
    Exit Sub
saveSettingExportERR:
    MsgBox Error, vbExclamation, App.Title
End Sub

Private Sub loadSettingExport()
    Dim s() As String
    Dim i As Integer
    Dim ii As Integer
    
    
    'imposto filtro
    rsTable.Filter = vbNullString
    rsTable.Filter = "nometabella='" & cmbTabelle.List(cmbTabelle.ListIndex) & "'"
    'load impostazioni sull'esportazione
    If Not IsNull(rsTable("parametri")) Then
        s = Split(rsTable("parametri"), SEPARATORE)
        On Error Resume Next
        i = -1
        i = UBound(s)
        On Error GoTo 0
        ii = 0
        If i > 0 Then
            For i = LBound(ArrayTipiDati) To (UBound(ArrayTipiDati) + 1) * 6 Step 7
                ArrayTipiDati(ii).Allineamento = s(i)
                ArrayTipiDati(ii).NumeroDecimali = s(i + 1)
                ArrayTipiDati(ii).Riempimento = s(i + 2)
                ArrayTipiDati(ii).SeparatoreData = s(i + 3)
                ArrayTipiDati(ii).SeparatoreDecimale = s(i + 4)
                ArrayTipiDati(ii).SeparatoreOra = s(i + 5)
                ArrayTipiDati(ii).Tipo = s(i + 6)
                ii = ii + 1
            Next
        End If
        'dal 21 id altri parametri
        mTitoliColonna = s(21)
        mExportXML = s(22)
        mTipoSeparatoreCampo = s(23)
        
        
        
        mArrayTipoSeparatoreCampo(TIPO_SEP_CAMP_ALTRO) = s(24)
        optSeparatore(mTipoSeparatoreCampo).Value = True
        txtSeparatore.Text = s(24)
        
        
        
        chkTitoli.Value = IIf(mTitoliColonna, 1, 0)
        chkXML.Value = IIf(mExportXML, 1, 0)
        
        
        cmbTipoCampo.ListIndex = -1
        cmbTipoCampo.ListIndex = 0
    Else
        TipiDatiDefault
    End If

End Sub


Private Function applyFormatField(testo As String, fld As Field)
    Dim data As Date
    Dim segno As Integer
    Dim simbDEC As String
    Dim size As Long
    
    If IsNumeric(testo) Then
        segno = Sgn(testo)
        testo = Abs(testo)
        size = Log(2 ^ (8 * fld.DefinedSize)) / Log(10#)
    Else
        size = fld.DefinedSize
    End If
    
    Select Case fld.Type
        'numeri interi
        Case DataTypeEnum.adInteger, DataTypeEnum.adBigInt
            If mTipoSeparatoreCampo = TIPO_SEP_CAMP_LARGHEZZA_FISSA Then
                testo = allinea(testo, ArrayTipiDati(TIPO_NUMERICO).Allineamento, ArrayTipiDati(TIPO_NUMERICO).Riempimento, size)
            Else
                testo = testo & mArrayTipoSeparatoreCampo(mTipoSeparatoreCampo)
            End If
            If segno <> 0 Then
                testo = IIf(segno < 1, "-", "+") & testo
            Else
                testo = " " & testo
            End If
        'numeri decimali
        Case DataTypeEnum.adDouble, DataTypeEnum.adDecimal, _
              DataTypeEnum.adDouble, DataTypeEnum.adNumeric
              
            If mTipoSeparatoreCampo = TIPO_SEP_CAMP_LARGHEZZA_FISSA Then
                testo = allinea(testo, ArrayTipiDati(TIPO_NUMERICO).Allineamento, ArrayTipiDati(TIPO_NUMERICO).Riempimento, size)
            Else
                testo = testo & mArrayTipoSeparatoreCampo(mTipoSeparatoreCampo)
            End If
            If segno <> 0 Then
                testo = IIf(segno < 1, "-", "+") & testo
            Else
                testo = " " & testo
            End If
            'sostituisco il simbolo di decimale
            simbDEC = Mid$(CStr(1.2), 2, 1)
            testo = Replace(testo, simbDEC, ArrayTipiDati(TIPO_NUMERICO).SeparatoreDecimale)
        'stringhe
        Case DataTypeEnum.adChar, DataTypeEnum.adVarChar, adVarWChar

            If mTipoSeparatoreCampo = TIPO_SEP_CAMP_LARGHEZZA_FISSA Then
                testo = allinea(testo, ArrayTipiDati(TIPO_TESTO).Allineamento, " ", size)
            Else
                testo = testo & mArrayTipoSeparatoreCampo(mTipoSeparatoreCampo)
            End If
        'date time
        Case DataTypeEnum.adDBTimeStamp, DataTypeEnum.adDate
                data = Format$(testo, "YYYY-MM-DD hh:nn:ss")
                'yyyy Year
                'q Quarter
                'm Month
                'y Day of year
                'd Day
                'w Weekday
                'ww Week
                'h Hour
                'n Minute
                's Second
                testo = DatePart("yyyy", data) & ArrayTipiDati(TIPO_DATA).SeparatoreData & _
                        allinea(DatePart("m", data), ADETRA, "0", 2) & ArrayTipiDati(TIPO_DATA).SeparatoreData & _
                        allinea(DatePart("d", data), ADETRA, "0", 2) & " " & _
                        DatePart("h", data) & ArrayTipiDati(TIPO_DATA).SeparatoreData & _
                        allinea(DatePart("n", data), ADETRA, "0", 2) & ArrayTipiDati(TIPO_DATA).SeparatoreData & _
                        allinea(DatePart("s", data), ADETRA, "0", 2) & " "
                testo = testo & mArrayTipoSeparatoreCampo(mTipoSeparatoreCampo)
                
        'boolean
        Case DataTypeEnum.adBoolean
                testo = IIf(CBool(testo), 1, 0)
                testo = testo & mArrayTipoSeparatoreCampo(mTipoSeparatoreCampo)
    End Select
    

    applyFormatField = testo
End Function

Private Function allinea(testo As String, tpa As TipoAllineamento, car_riemp As String, size As Long) As String
    Dim sTesto As String
    Dim sBlank As String
    Dim i As Long
    
    testo = testo & Space$(size - Len(testo))
    i = InStr(1, testo, " ")
    If i > 0 Then
        sTesto = Left$(testo, i - 1)
        sBlank = Mid$(testo, i)
    Else
        sTesto = testo
    End If
    
    If tpa = ADETRA Then
        sTesto = sBlank & sTesto
    Else
        sTesto = sTesto & sBlank
    End If
    
    allinea = Replace(sTesto, " ", car_riemp)
End Function


Private Sub applyFormat()
    Dim i As Long
    Dim s As String
    
    'applico il formato selezionato ai dati della tabella ...
    If rsRowExample Is Nothing Then Exit Sub
    rsRowExample.MoveFirst
    If rsRowExample.EOF Then Exit Sub
    
    txtEsempio = s
    If mTitoliColonna And mTipoSeparatoreCampo <> TIPO_SEP_CAMP_LARGHEZZA_FISSA Then
        For i = 0 To rsRowExample.fields.Count - 1
            s = s & rsRowExample(i).Name & mArrayTipoSeparatoreCampo(mTipoSeparatoreCampo)
        Next
         s = Left$(s, Len(s) - 1)
        txtEsempio = txtEsempio & s + vbCrLf
    End If
    
    Do While Not rsRowExample.EOF
        s = vbNullString
        For i = 0 To rsRowExample.fields.Count - 1
            
            s = s & applyFormatField(nullToString(rsRowExample(i).Value), rsRowExample(i))
        Next
        If mTipoSeparatoreCampo <> TIPO_SEP_CAMP_LARGHEZZA_FISSA Then
            s = Left$(s, Len(s) - 1)
        End If
        
        rsRowExample.MoveNext
        
        If rsRowExample.EOF Then
            txtEsempio = txtEsempio & s
        Else
            txtEsempio = txtEsempio & s + vbCrLf
        End If
    Loop
    
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If Not rsTable Is Nothing Then
        If rsTable.State = adStateOpen Then
            rsTable.Close
        End If
    End If
        
    If Not rsRowExample Is Nothing Then
        If rsRowExample.State = adStateOpen Then
            rsRowExample.Close
        End If
    End If
    
    Set rsTable = Nothing
    Set rsRowExample = Nothing
End Sub

Private Sub optSeparatore_Click(Index As Integer)
    mTipoSeparatoreCampo = Index
    
    txtSeparatore.Enabled = Index = TipoSeparatoreCampo.TIPO_SEP_CAMP_ALTRO
    txtSeparatore.BackColor = IIf(txtSeparatore.Enabled, vbWhite, BACKCOLOR_CAMPI_CALC)
    
    
End Sub

Private Sub TabStrip1_Click()
    
    pic(0).Visible = False
    pic(1).Visible = False
    pic(TabStrip1.SelectedItem.Index - 1).Visible = True
End Sub


Private Sub txtSeparatore_LostFocus()
    mArrayTipoSeparatoreCampo(TipoSeparatoreCampo.TIPO_SEP_CAMP_ALTRO) = txtSeparatore.Text
End Sub

Private Function nullToString(s As Variant) As String
    If IsNull(s) Then s = vbNullString
    nullToString = s
End Function


Private Sub esportazione()
    Dim f As Long
    Dim file_name As String
    Dim rs As ADODB.Recordset
    Dim s As String
    Dim i As Integer
    
    file_name = cmbTabelle.List(cmbTabelle.ListIndex)
    If file_name = vbNullString Then Exit Sub
    
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    Set rs.ActiveConnection = gDBC
    rs.Source = "select * from " & file_name & " where codComune='" & mCodComune & "'"
    rs.Open
    Set rs.ActiveConnection = Nothing
    
    
    prbEsportazione.Value = 0
    prbEsportazione.Min = 0
    prbEsportazione.Max = rs.RecordCount
    
    If mExportXML Then
        rs.save mPath_to_export & "\" & file_name & ".xml", adPersistXML
    Else
        f = FreeFile
        On Error Resume Next
        Kill mPath_to_export & "\" & file_name & ".txt"
        On Error GoTo esportazioneERR
        Open mPath_to_export & "\" & file_name & ".txt" For Output As #f

        If mTitoliColonna And mTipoSeparatoreCampo <> TIPO_SEP_CAMP_LARGHEZZA_FISSA Then
            For i = 0 To rs.fields.Count - 1
                s = s & applyFormatField(nullToString(rs(i).Name), rs(i))
            Next
            s = Left$(s, Len(s) - 1)
            Print #f, s
            prbEsportazione.Max = prbEsportazione.Max + 1
            prbEsportazione.Value = prbEsportazione.Value + 1
        End If
        
        Do While Not rs.EOF
            s = vbNullString
            For i = 0 To rs.fields.Count - 1
                s = s & applyFormatField(nullToString(rs(i).Value), rs(i))
            Next
            If mTipoSeparatoreCampo <> TIPO_SEP_CAMP_LARGHEZZA_FISSA Then
                s = Left$(s, Len(s) - 1)
            End If
            
            Print #f, s
            prbEsportazione.Value = prbEsportazione.Value + 1
            rs.MoveNext
            
        Loop
        prbEsportazione.Value = prbEsportazione.Max
        Close f
    End If
    
    rs.Close
    Exit Sub
esportazioneERR:

    MsgBox Error$, vbCritical, App.Title
End Sub


Private Sub txtSeparatoreDecimale_LostFocus()
    ArrayTipiDati(mTipoCampo).SeparatoreDecimale = txtSeparatoreDecimale.Text
End Sub

Private Sub txtSeparatoreData_LostFocus()
    ArrayTipiDati(mTipoCampo).SeparatoreData = txtSeparatoreData.Text
End Sub

Private Sub txtSeparatoreOra_LostFocus()
     ArrayTipiDati(mTipoCampo).SeparatoreOra = txtSeparatoreOra.Text
End Sub

Private Sub txtRiempimento_lostfocus()
    ArrayTipiDati(mTipoCampo).Riempimento = txtRiempimento.Text
End Sub

