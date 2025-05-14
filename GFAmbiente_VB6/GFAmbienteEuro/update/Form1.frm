VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUpdateDB 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Data Base - GFAmbiente"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   6870
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView ListDA 
      Height          =   1155
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   2037
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NomeDB"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.CommandButton cmdPathA 
      Caption         =   "Path"
      Height          =   315
      Left            =   3600
      TabIndex        =   11
      Top             =   1560
      Width           =   3135
   End
   Begin VB.CommandButton cmdPathDA 
      Caption         =   "Path"
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Avvia"
      Height          =   495
      Index           =   1
      Left            =   3960
      TabIndex        =   5
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Chiudi"
      Height          =   495
      Index           =   0
      Left            =   5400
      TabIndex        =   4
      Top             =   5400
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar PRBar 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin MSComctlLib.ListView lvwComuni 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3836
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListA 
      Height          =   1155
      Left            =   3600
      TabIndex        =   13
      Top             =   360
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   2037
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NomeDB"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Base a cui copiare"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Base da cui copiare"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ó"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   2520
      TabIndex        =   7
      Top             =   5400
      Width           =   165
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GFAmbiente - CopyRight   2002-2003 Analisi e Calcolo Tariffa Tarsu"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      Width           =   3735
   End
   Begin VB.Label lblInfoTable 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   6555
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lista Comuni"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   6495
   End
End
Attribute VB_Name = "frmUpdateDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gCON_OLD As ADODB.Connection
Dim gCON_NEW As ADODB.Connection
Dim gCAT  As ADOX.Catalog
Dim mpath_NEW As String
Dim mpath_OLD As String
Dim numPass As Integer

Const DATABASE_OLD = "\StudioAmbiente*.OLD"
Const DATABASE_NEW = "\StudioAmbiente*.MDB"



Private Type Table
    TableName As String
    KEY1 As String
    KEY2 As String
End Type



Dim listTable() As Table

Private Sub Tables_List()
    Dim adTab  As ADOX.Table
    
'    listTable.Clear
'    For Each adTab In gCAT.Tables
'        If adTab.Type = "TABLE" Then
'            listTable.AddItem adTab.Name
'        End If
'    Next
End Sub

Private Sub Comuni_List()
    Dim rs As ADODB.Recordset
    Dim Item As ListItem
    lvwComuni.ListItems.Clear
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.LockType = adLockBatchOptimistic
    rs.CursorType = adOpenStatic
    Set rs.ActiveConnection = gCON_OLD
    rs.Open "select * from CostoComunale ac inner join Comuni c on c.codIstat=ac.CodComune order by ac.codComune,annoriferimento"
    Set rs.ActiveConnection = Nothing
    
    lvwComuni.ColumnHeaders.Add , , "CodIstat", lvwComuni.Width * 0.15
    lvwComuni.ColumnHeaders.Add , , "Anno", lvwComuni.Width * 0.2
    lvwComuni.ColumnHeaders.Add , , "Comune", lvwComuni.Width * 0.6
    
    Do While Not rs.EOF
        Set Item = lvwComuni.ListItems.Add(, "|" & rs("CodComune").Value & "|" & rs("AnnoRiferimento").Value, rs("CodComune").Value)
        Item.ListSubItems.Add , rs("AnnoRiferimento").Name, rs("AnnoRiferimento").Value
        Item.ListSubItems.Add , rs("Comune").Name, rs("Comune").Value
        Item.Checked = True
        rs.MoveNext
    Loop
End Sub

Function ListField(rs As Recordset) As String
    Dim i As Long
    Dim s As String
    'ritorna una lista fields separati da una virgola
    For i = 0 To rs.Fields.Count - 1
        s = s & rs.Fields(i).Name & ","
    Next
    s = Left$(s, Len(s) - Len(","))
    ListField = s
End Function


Function ListFieldValue(rs As Recordset) As String
    Dim i As Long
    Dim s As String
    'ritorna una lista fields separati da una virgola
    For i = 0 To rs.Fields.Count - 1
        
        If IsNull(rs.Fields(i)) Then
            s = s & "Null,"
        Else
            Select Case rs.Fields(i).Type
                'Testo
                Case DataTypeEnum.adChar, DataTypeEnum.adLongVarChar, _
                     DataTypeEnum.adVarChar, DataTypeEnum.adBSTR, adVarWChar
                     s = s & "'" & Replace(rs.Fields(i), "'", "''") & "',"
                'Numerico
                Case DataTypeEnum.adBigInt, DataTypeEnum.adCurrency, _
                     DataTypeEnum.adDecimal, DataTypeEnum.adDouble, DataTypeEnum.adInteger, _
                     DataTypeEnum.adNumeric, DataTypeEnum.adSingle, DataTypeEnum.adSmallInt, _
                     DataTypeEnum.adTinyInt
                     s = s & "0" & Replace(CStr(rs.Fields(i)), ",", ".") & ","
                'Logico
                Case DataTypeEnum.adBoolean
                     s = s & IIf(CBool(rs.Fields(i)), 1, 0) & ","
                'Data
                Case DataTypeEnum.adDate, DataTypeEnum.adDBDate, DataTypeEnum.adDBTimeStamp
                     s = s & "'" & Format$(rs.Fields(i), "yyyy-mm-dd hh:nn:ss") & "',"
                'in tutti gli altri casi
                Case Else
                     s = s & "'" & rs.Fields(i) & "',"
            End Select
        End If
    Next
    s = Left$(s, Len(s) - Len(","))
    
    ListFieldValue = s
End Function

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            StartEsportazione
    End Select
End Sub


Private Sub CaricaTabelleOLDDB()
    Erase listTable
    ReDim listTable(17)
    listTable(0).TableName = "AnagraficaComunale"
    listTable(0).KEY1 = "codcomune"
    listTable(0).KEY2 = "annoriferimento"
    
    listTable(1).TableName = "CostoComunale"
    listTable(1).KEY1 = "codcomune"
    listTable(1).KEY2 = "annoriferimento"
    
    listTable(2).TableName = "DatiTecniciComune"
    listTable(2).KEY1 = "codcomune"
    listTable(2).KEY2 = "annoriferimento"

    listTable(3).TableName = "Accantonamenti"
    listTable(3).KEY1 = "codcomune"
    listTable(3).KEY2 = "annoriferimento"
    
    listTable(4).TableName = "Ammortamenti"
    listTable(4).KEY1 = "codcomune"
    listTable(4).KEY2 = "annoriferimento"
    
    listTable(5).TableName = "Attivazione"
    listTable(5).KEY1 = "codcomune"
    listTable(5).KEY2 = ""
    
    listTable(6).TableName = "BonusUtenzeAttivita"
    listTable(6).KEY1 = "codcomune"
    listTable(6).KEY2 = "annoriferimento"
    
    listTable(7).TableName = "BonusUtenzeDomestiche"
    listTable(7).KEY1 = "codcomune"
    listTable(7).KEY2 = "annoriferimento"
    
    listTable(8).TableName = "CategorieAttivitaComunali"
    listTable(8).KEY1 = "codcomune"
    listTable(8).KEY2 = "annoriferimento"
    
    listTable(9).TableName = "DatiFinaliCalcoloTariffa"
    listTable(9).KEY1 = "codcomune"
    listTable(9).KEY2 = "annoriferimento"
    
    listTable(10).TableName = "DatiFinaliCalcoloTarsu507"
    listTable(10).KEY1 = "codcomune"
    listTable(10).KEY2 = "annoriferimento"
    
    listTable(11).TableName = "ParametriSistema"
    listTable(11).KEY1 = "codIstat"

    
    listTable(12).TableName = "PersonaleImpiegatoNelComune"
    listTable(12).KEY1 = "codcomune"
    listTable(12).KEY2 = "annoriferimento"
    
    listTable(13).TableName = "RetribuzioneDelPersonale"
    listTable(13).KEY1 = "codcomune"
    listTable(13).KEY2 = "annoriferimento"
    
    listTable(14).TableName = "StatisticheComunali"
    listTable(14).KEY1 = "codcomune"
    listTable(14).KEY2 = "annoriferimento"
    
    listTable(15).TableName = "UtenzeDomesticheComunali"
    listTable(15).KEY1 = "codcomune"
    listTable(15).KEY2 = "annoriferimento"
    
    
    listTable(16).TableName = "ValoriCoefficientiCAComunali"
    listTable(16).KEY1 = "codcomune"
    listTable(16).KEY2 = "annoriferimento"
    
    
    listTable(17).TableName = "ValoriCoefficientiUDComunali"
    listTable(17).KEY1 = "codcomune"
    listTable(17).KEY2 = "annoriferimento"
    
    
    
    
    
End Sub

Private Sub CaricaTabelleNEWDB()
    
    Erase listTable
    ReDim listTable(19)
    listTable(0).TableName = "AnagraficaComunale"
    listTable(0).KEY1 = "codcomune"
    listTable(0).KEY2 = "annoriferimento"
    
    listTable(1).TableName = "CostoComunale"
    listTable(1).KEY1 = "codcomune"
    listTable(1).KEY2 = "annoriferimento"
    
    listTable(2).TableName = "DatiTecniciComune"
    listTable(2).KEY1 = "codcomune"
    listTable(2).KEY2 = "annoriferimento"

    listTable(3).TableName = "Accantonamenti"
    listTable(3).KEY1 = "codcomune"
    listTable(3).KEY2 = "annoriferimento"
    
    listTable(4).TableName = "Ammortamenti"
    listTable(4).KEY1 = "codcomune"
    listTable(4).KEY2 = "annoriferimento"
    
    listTable(5).TableName = "Attivazione"
    listTable(5).KEY1 = "codcomune"
    listTable(5).KEY2 = ""
    
    listTable(6).TableName = "BonusUtenzeAttivita"
    listTable(6).KEY1 = "codcomune"
    listTable(6).KEY2 = "annoriferimento"
    
    listTable(7).TableName = "BonusUtenzeDomestiche"
    listTable(7).KEY1 = "codcomune"
    listTable(7).KEY2 = "annoriferimento"
    
    listTable(8).TableName = "CategorieAttivitaComunali"
    listTable(8).KEY1 = "codcomune"
    listTable(8).KEY2 = "annoriferimento"
    
    listTable(9).TableName = "DatiFinaliCalcoloTariffa"
    listTable(9).KEY1 = "codcomune"
    listTable(9).KEY2 = "annoriferimento"
    
    listTable(10).TableName = "DatiFinaliCalcoloTarsu507"
    listTable(10).KEY1 = "codcomune"
    listTable(10).KEY2 = "annoriferimento"
    
    listTable(11).TableName = "ParametriSistema"
    listTable(11).KEY1 = "codIstat"
    listTable(11).KEY2 = "annoriferimento"
    
    listTable(12).TableName = "PersonaleImpiegatoNelComune"
    listTable(12).KEY1 = "codcomune"
    listTable(12).KEY2 = "annoriferimento"
    
    listTable(13).TableName = "RetribuzioneDelPersonale"
    listTable(13).KEY1 = "codcomune"
    listTable(13).KEY2 = "annoriferimento"
    
    listTable(14).TableName = "StatisticheComunali"
    listTable(14).KEY1 = "codcomune"
    listTable(14).KEY2 = "annoriferimento"
    
    listTable(15).TableName = "UtenzeDomesticheComunali"
    listTable(15).KEY1 = "codcomune"
    listTable(15).KEY2 = "annoriferimento"
    
    
    listTable(16).TableName = "ValoriCoefficientiCAComunali"
    listTable(16).KEY1 = "codcomune"
    listTable(16).KEY2 = "annoriferimento"
    
    
    listTable(17).TableName = "ValoriCoefficientiUDComunali"
    listTable(17).KEY1 = "codcomune"
    listTable(17).KEY2 = "annoriferimento"
    
    listTable(18).TableName = "ValoriCoefficientiIQSCAComunali"
    listTable(18).KEY1 = "codcomune"
    listTable(18).KEY2 = "annoriferimento"
        
    listTable(19).TableName = "ValoriCoefficienti507UDComunali"
    listTable(19).KEY1 = "codcomune"
    listTable(19).KEY2 = "annoriferimento"


End Sub

Private Sub CaricaTabelleNEWDB_WITHSIMU()
    
    Erase listTable
    ReDim listTable(22)
    listTable(0).TableName = "AnagraficaComunale"
    listTable(0).KEY1 = "codcomune"
    listTable(0).KEY2 = "annoriferimento"
    
    listTable(1).TableName = "CostoComunale"
    listTable(1).KEY1 = "codcomune"
    listTable(1).KEY2 = "annoriferimento"
    
    listTable(2).TableName = "DatiTecniciComune"
    listTable(2).KEY1 = "codcomune"
    listTable(2).KEY2 = "annoriferimento"

    listTable(3).TableName = "Accantonamenti"
    listTable(3).KEY1 = "codcomune"
    listTable(3).KEY2 = "annoriferimento"
    
    listTable(4).TableName = "Ammortamenti"
    listTable(4).KEY1 = "codcomune"
    listTable(4).KEY2 = "annoriferimento"
    
    listTable(5).TableName = "Attivazione"
    listTable(5).KEY1 = "codcomune"
    listTable(5).KEY2 = ""
    
    listTable(6).TableName = "BonusUtenzeAttivita"
    listTable(6).KEY1 = "codcomune"
    listTable(6).KEY2 = "annoriferimento"
    
    listTable(7).TableName = "BonusUtenzeDomestiche"
    listTable(7).KEY1 = "codcomune"
    listTable(7).KEY2 = "annoriferimento"
    
    listTable(8).TableName = "CategorieAttivitaComunali"
    listTable(8).KEY1 = "codcomune"
    listTable(8).KEY2 = "annoriferimento"
    
    listTable(9).TableName = "DatiFinaliCalcoloTariffa"
    listTable(9).KEY1 = "codcomune"
    listTable(9).KEY2 = "annoriferimento"
    
    listTable(10).TableName = "DatiFinaliCalcoloTarsu507"
    listTable(10).KEY1 = "codcomune"
    listTable(10).KEY2 = "annoriferimento"
    
    listTable(11).TableName = "ParametriSistema"
    listTable(11).KEY1 = "codIstat"
    listTable(11).KEY2 = "annoriferimento"
    
    listTable(12).TableName = "PersonaleImpiegatoNelComune"
    listTable(12).KEY1 = "codcomune"
    listTable(12).KEY2 = "annoriferimento"
    
    listTable(13).TableName = "RetribuzioneDelPersonale"
    listTable(13).KEY1 = "codcomune"
    listTable(13).KEY2 = "annoriferimento"
    
    listTable(14).TableName = "StatisticheComunali"
    listTable(14).KEY1 = "codcomune"
    listTable(14).KEY2 = "annoriferimento"
    
    listTable(15).TableName = "UtenzeDomesticheComunali"
    listTable(15).KEY1 = "codcomune"
    listTable(15).KEY2 = "annoriferimento"
    
    
    listTable(16).TableName = "ValoriCoefficientiCAComunali"
    listTable(16).KEY1 = "codcomune"
    listTable(16).KEY2 = "annoriferimento"
    
    
    listTable(17).TableName = "ValoriCoefficientiUDComunali"
    listTable(17).KEY1 = "codcomune"
    listTable(17).KEY2 = "annoriferimento"
    
    listTable(18).TableName = "ValoriCoefficientiIQSCAComunali"
    listTable(18).KEY1 = "codcomune"
    listTable(18).KEY2 = "annoriferimento"
        
    listTable(19).TableName = "ValoriCoefficienti507UDComunali"
    listTable(19).KEY1 = "codcomune"
    listTable(19).KEY2 = "annoriferimento"


    listTable(20).TableName = "ChartData"
    listTable(20).KEY1 = ""
    listTable(20).KEY2 = ""
    
    listTable(21).TableName = "ChartDefault"
    listTable(21).KEY1 = ""
    listTable(21).KEY2 = ""

    listTable(22).TableName = "Simulazioni"
    listTable(22).KEY1 = "codComune"
    listTable(22).KEY2 = "idSimulazione"

End Sub


Private Sub CaricaTabelleNEWDB_WITHRACCOLTA()
    
    Erase listTable
    ReDim listTable(30)
    listTable(0).TableName = "AnagraficaComunale"
    listTable(0).KEY1 = "codcomune"
    listTable(0).KEY2 = "annoriferimento"
    
    listTable(1).TableName = "CostoComunale"
    listTable(1).KEY1 = "codcomune"
    listTable(1).KEY2 = "annoriferimento"
    
    listTable(2).TableName = "DatiTecniciComune"
    listTable(2).KEY1 = "codcomune"
    listTable(2).KEY2 = "annoriferimento"

    listTable(3).TableName = "Accantonamenti"
    listTable(3).KEY1 = "codcomune"
    listTable(3).KEY2 = "annoriferimento"
    
    listTable(4).TableName = "Ammortamenti"
    listTable(4).KEY1 = "codcomune"
    listTable(4).KEY2 = "annoriferimento"
    
    listTable(5).TableName = "Attivazione"
    listTable(5).KEY1 = "codcomune"
    listTable(5).KEY2 = ""
    
    listTable(6).TableName = "BonusUtenzeAttivita"
    listTable(6).KEY1 = "codcomune"
    listTable(6).KEY2 = "annoriferimento"
    
    listTable(7).TableName = "BonusUtenzeDomestiche"
    listTable(7).KEY1 = "codcomune"
    listTable(7).KEY2 = "annoriferimento"
    
    listTable(8).TableName = "CategorieAttivita"
    listTable(8).KEY1 = "codcomune"
    listTable(8).KEY2 = "annoriferimento"
    
    
    listTable(8).TableName = "CategorieAttivitaComunali"
    listTable(8).KEY1 = "codcomune"
    listTable(8).KEY2 = "annoriferimento"
    
    listTable(9).TableName = "DatiFinaliCalcoloTariffa"
    listTable(9).KEY1 = "codcomune"
    listTable(9).KEY2 = "annoriferimento"
    
    listTable(10).TableName = "DatiFinaliCalcoloTarsu507"
    listTable(10).KEY1 = "codcomune"
    listTable(10).KEY2 = "annoriferimento"
    
    listTable(11).TableName = "ParametriSistema"
    listTable(11).KEY1 = "codIstat"
    listTable(11).KEY2 = "annoriferimento"
    
    listTable(12).TableName = "PersonaleImpiegatoNelComune"
    listTable(12).KEY1 = "codcomune"
    listTable(12).KEY2 = "annoriferimento"
    
    listTable(13).TableName = "RetribuzioneDelPersonale"
    listTable(13).KEY1 = "codcomune"
    listTable(13).KEY2 = "annoriferimento"
    
    listTable(14).TableName = "StatisticheComunali"
    listTable(14).KEY1 = "codcomune"
    listTable(14).KEY2 = "annoriferimento"
    
    listTable(15).TableName = "UtenzeDomesticheComunali"
    listTable(15).KEY1 = "codcomune"
    listTable(15).KEY2 = "annoriferimento"
    
    
    listTable(16).TableName = "ValoriCoefficientiCAComunali"
    listTable(16).KEY1 = "codcomune"
    listTable(16).KEY2 = "annoriferimento"
    
    
    listTable(17).TableName = "ValoriCoefficientiUDComunali"
    listTable(17).KEY1 = "codcomune"
    listTable(17).KEY2 = "annoriferimento"
    
    listTable(18).TableName = "ValoriCoefficientiIQSCAComunali"
    listTable(18).KEY1 = "codcomune"
    listTable(18).KEY2 = "annoriferimento"
        
    listTable(19).TableName = "ValoriCoefficienti507UDComunali"
    listTable(19).KEY1 = "codcomune"
    listTable(19).KEY2 = "annoriferimento"


    listTable(20).TableName = "ChartData"
    listTable(20).KEY1 = ""
    listTable(20).KEY2 = ""
    
    listTable(21).TableName = "ChartDefault"
    listTable(21).KEY1 = ""
    listTable(21).KEY2 = ""

    listTable(22).TableName = "Simulazioni"
    listTable(22).KEY1 = "codComune"
    listTable(22).KEY2 = "idSimulazione"

    listTable(23).TableName = "AnagraficaRaccolta"
    listTable(23).KEY1 = "CodComune"
    listTable(23).KEY2 = "annoriferimento"

    listTable(24).TableName = "ContenitoriComunali"
    listTable(24).KEY1 = "CodComune"
    listTable(24).KEY2 = "AnnoRiferimento"
    
    listTable(25).TableName = "ContenitoriInDotazioneCAComunali"
    listTable(25).KEY1 = "CodComune"
    listTable(25).KEY2 = "AnnoRiferimento"
    
    listTable(26).TableName = "ContenitoriInDotazioneUDComunali"
    listTable(26).KEY1 = "CodComune"
    listTable(26).KEY2 = "AnnoRiferimento"
    
    listTable(27).TableName = "ContenitoriSvuotamentiCAComunali"
    listTable(27).KEY1 = "CodComune"
    listTable(27).KEY2 = "AnnoRiferimento"
    
    listTable(28).TableName = "ContenitoriSvuotamentiUDComunali"
    listTable(28).KEY1 = "CodComune"
    listTable(28).KEY2 = "AnnoRiferimento"
    
    listTable(29).TableName = "Esportazione"
    listTable(29).KEY1 = "CodComune"
    listTable(29).KEY2 = "NomeTabella"
    
    listTable(30).TableName = "PercentualePesiRDComunali"
    listTable(30).KEY1 = "CodComune"
    listTable(30).KEY2 = "AnnoRiferimento"
    

End Sub

Private Sub CaricaTabelleNEWDB_WITHCHART()
    
    Erase listTable
    ReDim listTable(31)
    listTable(0).TableName = "AnagraficaComunale"
    listTable(0).KEY1 = "codcomune"
    listTable(0).KEY2 = "annoriferimento"
    
    listTable(1).TableName = "CostoComunale"
    listTable(1).KEY1 = "codcomune"
    listTable(1).KEY2 = "annoriferimento"
    
    listTable(2).TableName = "DatiTecniciComune"
    listTable(2).KEY1 = "codcomune"
    listTable(2).KEY2 = "annoriferimento"

    listTable(3).TableName = "Accantonamenti"
    listTable(3).KEY1 = "codcomune"
    listTable(3).KEY2 = "annoriferimento"
    
    listTable(4).TableName = "Ammortamenti"
    listTable(4).KEY1 = "codcomune"
    listTable(4).KEY2 = "annoriferimento"
    
    listTable(5).TableName = "Attivazione"
    listTable(5).KEY1 = "codcomune"
    listTable(5).KEY2 = ""
    
    listTable(6).TableName = "BonusUtenzeAttivita"
    listTable(6).KEY1 = "codcomune"
    listTable(6).KEY2 = "annoriferimento"
    
    listTable(7).TableName = "BonusUtenzeDomestiche"
    listTable(7).KEY1 = "codcomune"
    listTable(7).KEY2 = "annoriferimento"
    
    listTable(8).TableName = "CategorieAttivitaComunali"
    listTable(8).KEY1 = "codcomune"
    listTable(8).KEY2 = "annoriferimento"
    
    listTable(9).TableName = "DatiFinaliCalcoloTariffa"
    listTable(9).KEY1 = "codcomune"
    listTable(9).KEY2 = "annoriferimento"
    
    listTable(10).TableName = "DatiFinaliCalcoloTarsu507"
    listTable(10).KEY1 = "codcomune"
    listTable(10).KEY2 = "annoriferimento"
    
    listTable(11).TableName = "ParametriSistema"
    listTable(11).KEY1 = "codIstat"
    listTable(11).KEY2 = "annoriferimento"
    
    listTable(12).TableName = "PersonaleImpiegatoNelComune"
    listTable(12).KEY1 = "codcomune"
    listTable(12).KEY2 = "annoriferimento"
    
    listTable(13).TableName = "RetribuzioneDelPersonale"
    listTable(13).KEY1 = "codcomune"
    listTable(13).KEY2 = "annoriferimento"
    
    listTable(14).TableName = "StatisticheComunali"
    listTable(14).KEY1 = "codcomune"
    listTable(14).KEY2 = "annoriferimento"
    
    listTable(15).TableName = "UtenzeDomesticheComunali"
    listTable(15).KEY1 = "codcomune"
    listTable(15).KEY2 = "annoriferimento"
    
    
    listTable(16).TableName = "ValoriCoefficientiCAComunali"
    listTable(16).KEY1 = "codcomune"
    listTable(16).KEY2 = "annoriferimento"
    
    
    listTable(17).TableName = "ValoriCoefficientiUDComunali"
    listTable(17).KEY1 = "codcomune"
    listTable(17).KEY2 = "annoriferimento"
    
    listTable(18).TableName = "ValoriCoefficientiIQSCAComunali"
    listTable(18).KEY1 = "codcomune"
    listTable(18).KEY2 = "annoriferimento"
        
    listTable(19).TableName = "ValoriCoefficienti507UDComunali"
    listTable(19).KEY1 = "codcomune"
    listTable(19).KEY2 = "annoriferimento"


    listTable(20).TableName = "ChartData"
    listTable(20).KEY1 = ""
    listTable(20).KEY2 = ""
    
    listTable(21).TableName = "ChartDefault"
    listTable(21).KEY1 = ""
    listTable(21).KEY2 = ""

End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not gCON_OLD Is Nothing Then
        gCON_OLD.Close
        Set gCON_OLD = Nothing
    End If
    If Not gCON_NEW Is Nothing Then
        gCON_NEW.Close
        Set gCON_NEW = Nothing
    End If
        
End Sub


Private Sub StartEsportazione()
    Dim Item As ListItem
    Dim i As Long
    Dim rs As ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Dim sHead As String
    Dim sKey() As String
    Dim sWhere As String
    
    On Error GoTo StartEsportazioneERR
    
    gCON_NEW.BeginTrans
    
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockBatchOptimistic
    
    Set rs1 = CreateObject("ADODB.Recordset")
    rs1.CursorLocation = adUseClient
    rs1.CursorType = adOpenStatic
    rs1.MaxRecords = 1
    rs1.LockType = adLockBatchOptimistic
     
    'AnagraficaComunale
    PRBar.Min = 0
    PRBar.Value = 0
    
    Dim attivitaPersonalizzate As Boolean
    attivitaPersonalizzate = False
    For Each Item In lvwComuni.ListItems
        If Item.Checked Then
            lblInfoTable = "Comune:" & Item.ListSubItems("Comune") & " Anno:" & Item.ListSubItems("AnnoRiferimento")
            lblInfoTable.Refresh
            sKey = Split(Mid$(Item.Key, 2), "|")
            PRBar.Max = ContaComuni * (UBound(listTable) + 1)
            
            
            
            For i = LBound(listTable) To UBound(listTable)
                'lettura vecchi dati
                Set rs.ActiveConnection = gCON_OLD
                
                sWhere = " where "
                If listTable(i).KEY1 <> "" Then
                    sWhere = sWhere & " " & listTable(i).KEY1 & "='" & sKey(0) & "' and "
                End If
                If listTable(i).KEY2 <> "" Then
                    sWhere = sWhere & " " & listTable(i).KEY2 & "='" & sKey(1) & "' and "
                End If
                
                sWhere = Left$(sWhere, Len(sWhere) - Len(" and "))
                
                rs.Open "select * from " & listTable(i).TableName & " " & sWhere
                Set rs.ActiveConnection = Nothing
                
                If listTable(i).TableName = "CategorieAttivitaComunali" And attivitaPersonalizzate = False Then
                    attivitaPersonalizzate = True
                    CaricaCategorieAttivitaPersonalizzate
                End If
                
                
                If Not rs.EOF And listTable(i).TableName <> "ChartData" Then
                    'lettura nuovi dati per struttura tabella
                    Set rs1.ActiveConnection = gCON_NEW
                    rs1.Open "select * from " & listTable(i).TableName & " " & sWhere
                    Set rs1.ActiveConnection = Nothing
                    
                    If rs1.EOF Then
                        sHead = "insert into " & listTable(i).TableName & " (" & ListField(rs) & ")"
                        Do While Not rs.EOF
                            gCON_NEW.Execute (sHead & " " & " values (" & ListFieldValue(rs) & ")")
                            rs.MoveNext
                        Loop
                    End If
                    
                End If
                PRBar.Value = PRBar.Value + 1
                If rs.State = adStateOpen Then rs.Close
                If rs1.State = adStateOpen Then rs1.Close
            Next
        End If
    Next
    PRBar.Value = PRBar.Max
    MsgBox "importazione dati terminata!!!!"

    'sistema il tot_csl
    Set rs = gCON_NEW.Execute("select 1 from FamigliaAmmortamenti where codFamigliaAmmortamento=6")
    If Not rs.EOF Then
        gCON_NEW.Execute ("update CostoComunale set tot_csl = csl where  tot_csl=0 and csl>0 ")
        rs.Close
    End If
    Set rs = Nothing
    gCON_NEW.CommitTrans
    Exit Sub
StartEsportazioneERR:

    If Not gCON_NEW Is Nothing Then gCON_NEW.RollbackTrans
    MsgBox listTable(i).TableName & ": " & Error$, vbCritical, App.Title
End Sub


Sub CaricaCategorieAttivitaPersonalizzate()
    Dim rs As ADODB.Recordset
    Set rs = CreateObject("ADODB.Recordset")
    
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    
    On Error GoTo CaricaCategorieAttivitaPersonalizzateERR
    Set rs.ActiveConnection = gCON_OLD
    rs.Open "select * from CategorieAttivita where System=false"
 
    Do While Not rs.EOF
        gCON_NEW.Execute ("Insert into CategorieAttivita (CodAttivita,Descrizione,Istat,System) values (" & _
                rs("CodAttivita").Value & "," & _
             "'" & rs("Descrizione").Value & "'," & _
             "'" & rs("Istat").Value & "'," & _
                rs("System").Value & ")")
        rs.MoveNext
    Loop
    
    rs.Close
    rs.Open "SELECT CoefficientiCategorieAttivita.* " & _
            "FROM CategorieAttivita INNER JOIN CoefficientiCategorieAttivita ON " & _
            "CategorieAttivita.CodAttivita = CoefficientiCategorieAttivita.CodTipologiaAttivita " & _
            "WHERE (((CategorieAttivita.System)=false))"
 
 
    Do While Not rs.EOF
        gCON_NEW.Execute ("Insert into CoefficientiCategorieAttivita (IdCoefficiente,CodTipologiaAttivita,CodParGeografica,CodRange,ParteFissa,MinK,MaxK) values (" & _
                rs("IdCoefficiente").Value & "," & _
              rs("CodTipologiaAttivita").Value & "," & _
             "'" & rs("CodParGeografica").Value & "'," & _
              "'" & rs("CodRange").Value & "'," & _
                rs("ParteFissa").Value & "," & _
                rs("MinK").Value & "," & _
                       rs("MaxK").Value & ")")
        rs.MoveNext
    Loop
    
    On Error Resume Next
    rs.Close
CaricaCategorieAttivitaPersonalizzateERR:
    
End Sub

Function ContaComuni() As Integer
    Dim i As Integer
    Dim ii As Integer
    
    For i = 1 To lvwComuni.ListItems.Count
        If lvwComuni.ListItems(i).Checked Then
            ii = ii + 1
        End If
    Next
    
    ContaComuni = ii
End Function


Sub ApriDBOLD(s As String)
    Dim password As String
    numPass = 0
    If Not gCON_OLD Is Nothing Then
        gCON_OLD.Close
        Set gCON_OLD = Nothing
    End If

    Set gCON_OLD = CreateObject("ADODB.Connection")
    
    Do
        password = getPassword(numPass)
        On Error Resume Next
        gCON_OLD.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & s & ";Jet OLEDB:Database Password=" + password
        gCON_OLD.Open
        If gCON_OLD.Errors.Count > 0 Then
            If (gCON_OLD.Errors(0).NativeError <> -124782449) Then
                On Error GoTo 0
                Err.Raise gCON_OLD.Errors(0).NativeError, "OPEN_DB", gCON_OLD.Errors(0).Description
            End If
        End If
    Loop While password <> "" And gCON_OLD.State <> adStateOpen


End Sub

Sub ApriDBNEW(s As String)
    Dim password As String
    numPass = 0
    If Not gCON_NEW Is Nothing Then
        gCON_NEW.Close
        Set gCON_NEW = Nothing
    End If

    Set gCON_NEW = CreateObject("ADODB.Connection")
    
    Do
        password = getPassword(numPass)
        On Error Resume Next
        gCON_NEW.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & s & ";Jet OLEDB:Database Password=" + password
        
        gCON_NEW.Open
        If gCON_NEW.Errors.Count > 0 Then
            If (gCON_NEW.Errors(0).NativeError <> -124782449) Then
                On Error GoTo 0
                Err.Raise gCON_NEW.Errors(0).NativeError, "OPEN_DB", gCON_NEW.Errors(0).Description
            End If
        End If
    Loop While password <> "" And gCON_NEW.State <> adStateOpen

End Sub


Sub DIRDataBase(lvw As ListView, s As String)
  Dim dbName As String
  
  
  dbName = Dir$(s)
  lvw.ListItems.Clear
  
  Do While dbName <> vbNullString
    lvw.ListItems.Add , dbName, dbName
    dbName = Dir$
  Loop
End Sub


Private Sub cmdPathDA_Click()
    Dim s As String
    s = App.Path
    mpath_OLD = s
    
    DIRDataBase ListDA, mpath_OLD & DATABASE_OLD
    
End Sub


Private Sub cmdPathA_Click()
    Dim s As String
    
    s = App.Path
    mpath_NEW = s
    DIRDataBase ListA, mpath_NEW & DATABASE_NEW
End Sub

Private Sub ListA_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    ApriDBNEW mpath_NEW & "\" & Item.Key
End Sub

Private Sub ListDA_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim rs As ADODB.Recordset
    ApriDBOLD mpath_OLD & "\" & Item.Key
    Comuni_List
    
    On Error GoTo CaricaDatiVecchi
    Set rs = gCON_OLD.Execute("select 1 from ValoriCoefficienti507UDComunali")
    On Error GoTo CaricaDatiDBSenzaChart
    Set rs = gCON_OLD.Execute("select 1 from ChartData")
    On Error GoTo CaricaDatiConChart
    Set rs = gCON_OLD.Execute("select 1 from Simulazioni")
    On Error GoTo CaricaTabelleNEWDB_WITHSIMU
    Set rs = gCON_OLD.Execute("select 1 from MaterialeRaccolta")
    CaricaTabelleNEWDB_WITHRACCOLTA
    
    
    Exit Sub

CaricaDatiVecchi:
    CaricaTabelleOLDDB
    Exit Sub
CaricaDatiDBSenzaChart:
    CaricaTabelleNEWDB
    Exit Sub
CaricaDatiConChart:
    CaricaTabelleNEWDB_WITHCHART
    Exit Sub
CaricaTabelleNEWDB_WITHSIMU:
    CaricaTabelleNEWDB_WITHSIMU
End Sub




