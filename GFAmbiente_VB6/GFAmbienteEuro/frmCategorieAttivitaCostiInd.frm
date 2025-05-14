VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCategorieAttivitaCostiInd 
   Caption         =   "Form1"
   ClientHeight    =   10620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   ScaleHeight     =   10620
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmd 
      Caption         =   "&Chiudi"
      Height          =   480
      Left            =   11280
      TabIndex        =   1
      Top             =   9960
      Width           =   1410
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   360
      Left            =   105
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   10125
      Width           =   2280
   End
   Begin MSAdodcLib.Adodc ADO1 
      Height          =   360
      Left            =   4380
      Top             =   3330
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
      Caption         =   "Ado1"
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
      Height          =   480
      Left            =   9960
      TabIndex        =   2
      Top             =   9960
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   847
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy"
      Format          =   57737219
      UpDown          =   -1  'True
      CurrentDate     =   36967
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCategorieAttivitaCostiInd.frx":0000
      Height          =   9795
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12690
      _ExtentX        =   22384
      _ExtentY        =   17277
      _Version        =   393216
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "IdCategoriaAttivita"
         Caption         =   "IdCategoriaAttivita"
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
      BeginProperty Column02 
         DataField       =   "Descrizione"
         Caption         =   "Categoria"
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
         DataField       =   "Superficie"
         Caption         =   "Superficie"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "SuperficieCostiIndi"
         Caption         =   "Superf. Costi Ind."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "GettitoCosto1"
         Caption         =   "Gettito a 0,3"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """L."" ##,##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "GettitoCosto2"
         Caption         =   "Gettito a 0,4"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """L."" ##,##"
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
            ColumnWidth     =   1305,071
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   6254,929
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1920,189
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1590,236
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1440
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tot. Gettito"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   9870
      Width           =   975
   End
End
Attribute VB_Name = "frmCategorieAttivitaCostiInd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCodComune As String
Private sComune As String
Private mAnnoRiferimento As String


Sub TotaleGettito()
    Dim Totale As Double
    Dim i As Integer
    Dim CurrentCol As Integer
    Dim CurrentRow As Integer
    
    CurrentRow = 0
    CurrentCol = 0
    
    CurrentRow = DataGrid1.row
    CurrentCol = DataGrid1.col + 1
    
    On Error Resume Next
    Totale = 0
    For i = 0 To ADO1.Recordset.RecordCount - 1
        DataGrid1.row = i
        Totale = Totale + (0 & DataGrid1.Columns(DataGrid1.Columns.Count - 1).value) + _
                 (0 & DataGrid1.Columns(DataGrid1.Columns.Count - 2).value)
    
    Next

    If CurrentRow > -1 And CurrentCol < DataGrid1.Columns.Count Then
        DataGrid1.row = CurrentRow
        DataGrid1.col = CurrentCol
        
        Txt.Text = Format$(Totale, FORMATVALUTA)
   
        
    End If
        

End Sub




Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)

    If ColIndex <= 2 Then Exit Sub
    
    If (Round(DataGrid1.Columns(ColIndex).value, 2)) <> DataGrid1.Columns(ColIndex).value Then
       DataGrid1.Columns(ColIndex).value = Round(DataGrid1.Columns(ColIndex).value, 2)
    End If
    
    DataGrid1.Columns(DataGrid1.Columns.Count - 2).value = DataGrid1.Columns(ColIndex).value * 0.3
    DataGrid1.Columns(DataGrid1.Columns.Count - 1).value = (DataGrid1.Columns("Superficie").value - DataGrid1.Columns(ColIndex).value) * 0.4
     
End Sub

Private Sub DataGrid1_AfterUpdate()
    
    TotaleGettito
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        Sendkeys "{ENTER}", True
        DoEvents
    End If
End Sub


Private Sub DataGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    'I campi devono essere numerici
    If ColIndex = 2 Then Exit Sub
    If Not IsNumeric(DataGrid1.Tag) Then
        MsgBox "Il valore non è numerico", vbExclamation + vbOKOnly, App.Title
        Cancel = True
        Exit Sub
    ElseIf Conversion.CDbl(DataGrid1.Tag) > DataGrid1.Columns("Superficie").value Then
        MsgBox "Attenzione il valore non può essere maggiore di " & DataGrid1.Columns("Superficie").value, vbExclamation + vbOKOnly, App.Title
        Cancel = True
        Exit Sub
    End If
End Sub


Public Sub Carica(Optional sCodComune As String = vbNullString)
    Dim rs As Recordset
    Dim qry As ADODB.Command
    
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
    Set rs = gDBC.Execute("Select * from Comuni where CodIstat='" & mCodComune & "'")
    sComune = rs(0)
    
    Set rs = Nothing
    Set rs = DatiTecniciComune(mCodComune, mAnnoRiferimento)

    
    Set qry = gADOX.Procedures("InserisciCategorieAttivitaComunali").Command
    qry.Parameters.Refresh
    qry.Parameters(0) = mCodComune
    qry.Parameters(1) = mAnnoRiferimento
    qry.Execute
    
    CaricaDati mAnnoRiferimento
    
    Me.Show vbModal
End Sub


Private Sub DTPAnno_Change()
    mAnnoRiferimento = Year(DTPAnno.value)
    CaricaDati mAnnoRiferimento
End Sub


Private Sub Form_Activate()
    'Mi posiziono sulla griglia
    DataGrid1.SetFocus
End Sub

Private Sub Form_Load()
    
    MDIMain.mnuComprimi.Enabled = False
    Me.Caption = "Costi Indivisibili - Categorie Attivita del Comune: " & sComune

    ADO1.ConnectionString = gDBC

    
    DataGrid1.Left = 0
    DataGrid1.Top = 0
    DataGrid1.Height = Cmd.Top - 100
    DataGrid1.Width = Me.ScaleWidth
        
    DataGrid1.Columns(DataGrid1.Columns.Count - 1).NumberFormat = VALUTA_EURO & " " & FORMATVALUTA
    DataGrid1.Columns(DataGrid1.Columns.Count - 2).NumberFormat = VALUTA_EURO & " " & FORMATVALUTA
    
    DTPAnno.value = Format$(Now, "yyyy-mm-dd")
End Sub


Private Function CaricaDati(ByVal Anno As String) As Boolean
    
    CaricaDati = False
    
    On Error Resume Next
    ADO1.Recordset.Close
    On Error GoTo 0
    
    
    ADO1.CommandType = adCmdText
    ADO1.RecordSource = "SELECT CategorieAttivitaComunali.* from CategorieAttivitaComunali WHERE CategorieAttivitaComunali.CodComune='" & mCodComune & "' and CategorieAttivitaComunali.AnnoRiferimento='" & Anno & "' order by IdCategoriaAttivita"
    
    ADO1.Refresh

    CaricaDati = Not ADO1.Recordset.EOF
    
    If Not CaricaDati Then
        DataGrid1.RowHeight = 30
    Else
        DataGrid1.RowHeight = 270
        DataGrid1.Refresh
    End If
    
    TotaleGettito
    
End Function


Private Sub Form_Unload(Cancel As Integer)
    MDIMain.mnuComprimi.Enabled = True
End Sub



Private Sub Cmd_Click()
    Dim rs As ADODB.Recordset
    Dim qry As ADODB.Command
    

    Unload Me
      
    

End Sub


Private Sub DataGrid1_Change()
      ' Change is the easiest event to tap into.
      ' It is fired every time the cell text changes.

      ' Because you can only save the text when the cell is being edited,
      ' you have to store the value in another variable if you want
      ' to use that value in other events.

      ' Place the contents of the cell into the TAG property as a buffer.
      ' You could also use a variable to store the information.

      DataGrid1.Tag = GetCellText(DataGrid1)
End Sub







