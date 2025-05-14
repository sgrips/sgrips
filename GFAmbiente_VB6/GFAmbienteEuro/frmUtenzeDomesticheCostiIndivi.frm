VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmUtenzeDomesticheCostiIndivi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Costi Indivisibili Utenze Domestiche"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "&Chiudi"
      Height          =   495
      Left            =   9885
      TabIndex        =   0
      Top             =   5205
      Width           =   1425
   End
   Begin MSComCtl2.DTPicker DTPAnno 
      Height          =   405
      Left            =   8520
      TabIndex        =   1
      Top             =   5265
      Width           =   1155
      _ExtentX        =   2037
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
      Format          =   57737219
      UpDown          =   -1  'True
      CurrentDate     =   36967
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmUtenzeDomesticheCostiIndivi.frx":0000
      Height          =   4980
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   8784
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "CodComune"
         Caption         =   "CodComune"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "IdNumeriComponenti"
         Caption         =   "ID Componente"
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
         DataField       =   "Utenza"
         Caption         =   "Utenza"
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
            Format          =   "#.##0"
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
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "GettitoCosti1"
         Caption         =   "Gettito a 0,30 €"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """L."" #,##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "GettitoCosti2"
         Caption         =   "Gettito a 0,4 €"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """L."" #,##"
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
            ColumnWidth     =   2099,906
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   2954,835
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   2069,858
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   2204,788
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   2174,74
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADO1 
      Height          =   330
      Left            =   4200
      Top             =   5520
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   1
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
      DataSourceName  =   "GFAmbiente"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Gettito"
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
      Left            =   90
      TabIndex        =   4
      Top             =   5280
      Width           =   585
   End
End
Attribute VB_Name = "frmUtenzeDomesticheCostiIndivi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCodComune As String
Private mAnnoRiferimento As String

Private sComune As String
 

Public Sub Carica(Optional sCodComune As String = vbNullString)
    Dim rs As ADODB.Recordset
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
    
    Set qry = gADOX.Procedures("InserisciUtenzeDomesticheComunali").Command
    qry.Parameters.Refresh
    qry.Parameters(0).value = mCodComune
    qry.Parameters(1).value = mAnnoRiferimento
    
    qry.Execute
    
    Set qry = Nothing
    Set rs = Nothing
    
    DataGrid1.Columns(DataGrid1.Columns.Count - 1).NumberFormat = VALUTA_EURO & " " & FORMATVALUTA
    DataGrid1.Columns(DataGrid1.Columns.Count - 2).NumberFormat = VALUTA_EURO & " " & FORMATVALUTA
    CaricaDati mAnnoRiferimento
    
    Me.Show vbModal
End Sub

Private Sub Cmd_Click()
    Dim rs As ADODB.Recordset
    Dim qry As ADODB.Command
    
    Unload Me
       


End Sub

Sub TotaleGettito()
    Dim Totale  As Double
    Dim i As Integer
    Dim idRow As Integer
    
    On Error Resume Next
    
    Totale = 0
    idRow = DataGrid1.row
    For i = 0 To ADO1.Recordset.RecordCount - 1
        DataGrid1.row = i
        Totale = Totale + (0 & DataGrid1.Columns(DataGrid1.Columns.Count - 1).value) _
        + (0 & DataGrid1.Columns(DataGrid1.Columns.Count - 2).value)
    Next
    
    Txt = Format$(Totale, FORMATVALUTA)
    If Sgn(Txt) = 1 Then
        Txt.ForeColor = vbBlue
    Else
        Txt.ForeColor = vbRed
    End If

    DataGrid1.row = idRow

    

End Sub



Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    If (Round(DataGrid1.Columns(ColIndex).value, 2)) <> DataGrid1.Columns(ColIndex).value Then
       DataGrid1.Columns(ColIndex).value = Round(DataGrid1.Columns(ColIndex).value, 2)
    End If
    
    DataGrid1.Columns(DataGrid1.Columns.Count - 2).value = DataGrid1.Columns(ColIndex).value * 0.3
    DataGrid1.Columns(DataGrid1.Columns.Count - 1).value = (DataGrid1.Columns("Superficie").value - DataGrid1.Columns(ColIndex).value) * 0.4
     
    TotaleGettito
 
   'DataGrid1_KeyDown vbKeyTab, 0
End Sub

Private Sub DataGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    'I campi devono essere numerici
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

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        Sendkeys "{ENTER}", True
    End If

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
    ADO1.ConnectionString = gDBC
    Me.Caption = "Costi Indivisibili - Utenze Domestiche del Comune: " & sComune

    DataGrid1.Left = 0
    DataGrid1.Top = 0
    DataGrid1.Height = Cmd.Top - Cmd.Height
    DataGrid1.Width = Me.ScaleWidth
    
    DTPAnno.value = Format$(Now, "yyyy-mm-dd")

End Sub


Private Function CaricaDati(ByVal Anno As String) As Boolean
    Dim rs As ADODB.Recordset

    Txt = 0
    Set rs = Nothing
    Set rs = DatiTecniciComune(mCodComune, mAnnoRiferimento)
    If Not rs.EOF Then

        Txt = Format$(rs("TotSuperficieUtenzaDomestica"), FORMATVALUTA)
        rs.Close
    End If
    
    CaricaDati = False
    
    On Error Resume Next
    ADO1.Recordset.Close
    On Error GoTo 0
    
    ADO1.CommandType = adCmdText
    
    ADO1.RecordSource = "SELECT UtenzeDomesticheComunali.*, NumeroComponentiFamiglia.Descrizione AS Utenza " & _
                         "FROM UtenzeDomesticheComunali INNER JOIN NumeroComponentiFamiglia ON " & _
                         "UtenzeDomesticheComunali.IdNumeriComponenti = NumeroComponentiFamiglia.NumeroComponenti " & _
                         "WHERE UtenzeDomesticheComunali.CodComune='" & mCodComune & "' and UtenzeDomesticheComunali.AnnoRiferimento='" & Anno & "'"
    
    ADO1.Refresh

    CaricaDati = Not ADO1.Recordset.EOF
    
    If Not CaricaDati Then
        DataGrid1.RowHeight = 30
    Else
        DataGrid1.RowHeight = 270
        TotaleGettito
    End If
    
    
End Function



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




Private Sub Form_Unload(Cancel As Integer)
    MDIMain.mnuComprimi.Enabled = True
End Sub


