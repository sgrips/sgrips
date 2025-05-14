VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmCoefficienteKCKD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Coefficiente"
   ClientHeight    =   9930
   ClientLeft      =   2115
   ClientTop       =   1740
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmCoefficienteKCKD.frx":0000
      Height          =   9300
      Left            =   15
      TabIndex        =   1
      Top             =   -15
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   16404
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
         DataField       =   "IDCoefficiente"
         Caption         =   "IDCoefficiente"
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
         DataField       =   "AnnoRiferimento"
         Caption         =   "Anno"
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
         DataField       =   "Descrizione"
         Caption         =   "Attivita"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "MinK"
         Caption         =   "Minimo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "MaxK"
         Caption         =   "Massimo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Valore"
         Caption         =   "Valore"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
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
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   6809,953
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1124,787
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADO1 
      Height          =   345
      Left            =   2205
      Top             =   7830
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.CommandButton Cmd 
      Caption         =   "&Chiudi"
      Height          =   390
      Left            =   9360
      TabIndex        =   0
      Top             =   9465
      Width           =   1440
   End
   Begin MSComCtl2.DTPicker DTPAnno 
      Height          =   405
      Left            =   7575
      TabIndex        =   2
      Top             =   9450
      Width           =   1650
      _ExtentX        =   2910
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
      Format          =   19726339
      UpDown          =   -1  'True
      CurrentDate     =   36967
   End
End
Attribute VB_Name = "FrmCoefficienteKCKD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private mCodComune As String
Private mParteFissa As Boolean
Private mArea  As String
Private mAbitanti As Double
Private mAnnoRiferimento As String
Private sParteF As String
Private sParGeo As String
Private sAbitanti As String
Private bCaricaDati As Boolean

Public Sub Carica(bParteFissa As Boolean, Optional sCodComune As String = vbNullString)

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
    mParteFissa = bParteFissa
    DTPAnno.Value = Now
    bCaricaDati = CaricaDati
    
    Me.Show vbModal
End Sub

Private Sub Cmd_Click()
    Unload Me
    
End Sub

Private Sub DataGrid1_BeforeUpdate(Cancel As Integer)
    'Valida il valore inserito
    DoEvents
    If CDbl(DataGrid1.Columns("Valore")) > CDbl(DataGrid1.Columns("Massimo")) Or CDbl(DataGrid1.Columns("Valore")) < CDbl(DataGrid1.Columns("Minimo")) Then
        If MsgBox("Valore non ammesso nel range, Premere YES per forzatura valore.", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Cancel = False
        Else
            Cancel = True
        End If
    End If

End Sub

Private Sub DataGrid1_Error(ByVal DataError As Integer, Response As Integer)
    If DataError = 7011 Then
        Response = 0
    End If
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        SendKeys "{ENTER}", True
    End If

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
End Sub

Private Sub DTPAnno_Change()
    mAnnoRiferimento = Year(DTPAnno.Value)
    CaricaDati
End Sub

Private Sub Form_Activate()
    'Mi posiziono sulla griglia
    DataGrid1.SetFocus
End Sub

Private Sub Form_Load()
    MDIMain.mnuComprimi.Enabled = False
    
    ADO1.ConnectionString = gDBC
    ADO1.CommandType = adCmdText
    
    DataGrid1.Left = 0
    DataGrid1.Top = 0
    DataGrid1.Height = cmd.Top - 15 * 4
    DataGrid1.Width = Me.ScaleWidth

End Sub


Private Function CaricaDati() As Boolean
    Dim rs As Recordset
    Dim qry As New ADODB.Command
    
    CaricaDati = False
    
    On Error Resume Next
    ADO1.Recordset.Close
    On Error GoTo 0
    
    'Reperisco i dati comunali e anagrafici
    Set rs = DatiAnagraficiComunali(mCodComune, mAnnoRiferimento, sAbitanti)
    If rs.EOF Then
        MsgBox "Attenzione non sono stati trovati i dati anagrafici", vbExclamation + vbOKOnly, App.Title
        mArea = "No Dati"
        mAbitanti = 0
        sParGeo = "No Dati"
    Else
        'Area di appartenenza del comune , NumeroAbitanti
        mArea = rs("Descrizione")
        mAbitanti = rs("NumAbitantiResidenti")
        sParGeo = rs("AreaGeografica")
        'controllo se esistono già i coefficienti
        Set rs = Nothing
        
        Set qry = gADOX.Procedures("InsCategorieAttivitaInCoeffCategorieAttivitaComunali").Command
        qry.Parameters.Refresh
        qry.Parameters("Comune") = mCodComune
        qry.Parameters("[CodParGeo]") = sParGeo
        qry.Parameters("[CodR]") = sAbitanti
        qry.Parameters("ParteF") = mParteFissa
        qry.Parameters("[Anno]") = mAnnoRiferimento
        qry.Execute
        
        'Inserisco le attività
        Set qry = gADOX.Procedures("InserisciCategorieAttivitaComunali").Command
        qry.Parameters.Refresh
        qry.Parameters(0) = mCodComune
        qry.Parameters(1) = mAnnoRiferimento
        
        qry.Execute
    
    End If
    
    Me.Caption = "Coefficiente - " & IIf(mParteFissa, "Parte Fissa KC", "Parte Variabile KD") & " - " & mArea


    ADO1.RecordSource = "SELECT ValoriCoefficientiCAComunali.AnnoRiferimento,ValoriCoefficientiCAComunali.IdCoefficiente,ValoriCoefficientiCAComunali.CodComune,CoefficientiCategorieAttivita.MinK, CoefficientiCategorieAttivita.MaxK, ValoriCoefficientiCAComunali.Valore, CategorieAttivitaComunali.Descrizione " & _
                         "FROM (CoefficientiCategorieAttivita INNER JOIN ValoriCoefficientiCAComunali ON CoefficientiCategorieAttivita.IdCoefficiente = ValoriCoefficientiCAComunali.IdCoefficiente) " _
                         & "INNER JOIN CategorieAttivitaComunali ON CoefficientiCategorieAttivita.CodTipologiaAttivita = CategorieAttivitaComunali.IdCategoriaAttivita and ValoriCoefficientiCAComunali.CodComune = CategorieAttivitaComunali.CodComune and ValoriCoefficientiCAComunali.AnnoRiferimento = CategorieAttivitaComunali.AnnoRiferimento " _
                         & "WHERE ValoriCoefficientiCAComunali.CodComune='" & mCodComune & "' and CoefficientiCategorieAttivita.CodParGeografica='" & sParGeo & "' AND CoefficientiCategorieAttivita.CodRange='" & sAbitanti & "' AND CoefficientiCategorieAttivita.ParteFissa=" & IIf(mParteFissa, "True", "False") & " and " _
                         & "ValoriCoefficientiCAComunali.AnnoRiferimento='" & mAnnoRiferimento & "' order by CoefficientiCategorieAttivita.CodTipologiaAttivita"
    
    
    
    ADO1.Refresh

    CaricaDati = Not ADO1.Recordset.EOF
    
    If Not CaricaDati Then
        DataGrid1.RowHeight = 30
    Else
        DataGrid1.RowHeight = 270
        DataGrid1.Refresh
    End If
    
End Function


Private Sub Form_Unload(Cancel As Integer)
    'Calcola Quantità Rifiuti su UA
    If Not mParteFissa Then QuantitaRifiutiDaUtenzeAttivita mCodComune, mAnnoRiferimento
    
    MDIMain.mnuComprimi.Enabled = True
End Sub
