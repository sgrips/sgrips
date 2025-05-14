VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form FrmCategorieAttivita 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Categorie Attività Comunali"
   ClientHeight    =   10560
   ClientLeft      =   2115
   ClientTop       =   1740
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10560
   ScaleWidth      =   12750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd 
      Caption         =   "&Reimposta Descrizioni"
      Height          =   495
      Index           =   3
      Left            =   3600
      TabIndex        =   9
      Top             =   9975
      Width           =   1410
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "&Copia Anno Precedente"
      Height          =   495
      Index           =   2
      Left            =   8175
      TabIndex        =   8
      Top             =   9945
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
      Index           =   1
      Left            =   105
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   10095
      Width           =   1560
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
      Index           =   0
      Left            =   1830
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   10095
      Width           =   1560
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "&Chiudi"
      Height          =   480
      Index           =   0
      Left            =   11280
      TabIndex        =   3
      Top             =   9960
      Width           =   1410
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "&Tariffa Anno Precedente"
      Height          =   495
      Index           =   1
      Left            =   9690
      TabIndex        =   2
      Top             =   9945
      Width           =   1410
   End
   Begin MSAdodcLib.Adodc ADO1 
      Height          =   360
      Left            =   4380
      Top             =   3300
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
      Left            =   6900
      TabIndex        =   7
      Top             =   9975
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
      Format          =   56033283
      UpDown          =   -1  'True
      CurrentDate     =   36967
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmCategorieAttivita.frx":0000
      Height          =   9825
      Left            =   0
      TabIndex        =   1
      Top             =   -30
      Width           =   12690
      _ExtentX        =   22384
      _ExtentY        =   17330
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
      ColumnCount     =   8
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
         DataField       =   "NumRuoli"
         Caption         =   "Nr.Ruoli"
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
         DataField       =   "Tariffa"
         Caption         =   "Tariffa"
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
         DataField       =   "TariffaPF"
         Caption         =   "TariffaF"
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
      BeginProperty Column07 
         DataField       =   "TariffaPV"
         Caption         =   "TariffaV"
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
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1170,142
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1170,142
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tot.Ruoli"
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
      Index           =   1
      Left            =   90
      TabIndex        =   6
      Top             =   9840
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tot. Superficie"
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
      Left            =   1770
      TabIndex        =   4
      Top             =   9840
      Width           =   1245
   End
End
Attribute VB_Name = "FrmCategorieAttivita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCodComune As String
Private sComune As String
Private mAnnoRiferimento As String
Private Const TOTSUPERFICIE_ID = 0
Private Const NUMRUOLI_ID = 1


Sub Totali()
    Dim TotaleS As Double
    Dim TotaleN As Double
    Dim i As Integer
    Dim CurrentCol As Integer
    Dim CurrentRow As Integer
    
    CurrentRow = 0
    CurrentCol = 0
    
    CurrentRow = DataGrid1.row
    CurrentCol = DataGrid1.col + 1
    
    TotaleN = 0: TotaleS = 0
    For i = 0 To ADO1.Recordset.RecordCount - 1
        DataGrid1.row = i
        TotaleN = TotaleN + (0 & DataGrid1.Columns("Nr.Ruoli"))
        TotaleS = TotaleS + (0 & DataGrid1.Columns("Superficie"))
    Next

    If CurrentRow > -1 And CurrentCol < DataGrid1.Columns.Count Then
        DataGrid1.row = CurrentRow
        DataGrid1.col = CurrentCol
        
        Txt(TOTSUPERFICIE_ID) = Format$(TotaleS, FORMATVALUTA)
        Txt(NUMRUOLI_ID) = Format$(TotaleN, FORMATVALUTA)
        
    End If
        

End Sub




Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)

    If ColIndex <= 2 Then Exit Sub
    
    If (Round(DataGrid1.Columns(ColIndex).value, 2)) <> DataGrid1.Columns(ColIndex).value Then
       DataGrid1.Columns(ColIndex).value = Round(DataGrid1.Columns(ColIndex).value, 2)
    End If
End Sub

Private Sub DataGrid1_AfterUpdate()
    
    Totali
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
    ElseIf Val(DataGrid1.Tag) = 0 Then 'se imposto a 0 la superfice allora devo cancellare anche i relativi coefficienti UA 507, KD, KC e Riduzioni
       ' If MsgBox("Attenzione impostando la superfice a 0, l'attività sarà ignorata nel calcolo tariffa e i relativi coefficienti e riduzioni sarannò cancellati." + _
        '       vbCrLf + "Si vuole procedere ?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim qry As ADODB.Command
            Set qry = CreateObject("ADODB.Command")
            Set qry.ActiveConnection = gDBC
            
            gDBC.BeginTrans
            'KC e KD
            qry.CommandText = "DELETE ValoriCoefficientiCAComunali.* FROM CoefficientiCategorieAttivita " & _
                              "INNER JOIN ValoriCoefficientiCAComunali ON " & _
                              "CoefficientiCategorieAttivita.IdCoefficiente = ValoriCoefficientiCAComunali.IdCoefficiente " & _
                              "WHERE (((ValoriCoefficientiCAComunali.CodComune)='" + mCodComune + "') AND ((ValoriCoefficientiCAComunali.AnnoRiferimento)='" + mAnnoRiferimento + "') AND " & _
                              "((CoefficientiCategorieAttivita.CodTipologiaAttivita)=" + DataGrid1.Columns(0) + "))"
            qry.Execute
            
            
            'IQS
            qry.CommandText = "delete * from ValoriCoefficientiIQSCAComunali where CodComune='" & mCodComune & _
                "' and AnnoRiferimento='" & mAnnoRiferimento & _
                "' and CodAttivita=" & DataGrid1.Columns(0) & ""
            qry.Execute
            'Dotazione
            qry.CommandText = "delete * from ContenitoriInDotazioneCAComunali where CodComune='" & mCodComune & _
                "' and AnnoRiferimento='" & mAnnoRiferimento & _
                "' and CodAttivita=" & DataGrid1.Columns(0) & ""
            qry.Execute
            'Svuotamenti
            qry.CommandText = "delete * from ContenitoriSvuotamentiCAComunali where CodComune='" & mCodComune & _
                "' and AnnoRiferimento='" & mAnnoRiferimento & _
                "' and CodAttivita=" & DataGrid1.Columns(0) & ""
            qry.Execute
            
            
            'Riduzioni
            qry.CommandText = "delete * from BonusUtenzeAttivita where CodComune='" & mCodComune & _
                    "' and AnnoRiferimento='" & mAnnoRiferimento & _
                    "' and CodAttivita=" & DataGrid1.Columns(0) & ""
            qry.Execute
            gDBC.CommitTrans
            
            
            FrmCoefficiente.CreaCoefficenti tpo_coeff_KAKB, mCodComune, mAnnoRiferimento, True
            FrmCoefficiente.CreaCoefficenti tpo_coeff_KAKB, mCodComune, mAnnoRiferimento, False
            FrmCoefficiente.CreaCoefficenti tpo_coeff_KCKD, mCodComune, mAnnoRiferimento, True
            FrmCoefficiente.CreaCoefficenti tpo_coeff_KCKD, mCodComune, mAnnoRiferimento, False
            
            Set qry.ActiveConnection = Nothing
            Set qry = Nothing
        
        'Else
        '    Cancel = True
        '    Exit Sub
        'End If
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
    Me.Caption = "Categorie Attivita del Comune: " & sComune

    ADO1.ConnectionString = gDBC

    
    DataGrid1.Left = 0
    DataGrid1.Top = 0
    DataGrid1.Height = Cmd(0).Top - 100
    DataGrid1.Width = Me.ScaleWidth
        
    DataGrid1.Columns("Tariffa").NumberFormat = VALUTA_EURO & " " & FORMATVALUTA
    DataGrid1.Columns("TariffaF").NumberFormat = VALUTA_EURO & " " & FORMATVALUTA
    DataGrid1.Columns("TariffaV").NumberFormat = VALUTA_EURO & " " & FORMATVALUTA
    
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
        Cmd(1).Enabled = False
    Else
        Cmd(1).Enabled = True
        DataGrid1.RowHeight = 270
        DataGrid1.Refresh
    End If
    
    Totali
    
End Function


Private Sub Form_Unload(Cancel As Integer)
    'Calcola Quantità Rifiuti su UA
    
    QuantitaRifiutiDaUtenzeAttivita mCodComune, mAnnoRiferimento
    MDIMain.mnuComprimi.Enabled = True
End Sub



Private Sub Cmd_Click(index As Integer)
    Dim rs As ADODB.Recordset
    Dim qry As ADODB.Command
    

    Select Case index
         Case 0
            Unload Me
        Case 1
                         
            If Cmd(1).Caption = "&Tariffa Anno Precedente" Then
                Set rs = gDBC.Execute("Select * from AnagraficaComunale where CodComune='" & mCodComune & "' and AnnoRiferimento='" & (Val(mAnnoRiferimento) - 1) & "'")
                If rs.EOF Then gDBC.Execute "Insert into AnagraficaComunale (CodComune,AnnoRiferimento) values ('" & mCodComune & "','" & Val(mAnnoRiferimento) - 1 & "')"
                Set rs = gDBC.Execute("Select * from CostoComunale where CodComune='" & mCodComune & "' and AnnoRiferimento='" & (Val(mAnnoRiferimento) - 1) & "'")
                If rs.EOF Then gDBC.Execute "Insert into CostoComunale (CodComune,AnnoRiferimento) values ('" & mCodComune & "','" & Val(mAnnoRiferimento) - 1 & "')"
                
                Set rs = gDBC.Execute("Select * from CategorieAttivitaComunali where CodComune='" & mCodComune & "' and AnnoRiferimento='" & (Val(mAnnoRiferimento) - 1) & "'")
                If rs.EOF Then
                    gDBC.BeginTrans
                    'inserimento tariffa dell'anno precedente
                    Set qry = gADOX.Procedures("InserisciCategorieAttivitaAnnoPrecedente").Command
                    Set qry.ActiveConnection = gDBC
                    qry.Parameters.Refresh
                    qry.Parameters("[pCodComune]") = mCodComune
                    qry.Parameters("[pAnnoRiferimento]") = mAnnoRiferimento
                    qry.Parameters("[pAnno]") = mAnnoRiferimento - 1
                    qry.Execute
                    gDBC.CommitTrans
                

                End If
                
                Me.Caption = "Categorie Attivita del Comune: " & sComune
                CaricaDati Val(mAnnoRiferimento) - 1
                DataGrid1.Columns("Tariffa").Locked = False
                Cmd(1).Caption = "&Tariffa Anno Corrente"
                
                DTPAnno.value = (mAnnoRiferimento - 1) & "/" & Month(DTPAnno.value) & "/" & Day(DTPAnno.value)
            Else
                Me.Caption = "Categorie Attivita del Comune: " & sComune
                CaricaDati Val(mAnnoRiferimento)
                DataGrid1.Columns("Tariffa").Locked = True
                Cmd(1).Caption = "&Tariffa Anno Precedente"
            End If
        Case 2
        
            Set qry = gADOX.Procedures("CopiaDatiCategorieAttivita").Command
            Set qry.ActiveConnection = gDBC
            qry.Parameters.Refresh
            qry.Parameters("[pCodComune]") = mCodComune
            qry.Parameters("[pAnnoRiferimentoDaCopiare]") = mAnnoRiferimento - 1
            qry.Parameters("[pAnnoRiferimento]") = mAnnoRiferimento
            qry.Execute

            CaricaDati mAnnoRiferimento
        Case 3
        
            If MsgBox("Si vuole procedere con il ripristino delle descrizioni originali ?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            
                gDBC.BeginTrans
                Set qry = CreateObject("ADODB.Command")
                qry.CommandText = "UPDATE CategorieAttivita INNER JOIN CategorieAttivitaComunali " + _
                                  "ON CategorieAttivita.CodAttivita = CategorieAttivitaComunali.IdCategoriaAttivita SET CategorieAttivitaComunali.Descrizione = [CategorieAttivita].[Descrizione]" + _
                                  "WHERE (((CategorieAttivitaComunali.CodComune)='" + mCodComune + "') AND ((CategorieAttivitaComunali.AnnoRiferimento)='" + mAnnoRiferimento + "'))"
                Set qry.ActiveConnection = gDBC
                qry.Execute
                            
                gDBC.CommitTrans
                Set qry.ActiveConnection = Nothing
                Set qry = Nothing
                
                CaricaDati (mAnnoRiferimento)
            End If
    End Select

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





