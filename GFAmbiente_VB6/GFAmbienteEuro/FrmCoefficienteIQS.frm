VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmCoefficienti507 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Coefficiente"
   ClientHeight    =   9450
   ClientLeft      =   4320
   ClientTop       =   840
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1020
      Left            =   30
      TabIndex        =   1
      Top             =   8325
      Width           =   8895
      Begin VB.CommandButton Cmd 
         Caption         =   "&Chiudi"
         Default         =   -1  'True
         Height          =   390
         Index           =   0
         Left            =   7770
         TabIndex        =   8
         Top             =   450
         Width           =   1005
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "&IPS UA"
         Height          =   375
         Index           =   1
         Left            =   6750
         TabIndex        =   7
         Top             =   465
         Width           =   915
      End
      Begin VB.CommandButton Cmd 
         Caption         =   " &IPS UD"
         Height          =   375
         Index           =   2
         Left            =   5715
         TabIndex        =   6
         Top             =   465
         Width           =   885
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   2010
         TabIndex        =   3
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   135
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPAnno 
         Height          =   405
         Left            =   3990
         TabIndex        =   9
         Top             =   450
         Width           =   1665
         _ExtentX        =   2937
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
         Format          =   22806531
         UpDown          =   -1  'True
         CurrentDate     =   36967
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IQS UD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   1980
         TabIndex        =   5
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IPS UD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   660
      End
   End
   Begin MSAdodcLib.Adodc ADO1 
      Height          =   345
      Left            =   3375
      Top             =   585
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmCoefficienteIQS.frx":0000
      Height          =   8235
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   14526
      _Version        =   393216
      AllowArrows     =   -1  'True
      BorderStyle     =   0
      ForeColor       =   -2147483642
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
         DataField       =   "CodAttivita"
         Caption         =   "CodAttivita"
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
         DataField       =   "IQS"
         Caption         =   "IQS"
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
         DataField       =   "IPS"
         Caption         =   "IPS"
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
            ColumnWidth     =   6510,047
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   689,953
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmCoefficienti507"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCodComune As String
Private mAnnoRiferimento As String
Private bCaricaDati As Boolean

Private Const COEFF_IQSUD_ID = 2
Private Const COEFF_IPSUD_ID = 3

Private MaxCoeff As Currency
Private MinCoeff As Currency

Public Sub Carica(Optional sCodComune As String = vbNullString)
    Dim rs As Recordset
    Dim qry As New ADODB.Command
    Dim sAbitanti As String
    
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
    
    'Reperisco i dati comunali e anagrafici
    Set rs = DatiAnagraficiComunali(mCodComune, mAnnoRiferimento, sAbitanti)
    If rs.EOF Then MsgBox "Attenzione non sono stati trovati i dati anagrafici", vbExclamation + vbOKOnly, App.Title
    
    'Inserisco le attività
    Set rs = Nothing
    Set qry = gADOX.Procedures("InsCategorieAttivitaInCoeffIQSCAComunali").Command
    qry.Parameters.Refresh
    qry.Parameters("[pCodComune]") = mCodComune
    qry.Parameters("[pAnnoRiferimento]") = mAnnoRiferimento
    qry.Execute
    
    
    bCaricaDati = CaricaDati
    
    Me.Show vbModal
End Sub



Private Sub Cmd_Click(Index As Integer)
    Select Case Index
        Case 0
            LetParamValue mCodComune, "COEFFIQSUD", True, CDbl(txt(COEFF_IQSUD_ID).Text)
            'COEFFIPSUD
            LetParamValue mCodComune, "COEFFIPSUD", True, CDbl(txt(COEFF_IPSUD_ID).Text)
            Unload Me
        Case 1
            CalcolaIPS False
        Case 2
            CalcolaIPS True
    End Select
End Sub

Private Sub DataGrid1_BeforeUpdate(Cancel As Integer)
    
    'Valida il valore inserito
    If (CDbl(DataGrid1.Columns("IQS")) > MaxCoeff Or CDbl(DataGrid1.Columns("IQS")) < MinCoeff) And DataGrid1.Columns("IQS") <> 0 Then
        If MsgBox("Valore IQS non ammesso nel range, Premere YES per forzatura valore.", vbQuestion + vbYesNo, App.Title) = vbYes Then
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
    Me.Caption = "Coefficiente - IQS"

    ADO1.ConnectionString = gDBC
    ADO1.CommandType = adCmdText
    
    DataGrid1.Left = 0
    DataGrid1.Top = 0
    DataGrid1.Height = Me.Frame1.Top

   GetParamValue mCodComune, "COEFFIQSUD", MaxCoeff, MinCoeff
   
   txt(COEFF_IQSUD_ID) = Format$(GetParamValue(mCodComune, "COEFFIQSUD"), "0.00")
   txt(COEFF_IPSUD_ID) = Format$(GetParamValue(mCodComune, "COEFFIPSUD"), "0.00")
   
   DTPAnno.Value = Now
End Sub


Private Function CaricaDati() As Boolean
    CaricaDati = False
    
    On Error Resume Next
    ADO1.Recordset.Close
    On Error GoTo 0
    
    ADO1.RecordSource = "SELECT ValoriCoefficientiIQSCAComunali.AnnoRiferimento,ValoriCoefficientiIQSCAComunali.CodAttivita,ValoriCoefficientiIQSCAComunali.CodComune, ValoriCoefficientiIQSCAComunali.IQS,ValoriCoefficientiIQSCAComunali.IPS, CategorieAttivitaComunali.Descrizione " & _
                         "FROM  CategorieAttivitaComunali INNER JOIN ValoriCoefficientiIQSCAComunali ON ValoriCoefficientiIQSCAComunali.CodAttivita = CategorieAttivitaComunali.IdCategoriaAttivita and " & _
                         " ValoriCoefficientiIQSCAComunali.CodComune = CategorieAttivitaComunali.CodComune and " & _
                         " ValoriCoefficientiIQSCAComunali.AnnoRiferimento = CategorieAttivitaComunali.AnnoRiferimento " & _
                         "WHERE ValoriCoefficientiIQSCAComunali.CodComune='" & mCodComune & "'and ValoriCoefficientiIQSCAComunali.AnnoRiferimento='" & mAnnoRiferimento & "' order by ValoriCoefficientiIQSCAComunali.CodAttivita"
    
    ADO1.Refresh

    CaricaDati = Not ADO1.Recordset.EOF
    
    If Not CaricaDati Then
        DataGrid1.RowHeight = 30
    Else
        DataGrid1.RowHeight = 270
        DataGrid1.Refresh
    End If
    
End Function



Sub CalcolaIPS(UD As Boolean)
    Dim rsKD As ADODB.Recordset
    Dim KM As Double
    Dim rsDatiTecniciComune As ADODB.Recordset
    Dim rsANA As ADODB.Recordset
    Dim mParAbitanti As String
    Dim i  As Integer
    Dim iOld As Integer
    Dim IPSUD As Double
    
    Set rsANA = DatiAnagraficiComunali(mCodComune, mAnnoRiferimento, mParAbitanti)
    Set rsDatiTecniciComune = DatiTecniciComune(mCodComune, mAnnoRiferimento)
    
    'KM
    KM = rsDatiTecniciComune("QRifiutiTotali") / rsDatiTecniciComune("TotSuperficieIscrRuolo")
    
    
    If UD Then
        IPSUD = (rsDatiTecniciComune("QRifiutiUtenzaDomestica") / rsDatiTecniciComune("TotSuperficieUtenzaDomestica")) / KM
        txt(COEFF_IPSUD_ID).Text = Format$(IPSUD, "0.00")
    Else
        'KD
        Set rsKD = SelezionaCoeffK(CoeffKD, mCodComune, mAnnoRiferimento, rsANA("AreaGeografica"), mParAbitanti)
        
        i = 0
        iOld = DataGrid1.Row
        DataGrid1.Row = i
        Do While Not rsKD.EOF
            If DataGrid1.Columns("CodAttivita").Text = rsKD("CodTipologiaAttivita") Then
                DataGrid1.Columns("IPS") = rsKD("Valore") / KM
                i = i + 1
                If i < DataGrid1.VisibleRows Then DataGrid1.Row = i
            End If
            rsKD.MoveNext
        Loop
        DataGrid1.Row = iOld
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.mnuComprimi.Enabled = True
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    If Index = COEFF_IQSUD_ID And ((CDbl(txt(COEFF_IQSUD_ID)) > MaxCoeff Or CDbl(txt(COEFF_IQSUD_ID)) < MinCoeff) And CDbl(txt(COEFF_IQSUD_ID)) <> 0) Then
        If MsgBox("Valore IQS non ammesso nel range " & MinCoeff & " - " & MaxCoeff & vbCrLf & " Premere YES per forzatura valore.", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Cancel = False
        Else
            Cancel = True
        End If
    End If
    
End Sub
