VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmCoefficienti507UD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Coefficiente"
   ClientHeight    =   4935
   ClientLeft      =   4320
   ClientTop       =   840
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1020
      Left            =   0
      TabIndex        =   1
      Top             =   3900
      Width           =   8055
      Begin VB.CommandButton Cmd 
         Caption         =   "&Chiudi"
         Default         =   -1  'True
         Height          =   390
         Index           =   0
         Left            =   6570
         TabIndex        =   7
         Top             =   450
         Width           =   1005
      End
      Begin VB.CommandButton Cmd 
         Caption         =   " &IPS UD"
         Height          =   375
         Index           =   1
         Left            =   5580
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
         Left            =   1950
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
         Left            =   75
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPAnno 
         Height          =   405
         Left            =   3840
         TabIndex        =   8
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
         Format          =   56164355
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
         Left            =   1920
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
         Left            =   60
         TabIndex        =   4
         Top             =   240
         Width           =   660
      End
   End
   Begin MSAdodcLib.Adodc ADO1 
      Height          =   345
      Left            =   3180
      Top             =   2760
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
      Bindings        =   "FrmCoefficiente507UD.frx":0000
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   6588
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
      ColumnCount     =   5
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
         DataField       =   "CodUtenza"
         Caption         =   "CodUtenza"
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
         Caption         =   "Numero Componenti"
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
         DataField       =   "Valore"
         Caption         =   "Valore"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
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
            ColumnWidth     =   1035,213
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmCoefficienti507UD"
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

Private MaxCoeff As Double
Private MinCoeff As Double

Public Sub Carica(Optional sCodComune As String = vbNullString)

    
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
    

    
    
    bCaricaDati = CaricaDati
    
    
    
    Me.Show vbModal
End Sub



Private Sub Cmd_Click(Index As Integer)
    Select Case Index
        Case 0
            LetParamValue mAnnoRiferimento, mCodComune, "COEFFIQSUD", True, CDbl(txt(COEFF_IQSUD_ID).Text)
            'COEFFIPSUD
            LetParamValue mAnnoRiferimento, mCodComune, "COEFFIPSUD", True, CDbl(txt(COEFF_IPSUD_ID).Text)
            Unload Me
        Case 1
            CalcolaIPS
    End Select
End Sub

Private Sub DataGrid1_BeforeUpdate(Cancel As Integer)
    
    'Valida il valore inserito
    'If (CDbl(DataGrid1.Columns("valore")) > MaxCoeff Or CDbl(DataGrid1.Columns("valore")) < MinCoeff) And DataGrid1.Columns("IQS") <> 0 Then
    '    If MsgBox("Valore non ammesso nel range, Premere YES per forzatura valore.", vbQuestion + vbYesNo, App.Title) = vbYes Then
    '        Cancel = False
    '    Else
    '        Cancel = True
    '    End If
    'End If

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
    Me.Caption = "Coefficiente - IQS - IPS"

    ADO1.ConnectionString = gDBC
    ADO1.CommandType = adCmdText
    
    DataGrid1.Left = 0
    DataGrid1.Top = 0
    DataGrid1.Height = Me.Frame1.Top

   GetParamValue mAnnoRiferimento, mCodComune, "COEFFIQSUD", MaxCoeff, MinCoeff
   
   DTPAnno.Value = Now
End Sub


Private Function CaricaDati() As Boolean
    Dim rs As Recordset
    Dim qry As New ADODB.Command
    Dim sAbitanti As String
    
    CaricaDati = False
    
    'Reperisco i dati comunali e anagrafici
    Set rs = DatiAnagraficiComunali(mCodComune, mAnnoRiferimento, sAbitanti)
    If rs.EOF Then
        MsgBox "Attenzione non sono stati trovati i dati anagrafici", vbExclamation + vbOKOnly, App.Title
    Else
        'Inserisco le attività
        Set rs = Nothing
        Set qry = gADOX.Procedures("InsUtenzeDomesticheInCoeff507UDComunali").Command
        qry.Parameters.Refresh
        qry.Parameters("pCodComune") = mCodComune
        qry.Parameters("pAnnoRiferimento") = mAnnoRiferimento
        qry.Execute
        Set qry = Nothing
    End If
    
    On Error Resume Next
    ADO1.Recordset.Close
    On Error GoTo 0
    
    ADO1.RecordSource = "SELECT ValoriCoefficienti507UDComunali.CodComune,NumeroComponentiFamiglia.Descrizione, ValoriCoefficienti507UDComunali.AnnoRiferimento, ValoriCoefficienti507UDComunali.CodUtenza, ValoriCoefficienti507UDComunali.Valore " & _
                        "FROM NumeroComponentiFamiglia INNER JOIN (UtenzeDomesticheComunali INNER JOIN ValoriCoefficienti507UDComunali ON (UtenzeDomesticheComunali.IdNumeriComponenti = ValoriCoefficienti507UDComunali.CodUtenza) AND (UtenzeDomesticheComunali.AnnoRiferimento = ValoriCoefficienti507UDComunali.AnnoRiferimento) AND (UtenzeDomesticheComunali.CodComune = ValoriCoefficienti507UDComunali.CodComune)) ON NumeroComponentiFamiglia.NumeroComponenti = UtenzeDomesticheComunali.IdNumeriComponenti " & _
                        "WHERE (((ValoriCoefficienti507UDComunali.CodComune)='" & mCodComune & "') AND ((ValoriCoefficienti507UDComunali.AnnoRiferimento)='" & mAnnoRiferimento & "')) order by ValoriCoefficienti507UDComunali.CodUtenza"
    ADO1.Refresh

    CaricaDati = Not ADO1.Recordset.EOF
    
    If Not CaricaDati Then
        DataGrid1.RowHeight = 30
    Else
        DataGrid1.RowHeight = 270
        DataGrid1.Refresh
    End If
    
   txt(COEFF_IQSUD_ID) = Format$(GetParamValue(mAnnoRiferimento, mCodComune, "COEFFIQSUD"), "0.00")
   txt(COEFF_IPSUD_ID) = Format$(GetParamValue(mAnnoRiferimento, mCodComune, "COEFFIPSUD"), "0.00")
    
End Function



Sub CalcolaIPS()
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
    
    If rsDatiTecniciComune.EOF Then
        MsgBox "Attenzione dati tecnici mancanti", vbExclamation + vbOKOnly, App.Title
        Exit Sub
    ElseIf rsANA.EOF Then
        MsgBox "Attenzione dati anagrafici", vbExclamation + vbOKOnly, App.Title
        Exit Sub
    End If

    
    'KM
    KM = rsDatiTecniciComune("QRifiutiTotali") / rsDatiTecniciComune("TotSuperficieIscrRuolo")
    
    
    IPSUD = (rsDatiTecniciComune("QRifiutiUtenzaDomestica") / rsDatiTecniciComune("TotSuperficieUtenzaDomestica")) / KM
    txt(COEFF_IPSUD_ID).Text = Format$(IPSUD, "0.00")
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
