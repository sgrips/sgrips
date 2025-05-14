VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmUtenzeDomestiche 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Coefficiente"
   ClientHeight    =   5670
   ClientLeft      =   2115
   ClientTop       =   1740
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Index           =   1
      Left            =   2220
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   5175
      Width           =   1935
   End
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
      ForeColor       =   &H00808000&
      Height          =   360
      Index           =   0
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   5175
      Width           =   1965
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "&Tariffa Anno Precedente"
      Height          =   495
      Index           =   1
      Left            =   7485
      TabIndex        =   2
      Top             =   5100
      Width           =   1410
   End
   Begin MSAdodcLib.Adodc ADO1 
      Height          =   330
      Left            =   2505
      Top             =   100
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
   Begin VB.CommandButton Cmd 
      Caption         =   "&Chiudi"
      Height          =   495
      Index           =   0
      Left            =   9060
      TabIndex        =   0
      Top             =   5085
      Width           =   1425
   End
   Begin MSComCtl2.DTPicker DTPAnno 
      Height          =   405
      Left            =   6120
      TabIndex        =   7
      Top             =   5145
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
      Format          =   16515075
      UpDown          =   -1  'True
      CurrentDate     =   36967
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmUtenzeDomestiche.frx":0000
      Height          =   4860
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   8573
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
         DataField       =   "NumeroUtenze"
         Caption         =   "Num.Utenze"
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
      BeginProperty Column05 
         DataField       =   "TariffaNComp"
         Caption         =   "Tariffa n.Comp"
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
         DataField       =   "TariffaSuperficie"
         Caption         =   "Tariffa mq"
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
            ColumnWidth     =   2399,811
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1769,953
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Diff. Sup."
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
      Left            =   2220
      TabIndex        =   6
      Top             =   4920
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Rif. Sup."
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
      Left            =   105
      TabIndex        =   4
      Top             =   4905
      Width           =   750
   End
End
Attribute VB_Name = "FrmUtenzeDomestiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCodComune As String
Private mAnnoRiferimento As String

Private sComune As String
Private TotSuperficie As Double

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
    
    DataGrid1.Columns("Tariffa n.Comp").NumberFormat = VALUTA_EURO & " " & FORMATVALUTA
    DataGrid1.Columns("Tariffa mq").NumberFormat = VALUTA_EURO & " " & FORMATVALUTA
    CaricaDati mAnnoRiferimento
    
    Me.Show vbModal
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
                
                
                Set rs = gDBC.Execute("Select * from UtenzeDomesticheComunali where CodComune='" & mCodComune & "' and AnnoRiferimento='" & (Val(mAnnoRiferimento) - 1) & "'")
                If rs.EOF Then
                    gDBC.BeginTrans
                    'inserimento tariffa dell'anno precedente
                    Set qry = gADOX.Procedures("InserisciUtenzeDomesticheAnnoPrecedente").Command
                    Set qry.ActiveConnection = gDBC
                    qry.Parameters.Refresh
                    qry.Parameters("[pCodComune]") = mCodComune
                    qry.Parameters("[pAnnoRiferimento]") = mAnnoRiferimento
                    qry.Parameters("[pAnno]") = mAnnoRiferimento - 1
                    qry.Execute
                    gDBC.CommitTrans
                
                    
                End If
                
                Me.Caption = "Utenze Domestiche del Comune: " & sComune
                CaricaDati Val(mAnnoRiferimento) - 1
                DataGrid1.Columns("Tariffa n.Comp").Locked = False
                DataGrid1.Columns("Tariffa mq").Locked = False
                Cmd(1).Caption = "&Tariffa Anno Corrente"
                
                DTPAnno.value = (mAnnoRiferimento - 1) & "/" & Month(DTPAnno.value) & "/" & Day(DTPAnno.value)
             Else
                Me.Caption = "Utenze Domestiche del Comune: " & sComune
                CaricaDati Val(mAnnoRiferimento)
                DataGrid1.Columns("Tariffa n.Comp").Locked = True
                DataGrid1.Columns("Tariffa mq").Locked = True
                Cmd(1).Caption = "&Tariffa Anno Precedente"
            End If
    End Select


End Sub

Sub TotaleSuperficie()
    Dim TotaleS As Double
    Dim i As Integer
    Dim idRow As Integer
   
    TotaleS = 0
    idRow = DataGrid1.row
    For i = 0 To ADO1.Recordset.RecordCount - 1
        DataGrid1.row = i
        TotaleS = TotaleS + (0 & DataGrid1.Columns("Superficie"))
    Next
    
    Txt(1) = Format$(TotSuperficie - TotaleS, FORMATVALUTA)
    If Sgn(Txt(1)) = 1 Then
        Txt(1).ForeColor = vbBlue
    Else
        Txt(1).ForeColor = vbRed
    End If

    DataGrid1.row = idRow

    

End Sub



Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
 If (Round(DataGrid1.Columns(ColIndex).value, 2)) <> DataGrid1.Columns(ColIndex).value Then
    DataGrid1.Columns(ColIndex).value = Round(DataGrid1.Columns(ColIndex).value, 2)
End If
 TotaleSuperficie
 
 'DataGrid1_KeyDown vbKeyTab, 0
End Sub

Private Sub DataGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    'I campi devono essere numerici
    If Not IsNumeric(DataGrid1.Tag) Then
        MsgBox "Il valore non è numerico", vbExclamation + vbOKOnly, App.Title
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
    Me.Caption = "Utenze Domestiche del Comune: " & sComune

    DataGrid1.Left = 0
    DataGrid1.Top = 0
    DataGrid1.Height = Cmd(0).Top - Cmd(0).Height
    DataGrid1.Width = Me.ScaleWidth
    
    DTPAnno.value = Format$(Now, "yyyy-mm-dd")

End Sub


Private Function CaricaDati(ByVal Anno As String) As Boolean
    Dim rs As ADODB.Recordset

    Txt(0) = 0
    Set rs = Nothing
    Set rs = DatiTecniciComune(mCodComune, mAnnoRiferimento)
    If Not rs.EOF Then
        TotSuperficie = rs("TotSuperficieUtenzaDomestica")
        Txt(0) = Format$(rs("TotSuperficieUtenzaDomestica"), FORMATVALUTA)
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
        Cmd(1).Enabled = False
    Else
        Cmd(1).Enabled = True
        DataGrid1.RowHeight = 270
        TotaleSuperficie
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
