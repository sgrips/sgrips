VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAmmortamenti 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ammortamenti"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13950
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   13950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmAmmortamenti.frx":0000
      Height          =   5325
      Left            =   90
      TabIndex        =   5
      Top             =   0
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   9393
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   10
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
         DataField       =   "CodTipoAmmortamento"
         Caption         =   "CodTipoAmmortamento"
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
         Caption         =   "Descrizione"
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
         DataField       =   "Quantita"
         Caption         =   "Quantita"
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
         DataField       =   "ValoreCommerciale"
         Caption         =   "Valore Commerciale"
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
         DataField       =   "AnniAmmortamento"
         Caption         =   "Anni"
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
         DataField       =   "PercValoreResiduo"
         Caption         =   "Valore Residuo (%) "
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
         DataField       =   "PercInteresseComputo"
         Caption         =   "Interesse Computo (%)"
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
         DataField       =   "QuotaAnnua"
         Caption         =   "Quota Annua"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
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
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   1665,071
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   5430,047
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1500,095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   645,165
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1379,906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1620,284
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1604,976
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADO1 
      Height          =   375
      Left            =   11700
      Top             =   6555
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   661
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
   Begin VB.TextBox txtTotale 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11520
      TabIndex        =   3
      Text            =   "0"
      Top             =   5490
      Width           =   2265
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aggiungi"
      Height          =   555
      Index           =   0
      Left            =   10665
      TabIndex        =   2
      Top             =   8220
      Width           =   1470
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Totalizza"
      Height          =   540
      Index           =   1
      Left            =   12315
      TabIndex        =   1
      Top             =   8235
      Width           =   1500
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3300
      Left            =   60
      TabIndex        =   0
      Top             =   5460
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   5821
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Totale:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   10680
      TabIndex        =   4
      Top             =   5500
      Width           =   795
   End
End
Attribute VB_Name = "frmAmmortamenti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mComune As String
Private mAnnoRiferimento As String
Private mTotale As Double
Private OldTxtValue As String

Private Const cmd_Totalizza = 1
Private Const cmd_Aggiungi = 0

Function Carica(Totale As Double, Comune As String, AnnoRiferimento As String) As Double

    mComune = Comune
    mAnnoRiferimento = AnnoRiferimento
    
    mTotale = Totale
    
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
                    AggiungiAmmoratamento lItem
                End If
            Next
            
            CaricaDatiAmmortamento
            End With
    End Select
End Sub

Private Function EsisteGiaAmmortamento(Codice As Long) As Boolean
    Dim rs As Recordset
    
    Set rs = gDBC.Execute("Select * from Ammortamenti WHERE Ammortamenti.CodComune='" & mComune & "' and Ammortamenti.AnnoRiferimento='" & mAnnoRiferimento & "' and CodTipoAmmortamento=" & Codice)
    EsisteGiaAmmortamento = Not rs.EOF
End Function

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)

    DataGrid1.Columns(9) = CalcoloQuota
    txtTotale = Format$(Totalizza, FORMATVALUTA)
     
End Sub


Function CalcoloQuota(Optional ColIndex As Integer = -1, Optional ValCol As Double) As Double
    Dim Valore As Double
    Dim Quota As Double
    Dim valcol5 As Double, valcol4 As Double, valcol6 As Double, valcol7 As Double, valcol8 As Double
    
    On Error GoTo CalcoloQuotaERR
    valcol5 = IIf(ColIndex = 5, ValCol, DataGrid1.Columns(5))
    valcol4 = IIf(ColIndex = 4, ValCol, DataGrid1.Columns(4))
    valcol6 = IIf(ColIndex = 6, ValCol, DataGrid1.Columns(6))
    valcol7 = IIf(ColIndex = 7, ValCol, DataGrid1.Columns(7))
    valcol8 = IIf(ColIndex = 8, ValCol, DataGrid1.Columns(8))
    
    Valore = valcol5 * valcol4
    
    If valcol6 = 0 Then Exit Function
    
    
    CalcoloQuota = ((Valore - (Valore * (valcol7 / 100))) / valcol6) * (1 + (valcol8 / 100))
    Exit Function
CalcoloQuotaERR:
    'se errore
    CalcoloQuota = -1000
End Function

Function Totalizza() As Double
    Dim i As Integer
    
    If Not ADO1.Recordset.EOF Then
        mTotale = 0
        For i = 0 To ADO1.Recordset.RecordCount - 1
            DataGrid1.Row = i
            mTotale = mTotale + CDbl(DataGrid1.Columns(9).Value)
        Next
        ADO1.Recordset.MoveFirst
    Else
        mTotale = CDbl(txtTotale)
    End If
    
    Totalizza = mTotale
End Function



Private Sub DataGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    'I campi devono essere numerici
    If Not IsNumeric(DataGrid1.Tag) Then
        MsgBox "Il valore non è numerico", vbExclamation + vbOKOnly, App.Title
        Cancel = True
        Exit Sub
    End If
    
    'I campi devono essere delle percentuali
    If (ColIndex = 7 Or ColIndex = 8) And (CDbl(DataGrid1.Tag) < 0 Or CDbl(DataGrid1.Tag) > 100) Then
        MsgBox "Percentuale non valida", vbExclamation + vbOKOnly, App.Title
        Cancel = True
    End If
    
    'Attenzione la quota non può essere negativa
    If CalcoloQuota(ColIndex, DataGrid1.Tag) < 0 Then
        MsgBox "Il valore inserito porta la Quota ad un valore non ammesso", vbExclamation + vbOKOnly, App.Title
        Cancel = True
    End If
    
    'anni ammortamento non maggiore di 50
    If DataGrid1.Tag > 50 And ColIndex = 6 Then
        MsgBox "Anni Ammortamento massimo 50", vbExclamation + vbOKOnly, App.Title
        Cancel = True
    End If
    
End Sub

Private Sub DataGrid1_BeforeDelete(Cancel As Integer)
    Dim g As New ADODB.Connection
    
    
    g.ConnectionString = gDBC.ConnectionString
    g.Open
    If MsgBox("Cancellare Ammortamento:(" & DataGrid1.Columns(3) & ")", vbQuestion + vbYesNo, App.Title) = vbYes Then
            g.BeginTrans
            g.Execute ("Delete from Ammortamenti  WHERE Ammortamenti.CodComune='" & mComune & "' and Ammortamenti.AnnoRiferimento='" & mAnnoRiferimento & "' and Ammortamenti.CodTipoAmmortamento=" & DataGrid1.Columns(2))
            g.CommitTrans
            g.Close
            Set g = Nothing
    End If
    
    Cancel = 1
End Sub

Private Sub DataGrid1_Error(ByVal DataError As Integer, Response As Integer)
    If DataError = 7011 Then
        Response = 0
        CaricaDatiAmmortamento
        txtTotale = Format$(Totalizza, FORMATVALUTA)

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

Private Sub Form_Load()
    
    txtTotale = mTotale
    
    txtTotale = Format$(mTotale, FORMATVALUTA)
    
    ADO1.ConnectionString = gDBC.ConnectionString
    
    ADO1.CommandType = adCmdText
    ADO1.RecordSource = "SELECT Ammortamenti.CodComune, Ammortamenti.AnnoRiferimento, Ammortamenti.CodTipoAmmortamento,TipiAmmortamento.Descrizione, Ammortamenti.Quantita, Ammortamenti.ValoreCommerciale, Ammortamenti.AnniAmmortamento, Ammortamenti.PercValoreResiduo, Ammortamenti.PercInteresseComputo, Ammortamenti.QuotaAnnua" _
    & " FROM Ammortamenti INNER JOIN TipiAmmortamento ON Ammortamenti.CodTipoAmmortamento = TipiAmmortamento.CodTipoAmmortamento" _
    & " WHERE Ammortamenti.CodComune='" & mComune & "' and Ammortamenti.AnnoRiferimento='" & mAnnoRiferimento & "'"
    
    lvw.SmallIcons = MDIMain.ImgList
    
    CaricaInListView lvw, ICON_AMMORTAMENTO, Lvw_TipiAmmortamenti
    
    DataGrid1.Columns("Valore Commerciale").NumberFormat = FORMATVALUTA
    DataGrid1.Columns("Quota Annua").NumberFormat = VALUTA_EURO & " " & FORMATVALUTA
    If Not CaricaDatiAmmortamento Then
    End If


End Sub


Private Function CaricaDatiAmmortamento() As Boolean
    
    CaricaDatiAmmortamento = False
    
    On Error Resume Next
    ADO1.Recordset.Close
    On Error GoTo 0
    ADO1.Refresh
    Me.DataGrid1.ReBind
    
    CaricaDatiAmmortamento = Not ADO1.Recordset.EOF
    
    If Not CaricaDatiAmmortamento Then
        DataGrid1.RowHeight = 30
    Else
        DataGrid1.RowHeight = 270
    End If
End Function

Private Sub AggiungiAmmoratamento(Item As ListItem)
    If Item.Selected Then
        If EsisteGiaAmmortamento(Item.Text) Then
            MsgBox "Il tipo automezzo:" & Item.SubItems(1) & vbCrLf & "già presente tra gli ammortamenti", vbExclamation + vbOKOnly
        Else
            gDBC.BeginTrans
            gDBC.Execute "Insert into Ammortamenti (CodComune,AnnoRiferimento,CodTipoAmmortamento) Values(" & Chr$(34) & mComune & Chr$(34) & ",'" & mAnnoRiferimento & "'," & CDbl(Item.Text) & ")"
            gDBC.CommitTrans
        End If
    End If
End Sub


Private Sub Lvw_DblClick()
    If lvw.SelectedItem Is Nothing Then Exit Sub
    AggiungiAmmoratamento lvw.SelectedItem
    CaricaDatiAmmortamento
End Sub

Private Sub txtTotale_GotFocus()
    OldTxtValue = txtTotale
End Sub

Private Sub txtTotale_Validate(Cancel As Boolean)
        'Il valore deve essere numerico
        If Not IsNumeric(txtTotale) Then MsgBox "Attenzione il valore deve essere numerico", vbOKOnly + vbExclamation, App.Title: Cancel = True: txtTotale = OldTxtValue: Exit Sub

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



