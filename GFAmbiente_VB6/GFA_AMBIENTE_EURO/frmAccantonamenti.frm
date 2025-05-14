VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAccantonamenti 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accantonamenti"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAccantonamenti.frx":0000
      Height          =   4950
      Left            =   45
      TabIndex        =   3
      Top             =   0
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   8731
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
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
         DataField       =   "CodAccantonamento"
         Caption         =   "CodAccantonamento"
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
         DataField       =   "CostoUnitario"
         Caption         =   "Costo Unitario"
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
      BeginProperty Column06 
         DataField       =   "Totale"
         Caption         =   "Totale"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2910,047
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   2129,953
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ado1 
      Height          =   330
      Left            =   5235
      Top             =   5145
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   582
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
      Caption         =   "ado1"
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
      Left            =   930
      TabIndex        =   1
      Text            =   "0"
      Top             =   5085
      Width           =   2265
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "&Totalizza"
      Height          =   390
      Left            =   7290
      TabIndex        =   0
      Top             =   5070
      Width           =   1095
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
      Left            =   90
      TabIndex        =   2
      Top             =   5115
      Width           =   795
   End
End
Attribute VB_Name = "frmAccantonamenti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mComune As String
Private mAnnoRiferimento As String
Private mTotale As Double
Private OldTxtValue As String

Function Carica(Totale As Double, Comune As String, AnnoRiferimento As String) As Double


    mComune = Comune
    mAnnoRiferimento = AnnoRiferimento
    mTotale = Totale
    
    DoEvents
    
    Me.Show vbModal

    Carica = mTotale

End Function


Private Sub Cmd_Click()
    
    Totalizza
    Unload Me
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    DataGrid1.Columns(6) = DataGrid1.Columns(5) * CDbl(DataGrid1.Columns(4))

    txtTotale = Format$(Totalizza, FORMATVALUTA)
End Sub

Function Totalizza() As Double
    Dim i As Integer
    
    mTotale = 0
    For i = 0 To DataGrid1.VisibleRows - 1
        DataGrid1.Row = i
        mTotale = mTotale + CDbl(DataGrid1.Columns(6).Value)
    Next
    
    If mTotale > 0 Then
        Totalizza = mTotale
    Else
        mTotale = CDbl(txtTotale.Text)
        Totalizza = CDbl(txtTotale.Text)
    End If
    
End Function



Private Sub DataGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    'I campi devono essere numerici
    If Not IsNumeric(DataGrid1.Tag) Then
        MsgBox "Il valore non è numerico", vbExclamation + vbOKOnly, App.Title
        Cancel = True
    End If
    
End Sub

Private Sub DataGrid1_BeforeDelete(Cancel As Integer)
    
    If MsgBox("Cancellare Accantonamento:(" & DataGrid1.Columns(3) & ")", vbQuestion + vbYesNo, App.Title) = vbYes Then
            gDBC.BeginTrans
            gDBC.Execute ("Delete from Accantonamenti  WHERE Accantonamenti.CodComune='" & mComune & "' and Accantonamenti.AnnoRiferimento='" & mAnnoRiferimento & "' and Accantonamenti.CodTipoAmmortamento=" & DataGrid1.Columns(2))
            gDBC.CommitTrans
            CaricaDatiAccantonamento
            
            txtTotale = Format$(Totalizza, FORMATVALUTA)
    End If
    
    Cancel = 1

End Sub

Private Sub DataGrid1_Error(ByVal DataError As Integer, Response As Integer)
    If DataError = 7011 Then Response = 0
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
    
    
    txtTotale = CDbl(mTotale)
    
    ADO1.ConnectionString = gDBC
    ADO1.CommandType = adCmdText
    ADO1.RecordSource = "SELECT Accantonamenti.CodComune, Accantonamenti.AnnoRiferimento, Accantonamenti.CodAccantonamento, TipiAccantonamento.Descrizione, Accantonamenti.CostoUnitario, Accantonamenti.Quantita, Accantonamenti.Totale" _
    & " FROM Accantonamenti INNER JOIN TipiAccantonamento ON Accantonamenti.CodAccantonamento = TipiAccantonamento.CodAccantonamento" _
    & " WHERE Accantonamenti.CodComune='" & mComune & "' and Accantonamenti.AnnoRiferimento='" & mAnnoRiferimento & "'"
    
    DataGrid1.Left = 0
    DataGrid1.Top = 0
    DataGrid1.Height = cmd.Top - 15 * 4
    DataGrid1.Width = Me.ScaleWidth
    
    DataGrid1.Columns(4).NumberFormat = FORMAT6DEC
    DataGrid1.Columns(6).NumberFormat = VALUTA_EURO & " " & FORMATVALUTA
    
    txtTotale = Format$(mTotale, FORMATVALUTA)
    
    CaricaDatiAccantonamento
End Sub


Private Function CaricaDatiAccantonamento() As Boolean
    
    CaricaDatiAccantonamento = False
    
    On Error Resume Next
    ADO1.Recordset.Close
    On Error GoTo 0
    ADO1.Refresh
    
    If ADO1.Recordset.EOF Then
        'Carico gli accantonamenti per il comune
        gDBC.BeginTrans
        gDBC.Execute ("Insert into Accantonamenti (CodComune,AnnoRiferimento,CodAccantonamento) Select " & Chr$(34) & mComune & Chr$(34) & " as ESPR1,'" & mAnnoRiferimento & "' as ESPR2,CodAccantonamento From TipiAccantonamento")
        gDBC.CommitTrans
        
        ADO1.Refresh
    End If
        
    CaricaDatiAccantonamento = Not ADO1.Recordset.EOF
End Function


Private Sub txtTotale_GotFocus()
 OldTxtValue = txtTotale
End Sub

Private Sub txtTotale_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
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


