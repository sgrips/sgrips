VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmCategorieAttivitaPersonali 
   Caption         =   "Categorie Personalizzate"
   ClientHeight    =   11205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11205
   ScaleWidth      =   12240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd 
      Caption         =   "&Chiudi"
      Height          =   480
      Index           =   0
      Left            =   10800
      TabIndex        =   2
      Top             =   10560
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      Caption         =   "Categorie Attività"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmCategorieAttivitaPersonali.frx":0000
         Height          =   9825
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   11730
         _ExtentX        =   20690
         _ExtentY        =   17330
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   19
         TabAction       =   2
         WrapCellPointer =   -1  'True
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
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
         ColumnCount     =   4
         BeginProperty Column00 
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
         BeginProperty Column01 
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
         BeginProperty Column02 
            DataField       =   "System"
            Caption         =   "System"
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
            DataField       =   "Istat"
            Caption         =   "Istat"
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
            ScrollBars      =   2
            AllowRowSizing  =   -1  'True
            AllowSizing     =   -1  'True
            BeginProperty Column00 
               ColumnAllowSizing=   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   9915,024
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc ADO1 
      Height          =   360
      Left            =   120
      Top             =   10560
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   635
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
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
   Begin MSAdodcLib.Adodc adcCoeff 
      Height          =   360
      Left            =   2280
      Top             =   10560
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   635
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
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
End
Attribute VB_Name = "FrmCategorieAttivitaPersonali"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CaricaDati As Boolean
Dim primavolta As Boolean

Enum StateOpe
    Insert
    Update
    Delete
End Enum
Dim State As StateOpe

Private Sub Cmd_Click(index As Integer)
    Unload Me
End Sub


Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    
    If State = Insert Then
        DoEvents
        On Error GoTo err_manager
        ADO1.Recordset.Update
        ADO1.Recordset.MoveLast
    End If
    Exit Sub
err_manager:
    MsgBox Err.description
End Sub

Private Sub DataGrid1_AfterDelete()
    State = StateOpe.Delete
End Sub

Private Sub DataGrid1_AfterInsert()
    State = StateOpe.Insert
End Sub

Private Sub DataGrid1_AfterUpdate()
    If State = Delete Then Exit Sub
   
    If State = Insert Then
           If Not caricaCoeff Then
            'creo i parametri
            gDBC.BeginTrans
            gDBC.Execute ("insert into CoefficientiCategorieAttivita (CodTipologiaAttivita,CodParGeografica,CodRange,ParteFissa,mink,maxk) values(" & DataGrid1.Columns(0).value & ",'AGNORD','DCAMAG',1,0,99)")
            gDBC.Execute ("insert into CoefficientiCategorieAttivita (CodTipologiaAttivita,CodParGeografica,CodRange,ParteFissa,mink,maxk) values(" & DataGrid1.Columns(0).value & ",'AGNORD','DCAMAG',0,0,99)")
            gDBC.Execute ("insert into CoefficientiCategorieAttivita (CodTipologiaAttivita,CodParGeografica,CodRange,ParteFissa,mink,maxk) values(" & DataGrid1.Columns(0).value & ",'AGNORD','DCAMIN',1,0,99)")
            gDBC.Execute ("insert into CoefficientiCategorieAttivita (CodTipologiaAttivita,CodParGeografica,CodRange,ParteFissa,mink,maxk) values(" & DataGrid1.Columns(0).value & ",'AGNORD','DCAMIN',0,0,99)")
            gDBC.Execute ("insert into CoefficientiCategorieAttivita (CodTipologiaAttivita,CodParGeografica,CodRange,ParteFissa,mink,maxk) values(" & DataGrid1.Columns(0).value & ",'AGCENTRO','DCAMAG',1,0,99)")
            gDBC.Execute ("insert into CoefficientiCategorieAttivita (CodTipologiaAttivita,CodParGeografica,CodRange,ParteFissa,mink,maxk) values(" & DataGrid1.Columns(0).value & ",'AGCENTRO','DCAMAG',0,0,99)")
            gDBC.Execute ("insert into CoefficientiCategorieAttivita (CodTipologiaAttivita,CodParGeografica,CodRange,ParteFissa,mink,maxk) values(" & DataGrid1.Columns(0).value & ",'AGCENTRO','DCAMIN',1,0,99)")
            gDBC.Execute ("insert into CoefficientiCategorieAttivita (CodTipologiaAttivita,CodParGeografica,CodRange,ParteFissa,mink,maxk) values(" & DataGrid1.Columns(0).value & ",'AGCENTRO','DCAMIN',0,0,99)")
            gDBC.Execute ("insert into CoefficientiCategorieAttivita (CodTipologiaAttivita,CodParGeografica,CodRange,ParteFissa,mink,maxk) values(" & DataGrid1.Columns(0).value & ",'AGSUD','DCAMAG',1,0,99)")
            gDBC.Execute ("insert into CoefficientiCategorieAttivita (CodTipologiaAttivita,CodParGeografica,CodRange,ParteFissa,mink,maxk) values(" & DataGrid1.Columns(0).value & ",'AGSUD','DCAMAG',0,0,99)")
            gDBC.Execute ("insert into CoefficientiCategorieAttivita (CodTipologiaAttivita,CodParGeografica,CodRange,ParteFissa,mink,maxk) values(" & DataGrid1.Columns(0).value & ",'AGSUD','DCAMIN',1,0,99)")
            gDBC.Execute ("insert into CoefficientiCategorieAttivita (CodTipologiaAttivita,CodParGeografica,CodRange,ParteFissa,mink,maxk) values(" & DataGrid1.Columns(0).value & ",'AGSUD','DCAMIN',0,0,99)")
            gDBC.CommitTrans
        End If
        caricaCoeff
        State = Update
    End If
    
    
        
End Sub

Private Function caricaCoeff() As Boolean
    adcCoeff.RecordSource = "select * from CoefficientiCategorieAttivita where CodTipologiaAttivita=-1"
    adcCoeff.Refresh
    
    
    If DataGrid1.row = -1 Then Exit Function
    If DataGrid1.Columns(0).Text = "" Then Exit Function

    
    adcCoeff.RecordSource = "select * from CoefficientiCategorieAttivita where CodTipologiaAttivita=" & DataGrid1.Columns(0).value
    adcCoeff.Refresh
   
    
    caricaCoeff = Not adcCoeff.Recordset.EOF
End Function

Private Sub DataGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    If Not primavolta Then
        If ColIndex = 0 Then
        
   
            
            
        
            If Not IsNumeric(DataGrid1.Tag) Then
                MsgBox "Il valore non è numerico", vbExclamation + vbOKOnly, App.Title
                Cancel = True
                Exit Sub
            ElseIf DataGrid1.Tag < 30 Then
                MsgBox "Il valore deve essere maggiore di 30", vbExclamation + vbOKOnly, App.Title
                Cancel = True
                Exit Sub
            End If
        End If
        
        If ColIndex = 1 Then
            If DataGrid1.Tag = "" Then
                MsgBox "La descrizione della Categoria è obbligatoria", vbExclamation + vbOKOnly, App.Title
                Cancel = True
            End If
        End If
    Else
        Cancel = True
    End If
    
    primavolta = False
        
End Sub


Public Sub Carica()
   CaricaDati = False
    primavolta = True
    On Error Resume Next
    ADO1.Recordset.Close
    On Error GoTo 0
    
    
    ADO1.CommandType = adCmdText
    ADO1.RecordSource = "SELECT * from CategorieAttivita where [System]=0 order by CodAttivita"
    
    ADO1.Refresh

    CaricaDati = Not ADO1.Recordset.EOF
    
  
    caricaCoeff
    
    primavolta = False
    
    Me.Show vbModal
End Sub

Private Sub DataGrid1_BeforeDelete(Cancel As Integer)
    gDBC.Execute "DELETE ValoriCoefficientiCAComunali.* FROM CoefficientiCategorieAttivita " & _
                              "INNER JOIN ValoriCoefficientiCAComunali ON " & _
                              "CoefficientiCategorieAttivita.IdCoefficiente = ValoriCoefficientiCAComunali.IdCoefficiente " & _
                              "WHERE CoefficientiCategorieAttivita.CodTipologiaAttivita=" + DataGrid1.Columns(0)
                              
    gDBC.Execute "delete * from ValoriCoefficientiIQSCAComunali where CodAttivita=" & DataGrid1.Columns(0)

    'Dotazione
    gDBC.Execute "delete * from ContenitoriInDotazioneCAComunali where CodAttivita=" & DataGrid1.Columns(0)
    'Svuotamenti
    gDBC.Execute "delete * from ContenitoriSvuotamentiCAComunali where CodAttivita=" & DataGrid1.Columns(0)
            
    'Riduzioni
    gDBC.Execute "delete * from BonusUtenzeAttivita where CodAttivita=" & DataGrid1.Columns(0)
                              
    gDBC.Execute "delete from CategorieAttivitaComunali where IdCategoriaAttivita=" & DataGrid1.Columns(0).value
    gDBC.Execute "delete from CoefficientiCategorieAttivita where CodTipologiaAttivita=" & DataGrid1.Columns(0).value
    
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

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        Sendkeys "{ENTER}", True
        DoEvents
    End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  caricaCoeff
End Sub

Private Sub Form_Activate()
    'Mi posiziono sulla griglia
    DataGrid1.SetFocus
End Sub

Private Sub Form_Load()
    
    MDIMain.mnuComprimi.Enabled = False
    Me.Caption = "Categorie Attivita Personalizzate"

    ADO1.ConnectionString = gDBC
    adcCoeff.ConnectionString = gDBC
    

End Sub




Private Sub Form_Resize()
    On Error Resume Next
    Cmd(0).Top = Me.ScaleHeight - Cmd(0).Height - 15
     
    Frame1.Height = Cmd(0).Top - 180
    DataGrid1.Height = Frame1.Height - 500
End Sub

Private Sub Form_Unload(Cancel As Integer)
   MDIMain.mnuComprimi.Enabled = True
End Sub






