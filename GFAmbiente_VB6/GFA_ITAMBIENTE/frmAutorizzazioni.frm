VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAutorizzazioni 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Elenco Autorizzazioni"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbGrid 
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   6720
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   21102593
      CurrentDate     =   39431
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Chiudi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   9240
      TabIndex        =   1
      Top             =   6720
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAutorizzazioni.frx":0000
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   11456
      _Version        =   393216
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      TabAction       =   1
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "cmdAutorizzazioni"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "idAutorizzazione"
         Caption         =   "idAutorizzazione"
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
         DataField       =   "idSoggetto"
         Caption         =   "idSoggetto"
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
         DataField       =   "codSoggetto"
         Caption         =   "codSoggetto"
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
         DataField       =   "Numero"
         Caption         =   "Numero"
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
         DataField       =   "DataInizioValidita"
         Caption         =   "Data Inizio Validita"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "DataScadenza"
         Caption         =   "Data Scadenza"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "RilasciataDa"
         Caption         =   "Rilasciata Da"
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
         DataField       =   "flgDefault"
         Caption         =   "Default"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "Si"
            FalseValue      =   "No"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   7
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   1214,929
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1860,095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   4140,284
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   750,047
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAutorizzazioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private midSoggetto As Long
Private mcodSoggetto As String

Private inCancellazione As Boolean
Private inInserimento As Boolean
Private mColEdit As Long


Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0
            On Error Resume Next
            If DataGrid1.Row > -1 Then DataGrid1.Row = DataGrid1.Row + 1
            Unload Me
    End Select
End Sub



Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
    If Not inCancellazione Then
         DataGrid1.Columns("codSoggetto").Value = mcodSoggetto
         DataGrid1.Columns("idSoggetto").Value = midSoggetto
    End If
End Sub

Private Sub DataGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
     If Not inCancellazione And Not inInserimento And mColEdit = 1 Then Cancel = True: Exit Sub
End Sub

Private Sub DataGrid1_BeforeDelete(Cancel As Integer)
    inCancellazione = True
End Sub

Private Sub DataGrid1_ColEdit(ByVal ColIndex As Integer)
    mColEdit = ColIndex
End Sub

Private Sub DataGrid1_OnAddNew()
    inInserimento = True


End Sub

Private Sub DTPicker1_Change()
    If DTPicker1.Tag <> "" And DTPicker1.enabled = True Then DataGrid1.Columns(DTPicker1.Tag).Value = DTPicker1.Value
End Sub

Private Sub DTPicker1_Click()
    DTPicker1_Change
End Sub
Public Sub showMe(idSoggetto As Long, codSoggetto As String, Des As String)
    
    If (DevEnvMDB.rscmdAutorizzazioni.state = adStateOpen) Then DevEnvMDB.rscmdAutorizzazioni.Close
    DevEnvMDB.cmdAutorizzazioni codSoggetto, idSoggetto
    Me.Caption = "Elenco Autorizzazioni per: " & Des

    midSoggetto = idSoggetto
    mcodSoggetto = codSoggetto
    
    
    Set Me.DataGrid1.DataSource = Nothing
    Set Me.DataGrid1.DataSource = DevEnvMDB
    
    inInserimento = False
    inCancellazione = False
    
    If codSoggetto = "S" Then
        DataGrid1.AllowAddNew = frmLogin.Livello <> LVL_USER
        DataGrid1.AllowDelete = frmLogin.Livello <> LVL_USER
        DataGrid1.AllowUpdate = frmLogin.Livello <> LVL_USER
    End If
    
    Me.Show vbModal
End Sub

Private Sub cmbGrid_Change()
    If cmbGrid.Tag <> "" And cmbGrid.Locked = False Then DataGrid1.Columns(cmbGrid.Tag).Value = IIf(cmbGrid.Text = "Si", -1, IIf(cmbGrid.Text = "No", 0, cmbGrid.Text))
End Sub

Private Sub cmbGrid_Click()
    cmbGrid_Change
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If DataGrid1.col = -1 Or DataGrid1.Row = -1 Then Exit Sub

    
    DTPicker1.Visible = False
    DTPicker1.Move DataGrid1.left + DataGrid1.Columns(DataGrid1.col).left + 7, _
                DataGrid1.top + DataGrid1.RowTop(DataGrid1.Row), _
                DataGrid1.Columns(DataGrid1.col).width, DataGrid1.RowHeight
    
    
    cmbGrid.Visible = False
    cmbGrid.Move DataGrid1.left + DataGrid1.Columns(DataGrid1.col).left + 7, _
                DataGrid1.top + DataGrid1.RowTop(DataGrid1.Row) - 15, _
                DataGrid1.Columns(DataGrid1.col).width
    
    
    Select Case UCase$(DataGrid1.Columns(DataGrid1.col).DataField)
        Case UCase$("DATAINIZIOVALIDITA"), UCase$("DATASCADENZA")
            If (DataGrid1.Columns(DataGrid1.col).Text) = "" Then
                DTPicker1.Value = Now
                DataGrid1.Columns(DataGrid1.col).Value = Now
            Else
                DTPicker1.Value = DataGrid1.Columns(DataGrid1.col).Value
            End If
            
            DTPicker1.Tag = DataGrid1.col
            DTPicker1.Visible = True
        Case UCase$("FlgDefault")
            caricaComboSINO cmbGrid, DataGrid1.col
    
    End Select
    
    

End Sub


Private Sub caricaComboSINO(cmb As ComboBox, colID As Integer)
    cmb.Clear
    
    cmb.AddItem "Si"
    cmb.AddItem "No"
    On Error Resume Next
    cmb.Text = IIf(DataGrid1.Columns(DataGrid1.col).Value = -1, "Si", "No")
    
    cmb.Tag = colID
    
    cmb.Visible = True
End Sub


Private Sub Form_Load()
    cmbGrid.font = DataGrid1.font
    cmbGrid.font.size = DataGrid1.font.size

    

End Sub
