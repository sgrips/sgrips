VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAutomezzi2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Elenco Automezzi"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbGrid 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
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
      Left            =   6960
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAutomezzi.frx":0000
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      DataMember      =   "cmdAutomezzi"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "codUtente"
         Caption         =   "codUtente"
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
      BeginProperty Column02 
         DataField       =   "targa"
         Caption         =   "Targa"
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
      BeginProperty Column04 
         DataField       =   "descrizione"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   1170,142
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1844,787
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   4094,929
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "frmAutomezzi2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcodUtente As String
Private mcodSoggetto As String

Private inCancellazione As Boolean
Private inInserimento As Boolean
Private mColEdit As Long

Public Sub showMe(codUtente As Long, des As String, codSoggetto As String)

    If (DevEnvMDB.rscmdAutomezzi.State = adStateOpen) Then DevEnvMDB.rscmdAutomezzi.Close
    DevEnvMDB.cmdAutomezzi codUtente, codSoggetto
    mcodUtente = codUtente
    mcodSoggetto = codSoggetto
    Label1(0).Caption = des
    
    Set Me.DataGrid1.DataSource = Nothing
    Set Me.DataGrid1.DataSource = DevEnvMDB
    
    inInserimento = False
    inCancellazione = False
    
    Me.Show vbModal
End Sub

Private Sub cmbGrid_Change()
    If cmbGrid.Tag <> "" And cmbGrid.Locked = False Then DataGrid1.Columns(cmbGrid.Tag).Value = IIf(cmbGrid.Text = "Si", -1, IIf(cmbGrid.Text = "No", 0, cmbGrid.Text))
End Sub

Private Sub cmbGrid_Click()
    cmbGrid_Change
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0
            If DataGrid1.Row > -1 Then DataGrid1.Row = DataGrid1.Row + 1
            Unload Me
    End Select
End Sub



Private Sub DataGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
     If Not inCancellazione And Not inInserimento And mColEdit = 1 Then Cancel = True: Exit Sub
End Sub

Private Sub DataGrid1_BeforeDelete(Cancel As Integer)
    inCancellazione = True
End Sub

Private Sub DataGrid1_BeforeUpdate(Cancel As Integer)
    
    If Not inCancellazione Then
         DataGrid1.Columns(0).Value = mcodUtente
         DataGrid1.Columns(1).Value = mcodSoggetto
    End If

    
    inCancellazione = False
    inInserimento = False
End Sub

Private Sub DataGrid1_ColEdit(ByVal ColIndex As Integer)
    mColEdit = ColIndex
End Sub

Private Sub DataGrid1_OnAddNew()
'DataGrid1.Columns(0).Value = mcodUtente
inInserimento = True
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If DataGrid1.Col = -1 Or DataGrid1.Row = -1 Then Exit Sub
    
    cmbGrid.Visible = False
    cmbGrid.Move DataGrid1.Left + DataGrid1.Columns(DataGrid1.Col).Left + 7, _
                DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) - 15, _
                DataGrid1.Columns(DataGrid1.Col).Width
    Select Case UCase$(DataGrid1.Columns(DataGrid1.Col).DataField)
        Case UCase$("FlgDefault")
            caricaComboSINO cmbGrid, DataGrid1.Col
            
    End Select
    
    
End Sub

Private Sub caricaComboSINO(cmb As ComboBox, colID As Integer)
    cmb.Clear
    
    cmb.AddItem "Si"
    cmb.AddItem "No"
    On Error Resume Next
    cmb.Text = IIf(DataGrid1.Columns(DataGrid1.Col).Value = -1, "Si", "No")
    
    cmb.Tag = colID
    
    cmb.Visible = True
End Sub

Private Sub Form_Load()
    
    DataGrid1.AllowAddNew = frmLogin.Livello <> LVL_USER
    DataGrid1.AllowDelete = frmLogin.Livello <> LVL_USER
    DataGrid1.AllowUpdate = frmLogin.Livello <> LVL_USER
    
    
    cmbGrid.Font = DataGrid1.Font
    cmbGrid.Font.Size = DataGrid1.Font.Size
    
    cmbGrid.Locked = frmLogin.Livello <> LVL_ADMINISTRATOR
    
End Sub
