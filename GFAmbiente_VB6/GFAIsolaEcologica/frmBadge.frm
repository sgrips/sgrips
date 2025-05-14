VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBadge 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Elenco Badge"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbGrid 
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   4440
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
      Height          =   495
      Index           =   0
      Left            =   7440
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmBadge.frx":0000
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   6800
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      TabAction       =   1
      FormatLocked    =   -1  'True
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
      DataMember      =   "cmdBadge"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "CodUtente"
         Caption         =   "CodUtente"
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
         DataField       =   "CodBadge"
         Caption         =   "CodBadge"
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
         DataField       =   "ProgressivoBadge"
         Caption         =   "ProgressivoBadge"
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
         DataField       =   "DataBadgeInizio"
         Caption         =   "DataBadgeInizio"
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
         DataField       =   "DataBadgeFine"
         Caption         =   "DataBadgeFine"
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
         DataField       =   "default"
         Caption         =   "default"
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
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1635,024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1635,024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   705,26
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
Attribute VB_Name = "frmBadge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcodUtente As String

Public Sub showMe(codUtente As Long, Des As String)

    If (DevEnvMDB.rscmdBadge.state = adStateOpen) Then DevEnvMDB.rscmdBadge.Close
    DevEnvMDB.cmdBadge codUtente
    mcodUtente = codUtente
    Label1(0).Caption = Des
    
    Me.Show vbModal
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
    End Select
End Sub



Private Sub Form_Load()
    DataGrid1.AllowAddNew = False
    DataGrid1.AllowDelete = False
    DataGrid1.AllowUpdate = False
    
    cmbGrid.font = DataGrid1.font
    cmbGrid.font.size = DataGrid1.font.size
    
End Sub

Private Sub cmbGrid_Change()
    If cmbGrid.Tag <> "" And cmbGrid.Locked = False Then DataGrid1.Columns(cmbGrid.Tag).Value = IIf(cmbGrid.Text = "Si", -1, IIf(cmbGrid.Text = "No", 0, cmbGrid.Text))
End Sub

Private Sub cmbGrid_Click()
    cmbGrid_Change
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If DataGrid1.col = -1 Or DataGrid1.Row = -1 Then Exit Sub

    
    cmbGrid.Visible = False
    cmbGrid.Move DataGrid1.Left + DataGrid1.Columns(DataGrid1.col).Left + 7, _
                DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) - 15, _
                DataGrid1.Columns(DataGrid1.col).width
    
    
    Select Case UCase$(DataGrid1.Columns(DataGrid1.col).DataField)
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



