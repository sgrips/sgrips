VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGraficiImpostazioni 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impostazioni"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvwAnni 
      Height          =   2175
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   9596
      EndProperty
   End
   Begin VB.TextBox txt 
      Height          =   975
      Index           =   1
      Left            =   1080
      TabIndex        =   1
      Top             =   645
      Width           =   6435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Annulla"
      Height          =   495
      Index           =   1
      Left            =   5220
      TabIndex        =   5
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Conferma"
      Default         =   -1  'True
      Height          =   495
      Index           =   0
      Left            =   6420
      TabIndex        =   6
      Top             =   5880
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2670
      Width           =   2955
   End
   Begin VB.Frame Frame3 
      Caption         =   "Serie Unite"
      Height          =   555
      Left            =   3420
      TabIndex        =   14
      Top             =   5820
      Width           =   1575
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   1335
         TabIndex        =   17
         Top             =   200
         Width           =   1335
         Begin VB.OptionButton optSerieUnite 
            Caption         =   "Si"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   19
            Top             =   60
            Width           =   615
         End
         Begin VB.OptionButton optSerieUnite 
            Caption         =   "No"
            Height          =   315
            Index           =   1
            Left            =   720
            TabIndex        =   18
            Top             =   0
            Width           =   555
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Indicatori"
      Height          =   555
      Left            =   120
      TabIndex        =   13
      Top             =   5820
      Width           =   1635
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   1455
         TabIndex        =   23
         Top             =   160
         Width           =   1455
         Begin VB.OptionButton optIndicatori 
            Caption         =   "No"
            Height          =   315
            Index           =   1
            Left            =   720
            TabIndex        =   25
            Top             =   0
            Width           =   675
         End
         Begin VB.OptionButton optIndicatori 
            Caption         =   "Si"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   60
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Legenda"
      Height          =   555
      Left            =   1800
      TabIndex        =   12
      Top             =   5820
      Width           =   1575
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   240
         ScaleHeight     =   300
         ScaleWidth      =   1215
         TabIndex        =   20
         Top             =   160
         Width           =   1215
         Begin VB.OptionButton optLegenda 
            Caption         =   "Si"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   22
            Top             =   60
            Width           =   495
         End
         Begin VB.OptionButton optLegenda 
            Caption         =   "No"
            Height          =   315
            Index           =   1
            Left            =   720
            TabIndex        =   21
            Top             =   0
            Width           =   555
         End
      End
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   3
      Left            =   1080
      TabIndex        =   3
      Top             =   2175
      Width           =   6435
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   1710
      Width           =   6435
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   180
      Width           =   6435
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Anni di Confronto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   180
      TabIndex        =   15
      Top             =   3180
      Width           =   1845
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   11
      Top             =   2700
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Asse Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   10
      Top             =   2220
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Asse X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   9
      Top             =   1740
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Note"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   8
      Top             =   660
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Titolo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   570
   End
End
Attribute VB_Name = "frmGraficiImpostazioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pCancel As Boolean

Private pTitolo As String
Private pNote As String
Private pTitoloAsseX As String
Private pTitoloAsseY As String
Private pTipoGrafico As Long
Private pIndicatori As Boolean
Private pLegenda As Boolean
Private pSerieUnite As Boolean
Private pAnnoTariffa As Integer
Private pAnniConfronto As String

Public Function showImpostazioni(titolo As String, note As String, titoloAsseX As String, titoloAsseY As String, tipoGrafico As MSChart20Lib.VtChChartType, indicatori As Boolean, legenda As Boolean, serieUnite As Boolean, AnniConfronto As String, AnnoDiRiferimento As Integer, sCodIstat As String)
    Dim t As MSChart20Lib.VtChChartType
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim Item As ListItem
    Dim anni() As String
    
    txt(0).Text = titolo
    txt(1).Text = note
    txt(2).Text = titoloAsseX
    txt(3).Text = titoloAsseY
    
    If AnniConfronto <> vbNullString Then
        anni = Split(AnniConfronto, "|")
    End If
    'Lettura anni tariffa
    Set rs = gDBC.Execute("SELECT DISTINCT year(DataMovimento) from movimenti")
        
    lvwAnni.SortOrder = 0
    lvwAnni.ListItems.Clear
    lvwAnni.Sorted = True
    
    Do While Not rs.EOF
        
        Set Item = lvwAnni.ListItems.Add(, "|" & rs(0), rs(0))
        
    '    If Not IsNull(rs(1)) Then
   '         Item.ListSubItems.Add , , rs(1)
  '      Else
            Item.ListSubItems.Add , , vbNullString
'        End If
        If AnnoDiRiferimento = rs(0) Then
            Item.ForeColor = vbBlue
        End If
        If AnniConfronto <> vbNullString Then
            For i = LBound(anni) To UBound(anni)
                If rs(0) = anni(i) Then
                    Item.Checked = True
                End If
            Next
        End If
        rs.MoveNext
    Loop
    
    Combo1.Clear
    Combo1.AddItem "grafico a barre"
    Combo1.ItemData(0) = VtChChartType2dBar
    Combo1.AddItem "grafico a linee"
    Combo1.ItemData(1) = VtChChartType2dLine
    Combo1.AddItem "grafico a barre 3D"
    Combo1.ItemData(2) = VtChChartType3dBar
    For i = 0 To Combo1.ListCount - 1
        If Combo1.ItemData(i) = tipoGrafico Then
            Combo1.ListIndex = i
            Exit For
        End If
    Next
    
    optLegenda(IIf(legenda, 0, 1)) = True
    optIndicatori(IIf(indicatori, 0, 1)) = True
    optSerieUnite(IIf(serieUnite, 0, 1)) = True
    
    
    Me.Show vbModal
    
    If Not pCancel Then
        titolo = pTitolo
        note = pNote
        titoloAsseX = pTitoloAsseX
        titoloAsseY = pTitoloAsseY
        tipoGrafico = pTipoGrafico
        indicatori = pIndicatori
        legenda = pLegenda
        serieUnite = pSerieUnite
        AnniConfronto = pAnniConfronto

    End If
    
    showImpostazioni = pCancel
End Function



Private Sub Combo1_Click()
    pTipoGrafico = Combo1.ItemData(Combo1.ListIndex)
End Sub

'Private Sub Combo2_Click()
'    pAnnoTariffa = Val(Combo2.List(Combo2.ListIndex))
'End Sub
'
'Private Sub Combo3_Click()
'    pAnnoConfronto = Val(Combo3.List(Combo3.ListIndex))
'End Sub

Private Sub Command1_Click(Index As Integer)
    Dim anni As String
    Dim Item As ListItem
    
    Select Case Index
    
    Case 0
    
        For Each Item In lvwAnni.ListItems
            If Item.Checked Then
                anni = anni & Item.Text & "|"
            End If
        Next
        If anni = vbNullString Then
            MsgBox "Attenzione occorre selezionare almeno due anni", vbExclamation + vbOKOnly, App.Title
            Exit Sub
        Else
            anni = left$(anni, Len(anni) - 1)
            pAnniConfronto = anni
        End If
        
        pCancel = Index = 1
    
    End Select
    
    Unload Me
End Sub



Private Sub optIndicatori_Click(Index As Integer)
    pIndicatori = IIf(optIndicatori(0).Value = True, True, False)
End Sub

Private Sub optLegenda_Click(Index As Integer)
    pLegenda = IIf(optLegenda(0).Value = True, True, False)
End Sub

Private Sub optSerieUnite_Click(Index As Integer)
    pSerieUnite = IIf(optSerieUnite(0).Value = True, True, False)
End Sub

Private Sub txt_Change(Index As Integer)
    Select Case Index
        Case 0
            pTitolo = txt(Index)
        Case 1
            pNote = txt(Index)
        Case 2
            pTitoloAsseX = txt(Index)
        Case 3
            pTitoloAsseY = txt(Index)
    End Select
End Sub
