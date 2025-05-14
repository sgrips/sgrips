VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSimulazioni 
   Caption         =   "frmSimulazioni"
   ClientHeight    =   10605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14325
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10605
   ScaleWidth      =   14325
   Begin VB.CommandButton Command1 
      Caption         =   "&Stampe"
      Height          =   375
      Index           =   5
      Left            =   7560
      TabIndex        =   8
      Top             =   10080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Report XLS"
      Height          =   375
      Index           =   4
      Left            =   9120
      TabIndex        =   7
      Top             =   10080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Converti InTariffa"
      Height          =   375
      Index           =   3
      Left            =   6000
      TabIndex        =   6
      Top             =   10080
      Width           =   1455
   End
   Begin EasyTax.SgripsSplit SgripsSplit 
      Height          =   9735
      Left            =   4440
      TabIndex        =   5
      Top             =   0
      Width           =   105
      _ExtentX        =   185
      _ExtentY        =   16113
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancella"
      Height          =   375
      Index           =   2
      Left            =   10680
      TabIndex        =   4
      Top             =   10080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Crea"
      Height          =   375
      Index           =   1
      Left            =   12000
      TabIndex        =   3
      Top             =   10080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Chiudi"
      Height          =   375
      Index           =   0
      Left            =   13320
      TabIndex        =   2
      Top             =   10080
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   8655
      Left            =   5160
      ScaleHeight     =   8595
      ScaleWidth      =   8355
      TabIndex        =   1
      Top             =   360
      Width           =   8415
   End
   Begin MSComctlLib.TreeView tvwSimu 
      Height          =   9615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   16960
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   706
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   6
      FullRowSelect   =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSimulazioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CMD_ANNULLA = 0
Private Const CMD_CREA = 1
Private Const CMD_CANCELLA = 2
Private Const CMD_PUBBLICA = 3
Private Const CMD_REPORT = 4
Private Const CMD_PRINT = 5

Private pNodeSelected As Node
Private pLastParent As Long
Private currentObject As Object
Private currentObjectEditChild As Object

Public Sub showME()
    
    CreaAlbero
    
    Me.Show
End Sub

Function getKey(Key As String) As String
    getKey = "|" & Key
End Function

Function existsItem(items As Object, Key As String) As Boolean
    On Error Resume Next
    Dim Item As Node
    Set Item = Nothing
    Set Item = tvwSimu.Nodes(Key)
    existsItem = IIf(Item Is Nothing, False, True)
End Function


Private Sub Command1_Click(index As Integer)
    
    Select Case index
        Case 2
            mnuCancella_Click
        Case 1
            mnuCrea_Click
        Case 0
            Unload Me
        Case 3
            
            PublicaComeTariffa Right$(pNodeSelected.Key, 6), Mid$(pNodeSelected.Key, 2, 6), Mid$(pNodeSelected.Key, 8, 1)
             
            'deleteSimulation Right$(pNodeSelected.Key, 6), Mid$(pNodeSelected.Key, 2, 6), Mid$(pNodeSelected.Key, 8, 1)
            
            'CreaAlbero
            
        Case 4
            ReportTariffa Right$(pNodeSelected.Key, 6), Mid$(pNodeSelected.Key, 2, 6), Mid$(pNodeSelected.Key, 8, 1)
        Case 5
            frmStampe.Carica Mid$(pNodeSelected.Key, 2, 6), vbNullString, Right$(pNodeSelected.Key, 4)

    End Select
End Sub

Private Sub Form_Load()

    SgripsSplit.attachControl tvwSimu, Picture1
    
    Me.Move 0, 0, MDIMain.ScaleWidth - MDIMain.pic.ScaleWidth, MDIMain.ScaleHeight

End Sub

Private Sub Form_Resize()
    Dim i As Integer
    Dim nextX As Integer
    
    nextX = Me.ScaleWidth - Command1(0).Width - 100
    For i = Command1.LBound To Command1.UBound
        Command1(i).Move nextX, Me.ScaleHeight - Command1(i).Height - 100
        If i < Command1.UBound Then nextX = nextX - Command1(i + 1).Width - 100
    Next
    tvwSimu.Height = Command1(0).Top - 100
    SgripsSplit.Height = Me.ScaleHeight
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set SgripsSplit.LeftObj = Nothing
    Set SgripsSplit.RightObj = Nothing
    
    If Not currentObject Is Nothing Then
        currentObject.Visible = False
        SetParent currentObject.hwnd, pLastParent
        Unload currentObject
        Set currentObject = Nothing
    End If
    
    If Not currentObjectEditChild Is Nothing Then
            currentObjectEditChild.Visible = False
        SetParent currentObjectEditChild.hwnd, pLastParent
        Unload currentObjectEditChild
        Set currentObjectEditChild = Nothing
    End If
    
End Sub

Public Sub mnuCancella_Click()
    'cancellazione
    
    deleteSimulation Right$(pNodeSelected.Key, 6), Mid$(pNodeSelected.Key, 2, 6), Mid$(pNodeSelected.Key, 8, 1)
    On Error Resume Next
    tvwSimu.Nodes.Remove (pNodeSelected.Key)
End Sub

Public Sub mnuCrea_Click()
    'creazione
    createSimulation Right$(pNodeSelected.Key, 4), Mid$(pNodeSelected.Key, 2, 6), Mid$(pNodeSelected.Key, 8, 1)
    CreaAlbero
End Sub




Private Sub Picture1_Resize()
    If Not currentObject Is Nothing Then currentObject.Move 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
    If Not currentObjectEditChild Is Nothing Then currentObjectEditChild.Move 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
End Sub

Private Sub tvwSimu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim mnode As Node
    Dim s() As String
    Dim i As Integer
    
    If Button = vbRightButton Then
        Set mnode = tvwSimu.HitTest(X, Y)
        On Error Resume Next
        i = -1
        
        s = Split(mnode.FullPath, "\")
        i = UBound(s)
        If i = 2 Or i = 3 Then
            If i = 2 Then
                MDIMain.mnuCreaSimulazione.Visible = True
                MDIMain.mnuCancellaSimulazione.Visible = False
                
            Else
                MDIMain.mnuCancellaSimulazione.Visible = True
                MDIMain.mnuCreaSimulazione.Visible = False
            End If
            
            tvwSimu.Tag = mnode.Key
            Set pNodeSelected = mnode
            mnode.Selected = True
            Me.PopupMenu MDIMain.mnuPopup
            
            tvwSimu_NodeClick mnode
        End If
    End If
End Sub



Sub CreaAlbero()
    Dim tvwItem As Node
    Dim rs As Recordset
    Dim i As Integer
    
    Set rs = gDBC.Execute("select * from SelSimulazioni")
        
    tvwSimu.Sorted = True
    
    Do While Not rs.EOF
        If Not existsItem(tvwSimu, getKey(rs("codComune"))) Then
            tvwSimu.Nodes.Add , , getKey(rs("codComune")), getDesComune(rs("codComune")) & " :" & rs("codComune")
        End If
        
        If Not IsNull(rs("CodTipoTariffa")) Then
            If Not existsItem(tvwSimu, getKey(rs("codComune") & rs("CodTipoTariffa"))) Then
                tvwSimu.Nodes.Add getKey(rs("codComune")), tvwChild, getKey(rs("codComune") & rs("CodTipoTariffa")), rs("Descrizione")
            End If
            
            If Not existsItem(tvwSimu, getKey(rs("codComune") & rs("CodTipoTariffa") & rs("AnnoRiferimento"))) Then
                tvwSimu.Nodes.Add getKey(rs("codComune") & rs("CodTipoTariffa")), tvwChild, getKey(rs("codComune") & rs("CodTipoTariffa") & rs("AnnoRiferimento")), rs("AnnoRiferimento")
            End If
            
            If Not IsNull(rs("idsimulazione")) Then
                If Not existsItem(tvwSimu, getKey(rs("codComune") & rs("CodTipoTariffa") & rs("AnnoRiferimento") & rs("idSimulazione"))) Then
                    tvwSimu.Nodes.Add getKey(rs("codComune") & rs("CodTipoTariffa") & rs("AnnoRiferimento")), tvwChild, getKey(rs("codComune") & rs("CodTipoTariffa") & rs("AnnoRiferimento") & rs("idSimulazione")), rs("DesSimulazione")
                    tvwSimu.Nodes.Add getKey(rs("codComune") & rs("CodTipoTariffa") & rs("AnnoRiferimento") & rs("idSimulazione")), tvwChild, , "Coefficienti"
                    tvwSimu.Nodes.Add getKey(rs("codComune") & rs("CodTipoTariffa") & rs("AnnoRiferimento") & rs("idSimulazione")), tvwChild, , "Conto Economico"
                End If
            End If
        End If
        
        rs.MoveNext
    Loop
    
    If tvwSimu.Tag <> vbNullString Then
        If existsItem(tvwSimu.Nodes, tvwSimu.Tag) Then
            tvwSimu.Nodes(tvwSimu.Tag).EnsureVisible
        End If
    End If
End Sub



Private Sub tvwSimu_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim mnode As Node
    Dim s() As String
    Dim i As Integer
    
    Set mnode = Node
    On Error Resume Next
    i = -1
    
    s = Split(mnode.FullPath, "\")
    i = UBound(s)
    
    Command1(CMD_CANCELLA).Enabled = i = 3
    Command1(CMD_PUBBLICA).Enabled = i = 3
    Command1(CMD_CREA).Enabled = i = 2
    Command1(CMD_REPORT).Enabled = i = 3
    Command1(CMD_PRINT).Enabled = i = 2
    
    If Not currentObject Is Nothing Then
        currentObject.Visible = False
        SetParent currentObject.hwnd, pLastParent
        Unload currentObject
        Set currentObject = Nothing
    End If
    
    If Not currentObjectEditChild Is Nothing Then
        currentObjectEditChild.Visible = False
        SetParent currentObjectEditChild.hwnd, pLastParent
        Unload currentObjectEditChild
        Set currentObjectEditChild = Nothing
    End If
    
    If i = 2 Then
        tvwSimu.Tag = mnode.Key
        Set pNodeSelected = Node
        If i = 2 Then
            caricaCalcolo i
        End If
    ElseIf i = 4 And mnode.Text = "Coefficienti" Then
        Set pNodeSelected = Node.Parent
        caricaCoeff
    ElseIf i = 4 And mnode.Text = "Conto Economico" Then
        Set pNodeSelected = Node.Parent
        caricaContoEconomico
    ElseIf i = 3 Then
        Set pNodeSelected = Node
        caricaCalcolo i
    End If
    
    
End Sub

Sub caricaCoeff()
    Dim FrmCoeff As FrmCoefficiente
    
    Set FrmCoeff = New FrmCoefficiente
    
    pLastParent = 0
    
    
    SetParent FrmCoeff.hwnd, Picture1.hwnd
    
    FrmCoeff.Carica Mid$(pNodeSelected.Key, 2, 6), Right$(pNodeSelected.Key, 6), Mid$(pNodeSelected.Key, 8, 1), False
    TitleBar FrmCoeff, False
    
    Set currentObject = FrmCoeff
    
    Me.SetFocus
    
    Picture1_Resize
End Sub

Sub caricaContoEconomico()
    Dim frmCostiDelCom As FrmCostiDelComune
    
    Set frmCostiDelCom = New FrmCostiDelComune
    
    pLastParent = 0
   
    frmCostiDelCom.Carica Mid$(pNodeSelected.Key, 2, 6), Right$(pNodeSelected.Key, 6), Mid$(pNodeSelected.Key, 8, 1), False, Picture1.hwnd
    
    Set currentObjectEditChild = frmCostiDelCom.GetEditForm
    
    TitleBar currentObjectEditChild, False
    
    Set currentObject = frmCostiDelCom
    
    Me.SetFocus
    
    Picture1_Resize
End Sub



Sub caricaCalcolo(index As Integer)
    Dim CodComune As String
    Dim Anno As String
    Dim TipoTariffa As Long
    
    pLastParent = 0
    
    SetParent frmSimuCalcolo.hwnd, Picture1.hwnd

    TitleBar frmSimuCalcolo, False
    
    If index = 2 Then
        CodComune = Mid$(pNodeSelected.Key, 2, 6)
        Anno = Right$(pNodeSelected.Key, 4)
        TipoTariffa = Mid$(pNodeSelected.Key, 8, 1)
    Else
        CodComune = Mid$(pNodeSelected.Key, 2, 6)
        Anno = Right$(pNodeSelected.Key, 6)
        TipoTariffa = Mid$(pNodeSelected.Key, 8, 1)
    End If
    frmSimuCalcolo.Carica CodComune, Anno, TipoTariffa
    frmSimuCalcolo.Visible = False
    
    Set currentObject = frmSimuCalcolo
    
    Me.SetFocus
    
    Picture1_Resize
    
    frmSimuCalcolo.Visible = True
End Sub

