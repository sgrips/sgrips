VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLottiLottoUscita 
   Caption         =   "Form1"
   ClientHeight    =   8310
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   7920
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd 
      Caption         =   "Chiudi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin MSComctlLib.TreeView tvwLottiScaricati 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4048
      _Version        =   393217
      LabelEdit       =   1
      Style           =   6
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmLottiLottoUscita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private offsetX As Long
Private offsetY As Long
Private pCodLottoUscita As String


Public Sub showMe(codLottoUscita As String)
    pCodLottoUscita = codLottoUscita
    
    caricaLotti
    
    Me.Show vbModal
End Sub


Private Sub cmd_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    offsetX = 10 * Screen.TwipsPerPixelX
    offsetY = 10 * Screen.TwipsPerPixelX
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    Dim posLeft As Long
    Dim posTop As Long
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    On Error Resume Next
    posTop = Me.ScaleHeight - cmd(i).Height - offsetX
    posLeft = Me.ScaleWidth - cmd(i).width - offsetX
    
    For i = cmd.LBound To cmd.ubound
        If i > 0 Then posLeft = posLeft - offsetX - cmd(i).width
        
        cmd(i).Move posLeft, posTop
    Next
    
     
    tvwLottiScaricati.Move 0, 0, Me.ScaleWidth, cmd(0).Top - offsetY
End Sub


Private Sub caricaLotti()
    Dim root As String
    Dim nodeT As Node
    
    If (DevEnvMDB.rscmdLottiLottoUscita.state = adStateOpen) Then
        DevEnvMDB.rscmdLottiLottoUscita.Close
    End If
    
    DevEnvMDB.cmdLottiLottoUscita pCodLottoUscita
    
    root = DevEnvMDB.rscmdLottiLottoUscita!LottoUscita
    tvwLottiScaricati.Nodes.Clear
    Set nodeT = tvwLottiScaricati.Nodes.Add(, TreeRelationshipConstants.tvwFirst, root, root)
    nodeT.Expanded = True
    
    With DevEnvMDB.rscmdLottiLottoUscita
        Do While Not .EOF
            tvwLottiScaricati.Nodes.Add root, TreeRelationshipConstants.tvwChild, root & "|" & .Fields("CodLotto"), .Fields("Descrizione")
           .MoveNext
        Loop
        
    End With
    
End Sub
