VERSION 5.00
Begin VB.Form frmFiltroMovimentiScuole 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filtro Movimenti Scuole"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   2520
      Width           =   6375
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
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1800
      Width           =   6375
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
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1080
      Width           =   6375
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
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   360
      Width           =   6375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Applica"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   2565
      TabIndex        =   0
      Top             =   3600
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sezioni"
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
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Classi"
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
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Scuole"
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
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Istituto"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   585
   End
End
Attribute VB_Name = "frmFiltroMovimentiScuole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pRetValue(3) As Long

Private Sub Combo1_Click(Index As Integer)

    If Combo1(Index).ListIndex = 0 Then Exit Sub

    CaricaCombo Combo1(Index + 1), Choose(Index + 1, "S", "C", "Z"), Combo1(Index).ItemData(Combo1(Index).ListIndex)
End Sub

Private Sub Command1_Click()
    Dim i As Long

    For i = Combo1.LBound To Combo1.UBound
        If Combo1(i).ListIndex > 0 Then
            pRetValue(i) = Combo1(i).ItemData(Combo1(i).ListIndex)
        Else
            pRetValue(i) = -1
        End If
    Next


    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    For i = Combo1.LBound To Combo1.UBound
        Combo1(i).Clear
        Combo1(i).Visible = i + 1 <= gPuntiAssegnatiPer
        Label1(i).Visible = i + 1 <= gPuntiAssegnatiPer
    Next
    
    CaricaCombo Combo1(0), "I"
    CaricaCombo Combo1(1), "S"
    CaricaCombo Combo1(2), "C"
    CaricaCombo Combo1(3), "Z"
End Sub



Private Sub CaricaCombo(cmb As ComboBox, TipoUtente As String, Optional codPadre As Long = 0)
    Dim rs As ADODB.Recordset
    
    cmb.Clear
    cmb.AddItem "Tutti"
    
    If codPadre <> 0 Then
        Set rs = gDBC.Execute("select * from Utenti where TipoAnagrafe='" & TipoUtente & "' and codPadre=" & codPadre)
    Else
        Set rs = gDBC.Execute("select * from Utenti where TipoAnagrafe='" & TipoUtente & "'")
    End If
    
    Do While Not rs.EOF
        cmb.AddItem rs("RagioneSociale").Value
        cmb.ItemData(cmb.ListCount - 1) = rs("CodUtente")
        rs.MoveNext
    Loop
    cmb.ListIndex = 0
    rs.Close
End Sub

Public Function ShowMe() As Variant
    Dim i As Long
    Dim retValue(3) As Long
    
    Me.Show vbModal
    
    
    CopyMemory retValue(0), pRetValue(0), LenB(retValue(0)) * (UBound(retValue) + 1)
    
    ShowMe = retValue
    
End Function
