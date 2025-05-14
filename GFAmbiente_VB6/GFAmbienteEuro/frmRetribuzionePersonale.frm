VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRetribuzionePersonale 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Retribuzione del Personale"
   ClientHeight    =   6105
   ClientLeft      =   2085
   ClientTop       =   2280
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imglist 
      Left            =   7275
      Top             =   3765
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRetribuzionePersonale.frx":0000
            Key             =   "Icon_Root_Personale"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRetribuzionePersonale.frx":4DBC
            Key             =   "Icon_TipoPersonale"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtValore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6360
      TabIndex        =   3
      Text            =   "0"
      Top             =   2430
      Width           =   3480
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Annulla"
      Height          =   405
      Index           =   1
      Left            =   7350
      TabIndex        =   2
      Top             =   5595
      Width           =   1170
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Conferma"
      Height          =   405
      Index           =   0
      Left            =   8685
      TabIndex        =   1
      Top             =   5595
      Width           =   1185
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   5985
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   10557
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Personale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   4
      Left            =   6330
      TabIndex        =   8
      Top             =   900
      Width           =   3510
   End
   Begin VB.Label lbl 
      Caption         =   "Tipo Personale:"
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   3
      Left            =   6300
      TabIndex        =   7
      Top             =   660
      Width           =   3390
   End
   Begin VB.Label lbl 
      Caption         =   "Livello:"
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   2
      Left            =   6315
      TabIndex        =   6
      Top             =   1410
      Width           =   3390
   End
   Begin VB.Label lbl 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Livello"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   1
      Left            =   6330
      TabIndex        =   5
      Top             =   1650
      Width           =   3510
   End
   Begin VB.Label lbl 
      Caption         =   "Retribuzione Annua:"
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   0
      Left            =   6360
      TabIndex        =   4
      Top             =   2190
      Width           =   3390
   End
End
Attribute VB_Name = "frmRetribuzionePersonale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT = "ROOT"
Private Const Icon_Root_Personale = "Icon_Root_Personale"
Private Const Icon_TipoPersonale = "Icon_TipoPersonale"
Private Const Icon_LivelloPersonale = "Icon_TipoPersonale"

Private Const KEYTAGDIRIGENTE = SEPARATORE & "DIRIGENTE"
Private Const KEYTAGRETRIBUZIONE = SEPARATORE & "RETRIBUZIONE"

Private Const CMD_CONFERMA = 0
Private Const CMD_ANNULLA = 1

Private Const SEPARATORE_DATO_NODE = " - "

Private OldTxtValue As String

Private mComune As String
Private mAnnoRiferimento As String

Public Sub Carica(Comune As String, AnnoRiferimento As String)
    
    mComune = Comune
    mAnnoRiferimento = AnnoRiferimento
    Me.Show vbModal
End Sub


Private Sub Cmd_Click(Index As Integer)
    Select Case Index
        Case 0
            SalvaRetribuzioni
            Unload Me
        Case 1
            Unload Me
    End Select
    
    
End Sub

Private Sub Form_Load()

    Me.Caption = mComune & " - Retribuzione del personale"
    tvw.ImageList = imglist
    
    If Not CaricaTree Then Unload Me
End Sub


Private Function CaricaTree()
    Dim rs As Recordset
    Dim rs1 As Recordset
    Dim Node As Node
    Dim KEYNODE As String
    Dim Retribuzione As Double
    
    Set rs = gDBC.Execute("Select TipiPersonaleImpiegato.* from TipiPersonaleImpiegato inner join PersonaleImpiegatoNelComune ON TipiPersonaleImpiegato.CodTipoPersonale = PersonaleImpiegatoNelComune.CodTipoPersonale ")
    Set rs1 = gDBC.Execute("Select * from LivelliPersonaleImpiegato")
    
    On Error GoTo CaricaTreeERR
    CaricaTree = False
    
    If rs.EOF Then
        MsgBox "Attenzione la tabella: TipiPersonaleImpiegato " + vbCrLf + "è vuota, controllare il DataBase", vbCritical + vbOKOnly
        Exit Function
    End If
    
    If rs1.EOF Then
        MsgBox "Attenzione la tabella: LivelliPersonaleImpiegato " + vbCrLf + "è vuota, controllare il DataBase", vbCritical + vbOKOnly
        Exit Function
    End If
    
    
    With tvw.Nodes
        Set Node = .Add(, , ROOT, "Personale", Icon_Root_Personale)
        Node.Expanded = True
        
        'carico nel treeview la tipologia del personale
        Do While Not rs.EOF
            Set Node = .Add(ROOT, tvwChild, KEYLIV1 & rs("codTipoPersonale"), rs("Descrizione"), Icon_TipoPersonale)
                   
            Node.Tag = SetValueInTag(KEYTAGDIRIGENTE, rs("dirigente"), Node.Tag)
            KEYNODE = Node.Key
            
            'Aggiungo i livelli per il personale
            Do While Not rs1.EOF
                If rs("Dirigente") = rs1("Dirigente") Then
                    Retribuzione = GetRetribuzione(mComune, mAnnoRiferimento, rs("codTipoPersonale"), rs1("CodLibelloPersonale"))
                    
                    Set Node = .Add(KEYNODE, tvwChild, KEYNODE & KEYLIV2 & rs1("CodLibelloPersonale"), rs1("Descrizione") & SEPARATORE_DATO_NODE & Format$(Retribuzione, FORMATVALUTA), Icon_LivelloPersonale)
                    
                    Node.Tag = SetValueInTag(KEYTAGDIRIGENTE, rs1("Dirigente"), Node.Tag)
                    Node.Tag = SetValueInTag(KEYTAGRETRIBUZIONE, Retribuzione, Node.Tag)
                End If
                rs1.MoveNext
            Loop
            
            rs1.MoveFirst
            rs.MoveNext
        Loop
        
        
    End With
    
    CaricaTree = True
    Exit Function
CaricaTreeERR:
    MsgBox Error$
End Function

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
        
    If InStr(1, Node.Key, ROOT) > 0 Then Exit Sub
    
    If InStr(1, Node.Key, KEYLIV2) = 0 Then
        'Se primo livello cambio Lbl Tipo Personale
        lbl(4) = Node.Text
        txtValore.Tag = vbNullString
        txtValore.Enabled = False
    Else
        'retribuzione
        i = InStr(1, Node.Text, SEPARATORE_DATO_NODE)
        lbl(4) = Node.Parent.Text
        lbl(1) = Left$(Node.Text, i - 1)
        txtValore = Format$(GetValueInTag(KEYTAGRETRIBUZIONE, Node.Tag), FORMATVALUTA)
        
        txtValore.Tag = Node.Key
        txtValore.Enabled = True
    End If
End Sub


Private Sub SalvaRetribuzioni()
    Dim oNode As Node
    Dim onNode As Node
    Dim i As Integer
    Dim Retribuzione As Long
    
    'Se Retribuzione 0 cancello livello
    'se >0 la Inserisco se nuova altrimenti la Modifico
    With Me.tvw
    Set oNode = .Nodes.Item(1).Child
        For i = 1 To .Nodes.Item(1).children
            Set onNode = oNode.Child
            Do While Not onNode Is Nothing

                Retribuzione = GetRetribuzione(mComune, mAnnoRiferimento, GetValueOfKeyTVW(KEYLIV1, onNode.Key), GetValueOfKeyTVW(KEYLIV2, onNode.Key))
                 
                Select Case CDbl(GetValueInTag(KEYTAGRETRIBUZIONE, onNode.Tag))
                    Case 0
                        If Retribuzione > 0 Then gDBC.Execute "Delete from RetribuzioneDelPersonale Where CodComune='" & mComune & "' and CodTipoPersonale =" & GetValueOfKeyTVW(KEYLIV1, onNode.Key) & " and CodLivelloPersonale ='" & GetValueOfKeyTVW(KEYLIV2, onNode.Key) & "'"
                    Case Is > 0
                        If Retribuzione = 0 Then
                            gDBC.Execute "Insert Into RetribuzioneDelPersonale (RetribuzioneAnnuale,AnnoRiferimento,CodComune,CodTipoPersonale,CodLivelloPersonale) values(" & FormatFLD(CDbl(GetValueInTag(KEYTAGRETRIBUZIONE, onNode.Tag)), adDouble) & ",'" & mAnnoRiferimento & "','" & mComune & "' ," & GetValueOfKeyTVW(KEYLIV1, onNode.Key) & " ,'" & GetValueOfKeyTVW(KEYLIV2, onNode.Key) & "')"
                        Else
                            gDBC.Execute "Update RetribuzioneDelPersonale set RetribuzioneAnnuale=" & FormatFLD(CDbl(GetValueInTag(KEYTAGRETRIBUZIONE, onNode.Tag)), adDouble) & "  Where  CodComune='" & mComune & "' and AnnoRiferimento='" & mAnnoRiferimento & "' and CodTipoPersonale =" & GetValueOfKeyTVW(KEYLIV1, onNode.Key) & " and CodLivelloPersonale ='" & GetValueOfKeyTVW(KEYLIV2, onNode.Key) & "'"
                        End If
                End Select
                Set onNode = onNode.Next
            Loop
            Set oNode = oNode.Next
        Next i
    End With
End Sub


Private Sub txtValore_GotFocus()
    'Imposto il numero MAX di caratteri in EDIT MODE
    txtValore = CDbl(txtValore.Text)
    txtValore.MaxLength = 12
    'selezione testo
    txtValore.SelStart = 0
    txtValore.SelLength = Len(txtValore.Text)
     OldTxtValue = txtValore

End Sub

Private Sub txtValore_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
End Sub

Private Sub txtValore_LostFocus()
    Dim i As Integer
    
    txtValore.MaxLength = 15
    If txtValore.Tag <> vbNullString Then
        i = InStr(1, tvw.Nodes(txtValore.Tag).Text, SEPARATORE_DATO_NODE)
        tvw.Nodes(txtValore.Tag).Text = Left$(tvw.Nodes(txtValore.Tag).Text, i + Len(SEPARATORE_DATO_NODE) - 1) & Format$(CDbl(txtValore.Text), FORMATVALUTA)
        tvw.Nodes(txtValore.Tag).Tag = SetValueInTag(KEYTAGRETRIBUZIONE, CDbl(txtValore.Text), tvw.Nodes(txtValore.Tag).Tag)
    End If

End Sub



Private Sub txtValore_Validate(Cancel As Boolean)
        'Il valore deve essere numerico
        If Not IsNumeric(txtValore) Then MsgBox "Attenzione il valore deve essere numerico", vbOKOnly + vbExclamation, App.Title: Cancel = True: txtValore = OldTxtValue: Exit Sub

End Sub
