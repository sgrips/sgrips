VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBonus 
   Caption         =   "Bonus"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   9060
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmdTipiTariffa 
      Height          =   315
      Left            =   5220
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Frame frm 
      Caption         =   "Bonus"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2910
      Left            =   5220
      TabIndex        =   15
      Top             =   225
      Width           =   3690
      Begin VB.CommandButton Command1 
         Caption         =   "Riduzione Single +PF"
         Enabled         =   0   'False
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   1890
      End
      Begin VB.TextBox txt 
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
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   2400
         Width           =   3435
      End
      Begin VB.CommandButton Command1 
         Caption         =   "+ PF"
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   2355
         TabIndex        =   7
         Top             =   1680
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aggiorna"
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   105
         TabIndex        =   4
         Top             =   1215
         Width           =   930
      End
      Begin VB.CommandButton Command1 
         Caption         =   "+ PV"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   2985
         TabIndex        =   8
         Top             =   1680
         Width           =   570
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cancella/Ripristina"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   1350
         TabIndex        =   5
         Top             =   1215
         Width           =   2205
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
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
         Index           =   0
         Left            =   150
         TabIndex        =   0
         Text            =   "0"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   2415
         TabIndex        =   2
         Text            =   "0"
         Top             =   600
         Width           =   855
      End
      Begin MSComCtl2.UpDown UpDwn 
         Height          =   420
         Index           =   1
         Left            =   3270
         TabIndex        =   3
         Top             =   600
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   741
         _Version        =   393216
         Value           =   10
         BuddyControl    =   "txt(1)"
         BuddyDispid     =   196612
         BuddyIndex      =   1
         OrigLeft        =   1305
         OrigTop         =   1560
         OrigRight       =   1545
         OrigBottom      =   1980
         Max             =   90
         Min             =   10
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDwn 
         Height          =   420
         Index           =   0
         Left            =   1650
         TabIndex        =   1
         Top             =   600
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   741
         _Version        =   393216
         BuddyControl    =   "txt(0)"
         BuddyDispid     =   196612
         BuddyIndex      =   0
         OrigLeft        =   1335
         OrigTop         =   645
         OrigRight       =   1575
         OrigBottom      =   1065
         Max             =   100000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Descrizione"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   18
         Top             =   2160
         Width           =   1065
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Valore "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   17
         Top             =   345
         Width           =   645
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "% Riduzione"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2400
         TabIndex        =   16
         Top             =   345
         Width           =   1125
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "A&nnulla"
      Default         =   -1  'True
      Height          =   375
      Index           =   4
      Left            =   6555
      TabIndex        =   13
      Top             =   5400
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C&onferma"
      Height          =   375
      Index           =   3
      Left            =   5280
      TabIndex        =   12
      Top             =   5400
      Width           =   1140
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   7635
      Left            =   105
      TabIndex        =   14
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   13467
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   531
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker DTPAnno 
      Height          =   405
      Left            =   5265
      TabIndex        =   11
      Top             =   4920
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy"
      Format          =   48562179
      UpDown          =   -1  'True
      CurrentDate     =   36967
   End
End
Attribute VB_Name = "frmBonus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT = "Root"
Private Const UA = "UA"
Private Const UD = "UD"
Private Const PF = "PF"
Private Const PV = "PV"
Private Const SEPARATOREVALORE = " - "
Private Const SEPARATOREDESCRIZIONE = " >"

Private mAnnoRiferimento As String
Private mCodComune As String
Private mTipoTariffa As Integer

Private bordo As Long
Private oNode As Node
Private Livello As Long

Public Enum tpoParte
    ParteFissa
    ParteVariabile
End Enum

Private noEvent As Boolean

Private Sub cmdTipiTariffa_Click()
    If noEvent Then Exit Sub
    mTipoTariffa = cmdTipiTariffa.ItemData(cmdTipiTariffa.ListIndex)
    Carica mCodComune, mAnnoRiferimento
End Sub

Private Sub Command1_Click(index As Integer)
    Dim oNodeP As Node
    Dim oNodeP1 As Node
    Dim i As Integer
    
    Select Case index
        Case 0
            'Agg PV
            GestioneNodo ParteVariabile, False
            
        Case 1
            'Cancella
            If Not oNode Is Nothing Then
            
                Select Case Livello
                    Case 3
                       For i = 1 To 2
                            If i = 1 Then
                                Set oNodeP1 = oNode.Child
                            Else
                                Set oNodeP1 = oNodeP1.Next
                            End If
                            
                            Set oNodeP = oNodeP1.Child
                            Do While Not oNodeP Is Nothing
                                oNodeP.ForeColor = IIf(oNodeP.ForeColor = vbRed, vbBlack, vbRed)
                                Set oNodeP = oNodeP.Next
                            Loop
                        Next
                    
                    Case 4
                        Set oNodeP = oNode.Child
                        Do While Not oNodeP Is Nothing
                            oNodeP.ForeColor = IIf(oNodeP.ForeColor = vbRed, vbBlack, vbRed)
                            Set oNodeP = oNodeP.Next
                        Loop
                    Case 5
                        oNode.ForeColor = IIf(oNode.ForeColor = vbRed, vbBlack, vbRed)
                End Select
                
                                
            End If
        Case 2
            'Modifica
            If oNode.Parent.Text = PF Then
                GestioneNodo ParteFissa, True
            Else
                GestioneNodo ParteVariabile, True
            End If
            
        Case 3
            'Conferma
            SalvaSuDB
            Unload Me
            Exit Sub
        Case 4
            'Annulla
            Unload Me
            Exit Sub
        Case 5
            'Agg PF
            GestioneNodo ParteFissa, False
            
        Case 6
            'Agg PF Single
            GestioneNodo ParteFissa, False, RIDUZIONE_SINGLE
            
    End Select

    txt(2).Enabled = index <> 6
End Sub

Private Sub DTPAnno_Change()
    mAnnoRiferimento = Year(DTPAnno.value)
    mTipoTariffa = GetParamValue(mAnnoRiferimento, mCodComune, "TIPOMETODO")
    If cmdTipiTariffa.ListCount > 0 Then
        On Error Resume Next
        cmdTipiTariffa.ListIndex = mTipoTariffa
    End If
    Carica mCodComune, mAnnoRiferimento
End Sub

Private Sub Form_Load()
    noEvent = True
    MDIMain.mnuComprimi.Enabled = False
    Me.Width = 11500
    Set tvw.ImageList = MDIMain.imglist
    bordo = (Me.Width - Me.ScaleWidth) / 2
    DTPAnno.value = Format$(Now, "yyyy-mm-dd")
    mTipoTariffa = GetParamValue(mAnnoRiferimento, mCodComune, "TIPOMETODO")
    
    'selezione tipi tariffa
    Dim rs As ADODB.Recordset
    Set rs = gDBC.Execute("select * from tipitariffa where visibile=-1 order by codTipoTariffa")
    Dim i As Integer
    
    i = 0
    Do While Not rs.EOF
        cmdTipiTariffa.AddItem rs("Descrizione")
        cmdTipiTariffa.ItemData(cmdTipiTariffa.ListCount - 1) = rs("codTipoTariffa")
        If rs("codTipoTariffa") = mTipoTariffa Then
            cmdTipiTariffa.ListIndex = i
        End If
        i = i + 1
        rs.MoveNext
    Loop
    If cmdTipiTariffa.ListIndex = -1 Then cmdTipiTariffa.ListIndex = 0
    noEvent = False
End Sub

Sub Carica(Optional ByVal sCodComune As String, Optional ByVal sAnno As String)
    Dim oNodeP As Node
    mAnnoRiferimento = IIf(sAnno = vbNullString, Year(Now), sAnno)
    
    'se la modalità è tutti i comuni chiedo di quale comune si vanno ad inserire i dati
    If sCodComune = vbNullString Then
        If gComuneAttivo = ALL_COMUNI Then
            frmComuneDaLavorare.Show vbModal
            sCodComune = GetSetting(App.Title, SEZ_ULTIMI_VALORI, KEY_ULTIMI_VALORI_COMUNESEL, "")
        Else
            sCodComune = gComuneAttivo
        End If
    End If
    
    mCodComune = sCodComune
    
    Load Me
    
    Me.Caption = "Riduzioni:" & getDesComune(sCodComune)
    
    tvw.Visible = False
    tvw.Nodes.Clear
    Set oNodeP = tvw.Nodes.Add(, tvwFirst, ROOT, "Utenze", ICON_DOMESTICA_ATTIVITA)
    oNodeP.Expanded = True
    Set oNodeP = tvw.Nodes.Add(ROOT, tvwChild, UA, "Utenze Attivita", ICON_ATTIVITA)
    oNodeP.Expanded = True
    Set oNodeP = tvw.Nodes.Add(ROOT, tvwChild, UD, "Utenze Domestica", ICON_DOMESTICA)
    oNodeP.Expanded = True
    
    
    '-----------------------------------------------------------
    'Carico dati UTENZA ATTIVITA
    '-----------------------------------------------------------
    CreaAlbero BonusUtenzaAttivita
    
    '-----------------------------------------------------------
    'Carico dati UTENZA DOMESTICA
    '-----------------------------------------------------------
    CreaAlbero BonusUtenzaDomestica

    tvw.Visible = True
    
    mTipoTariffa = cmdTipiTariffa.ItemData(cmdTipiTariffa.ListIndex)
    
    If Not Me.Visible Then
        Me.Show vbModal
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        'Cancello quelli con percentuale a ZERO
        gDBC.Execute ("Delete from BonusUtenzeAttivita  where Riduzione=0")
        gDBC.Execute ("Delete from BonusUtenzeDomestiche  where Riduzione=0")

End Sub

Private Sub Form_Resize()
    If WindowState = vbMinimized Then Exit Sub

On Error Resume Next
    frm.Move Me.ScaleWidth - frm.Width - bordo, bordo
    
    If frm.Left - bordo < 0 Then Exit Sub
    
    tvw.Move Me.ScaleLeft, Me.ScaleTop, frm.Left - bordo, Me.ScaleHeight
    
    Command1(4).Move Me.ScaleWidth - Command1(4).Width - bordo, Me.ScaleHeight - bordo - Command1(4).Height, Command1(4).Width
    Command1(3).Move Command1(4).Left - Command1(4).Width - bordo, Command1(4).Top
    DTPAnno.Move Command1(3).Left - Command1(3).Width, Command1(3).Top
    
    cmdTipiTariffa.Move frm.Left, frm.Top + frm.Height + bordo, frm.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.mnuComprimi.Enabled = True
End Sub



Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim UAoUD As String
    
    Set oNode = Nothing
    txt(0) = 0
    txt(1) = 0
    frm.Caption = "Bonus"
    Livello = CheLivello(Node.FullPath)
    lbl(0) = "Valore"
    Command1(0).Enabled = False
    Command1(5).Enabled = False
    Command1(1).Enabled = False
    Command1(2).Enabled = False
    
    txt(2).Enabled = True
    UAoUD = ""
    
    Select Case Livello
        Case 1
        Case 2
        Case 3
            Set oNode = Node
            UAoUD = oNode.Parent.Key
            
            
            Command1(0).Enabled = mTipoTariffa <> tpoTariffa.Tariffa507
            Command1(5).Enabled = True
            Command1(1).Enabled = True
            Command1(2).Enabled = False
        Case 4
            Command1(1).Enabled = True
            Set oNode = Node
            UAoUD = oNode.Parent.Parent.Key
            If Node.Text = "PV" Then
                Command1(0).Enabled = mTipoTariffa <> tpoTariffa.Tariffa507
                Command1(5).Enabled = False
            Else
                Command1(5).Enabled = True
                
                Command1(0).Enabled = False
            End If
        Case 5
            
            Set oNode = Node
            UAoUD = oNode.Parent.Parent.Parent.Key
            Command1(1).Enabled = True
            Command1(2).Enabled = True
            txt(0) = PrendiValore(Node.Text)
            txt(1) = PrendiPercentuale(Node.Text)
            txt(2) = PrendiDescrizione(Node.Text)
            'testo se cancellato logicamente
            If Node.Parent.Text = "PV" Then
                frm.Caption = "Bonus - Parte Variabile"
                lbl(0) = IIf(Node.Parent.Parent.Parent.Key = UA, "Superfice", "Nr.Famiglie")
                Command1(0).Enabled = mTipoTariffa <> tpoTariffa.Tariffa507
                Command1(5).Enabled = False
            Else
                frm.Caption = "Bonus - Parte Fissa"
                lbl(0) = "Superfice"
                Command1(5).Enabled = True
                Command1(0).Enabled = False
            End If
    End Select
    
    Command1(6).Visible = False
    'Command1(6).Enabled = Command1(5).Enabled And UAoUD = UD And mTipoTariffa = tpoTariffa.Tariffa507
    'If Command1(6).Enabled And Txt(2).Text = RIDUZIONE_SINGLE Then
    '    Txt(2).Enabled = False
    'End If
End Sub

Function CheLivello(ByVal s As String) As Long
    Dim i As Long
    
    i = InStr(1, s, Me.tvw.PathSeparator)
    CheLivello = 1
    Do While i > 0
        CheLivello = CheLivello + 1
        s = Mid$(s, i + 1)
        i = InStr(1, s, Me.tvw.PathSeparator)
    Loop
End Function


Function PrendiValore(ByVal s As String) As String
    Dim i As Long
    Dim ii As Long
    
    ii = InStr(1, s, SEPARATOREVALORE)
    i = InStrRev(s, SEPARATOREDESCRIZIONE)
    PrendiValore = Mid$(Replace(s, ",", "."), ii + Len(SEPARATOREVALORE), i - ii - Len(SEPARATOREVALORE))

End Function


Function PrendiPercentuale(ByVal s As String) As String
    Dim i As Long
    i = InStr(1, s, "%")
    
    PrendiPercentuale = Left$(s, i - 1)
End Function


Function PrendiSingle(ByVal s As String) As String
    Dim i As Long
    i = InStr(1, s, "%")
    
    PrendiSingle = Left$(s, i - 1)
End Function


Function PrendiDescrizione(ByVal s As String) As String
    Dim i As Long
    
    i = InStr(1, s, SEPARATOREDESCRIZIONE)
    PrendiDescrizione = Mid$(s, i + Len(SEPARATOREDESCRIZIONE))
    
End Function


Sub CreaAlbero(ByVal Tipo As TpoBonus)
    Dim oNodeP As Node
    Dim oNodeF As Node
    Dim oNodeV As Node
    Dim i As Long
    Dim KEYNODE As String
    Dim Key As String
    Dim rs As Recordset
    Dim qry As Command
    Dim sQry As String
    Dim sTable As String
    Dim pKeySave As String
    Dim pKey As String
    Dim sKeySingle As String
    
    Set oNodeP = Nothing
    Set oNodeF = Nothing
    Set oNodeV = Nothing
    
    If Tipo = BonusUtenzaAttivita Then
        sTable = "BonusUtenzeAttivita"
        sQry = "InsCategorieAttivitaInBonusCA"
    Else
        sQry = "InsUtenzeDomesticheInBonusUD"
        sTable = "BonusUtenzeDomestiche"
    End If
    
    'Inserisce se vi sono le nuove Utenze Domestiche
    Set qry = gADOX.Procedures(sQry).Command
    qry.Parameters.Refresh
    qry.Parameters("pCodComune") = mCodComune
    qry.Parameters("pAnnoRiferimento") = mAnnoRiferimento
    qry.Parameters("pCodTipoTariffa") = mTipoTariffa
    qry.Execute
    
    Set rs = SelezionaBonus(Tipo, mCodComune, mAnnoRiferimento, mTipoTariffa)

    pKeySave = CreaLaPkey(sTable)
    
    Do While Not rs.EOF
        
        On Error Resume Next
        
        
        
        If Tipo = BonusUtenzaAttivita Then
            KEYNODE = UA
            Key = rs("codAttivita")
        Else
            KEYNODE = UD
            Key = rs("IdNumeriComponenti")
        
            'If mTipoTariffa = tpoTariffa.Tariffa507 And Key <> "1" Then
            '    Exit Sub
            'End If
        End If
        
        
        Set oNode = tvw.Nodes(KEYNODE & SEPARATORE & "Liv1" & SEPARATORE & Key)
        If Err > 0 Then
            Set oNodeP = tvw.Nodes.Add(KEYNODE, tvwChild, KEYNODE & SEPARATORE & "Liv1" & SEPARATORE & Key, rs("Descrizione"), IIf(KEYNODE = UA, ICON_FACTORY, ICON_MAN))
            oNodeP.Expanded = False
            Err.Clear
            
            pKey = pKeySave
            For i = 0 To rs.fields.Count - 1
                If InStr(pKeySave, SEPARATORE & rs.fields(i).Name) And rs.fields(i).Name <> "Riduzione" And rs.fields(i).Name <> "ParteFissa" And rs.fields(i).Name <> "indSingle" Then
                    MettivaloreInPKEY pKey, SEPARATORE & rs.fields(i).Name, rs.fields(i).value, rs.fields(i).Type
                End If
            Next
            oNodeP.Tag = pKey
            oNodeP.Sorted = True
        End If
        
        If rs("riduzione") > 0 Then
            Set oNode = tvw.Nodes(oNodeP.Key & "Liv2" & SEPARATORE & PF)
            If Err > 0 And rs("ParteFissa") = True Then
                Err.Clear
                Set oNodeF = tvw.Nodes.Add(oNodeP.Key, tvwChild, oNodeP.Key & "Liv2" & SEPARATORE & PF, "PF", ICON_PF)
                oNodeF.Expanded = False
                oNodeF.Sorted = True
            End If
            
            Set oNode = tvw.Nodes(oNodeP.Key & "Liv2" & SEPARATORE & PV)
            If Err > 0 And rs("ParteFissa") = False Then
                Err.Clear
                Set oNodeV = tvw.Nodes.Add(oNodeP.Key, tvwChild, oNodeP.Key & "Liv2" & SEPARATORE & PV, "PV", ICON_PV)
                oNodeV.Expanded = False
                oNodeV.Sorted = True
            End If
            
            
            If Tipo = BonusUtenzaDomestica Then
                sKeySingle = SEPARATORE & IIf(rs("indsingle").value = True, "1", "0")
            Else
                sKeySingle = vbNullString
            End If
            
            If rs("ParteFissa") = False Then
                Set oNode = tvw.Nodes.Add(oNodeV.Key, tvwChild, oNodeV.Key & "Liv3" & SEPARATORE & rs("Riduzione") & sKeySingle, rs("Riduzione") & "%" & SEPARATOREVALORE & rs("valore") & SEPARATOREDESCRIZIONE & rs("DesBonus"))
            Else
                Set oNode = tvw.Nodes.Add(oNodeF.Key, tvwChild, oNodeF.Key & "Liv3" & SEPARATORE & rs("Riduzione") & sKeySingle, rs("Riduzione") & "%" & SEPARATOREVALORE & rs("valore") & SEPARATOREDESCRIZIONE & rs("DesBonus"))
            End If
            pKey = pKeySave
            For i = 0 To rs.fields.Count - 1
                If InStr(pKeySave, SEPARATORE & rs.fields(i).Name) Then
                    MettivaloreInPKEY pKey, SEPARATORE & rs.fields(i).Name, rs.fields(i).value, rs.fields(i).Type
                End If
            Next
            oNode.Tag = pKey
            oNodeP.ForeColor = vbBlue
        End If
        
        oNode.Expanded = False
        
        
        
        
        rs.MoveNext
    Loop

    Set qry = Nothing
    Set rs = Nothing


End Sub





Sub SalvaSuDB()
    Dim i As Long
    Dim sParte As Boolean
    Dim sKey As String
    Dim rs As Recordset
    Dim sTable As String
    Dim s As String
    
    For i = 1 To tvw.Nodes.Count
        Set oNode = tvw.Nodes(i)
        Select Case CheLivello(oNode.FullPath)
            Case 5
                
                If oNode.Parent.Parent.Parent.Key = "UA" Then
                    sTable = "BonusUtenzeAttivita"
                Else
                    sTable = "BonusUtenzeDomestiche"
                End If
                
                If oNode.Parent.Text = PF Then
                    sParte = True
                Else
                    sParte = False
                End If
                
                If oNode.Tag <> vbNullString Then
                    sKey = oNode.Tag
                    'cancello il record per permettere anche la variazione della percentuale
                    gDBC.Execute ("Delete from " & sTable & "  where " & GetPkeyForWhereCondition(sKey))
                End If
                
                sKey = oNode.Parent.Parent.Tag
                MettivaloreInPKEY sKey, SEPARATORE & "ParteFissa", sParte, adBoolean
                MettivaloreInPKEY sKey, SEPARATORE & "Riduzione", PrendiPercentuale(oNode.Text), adInteger
                MettivaloreInPKEY sKey, SEPARATORE & "indSingle", IIf(InStr(1, oNode.Text, RIDUZIONE_SINGLE) > 0, True, False), adBoolean
                
                If oNode.ForeColor <> vbRed Then
                    gDBC.Execute "Insert into " & sTable & "(" & getCampiDapKeyPerInsert(sKey) & ",valore,descrizione) values(" & getValoriDapKeyPerInsert(sKey) & "," & Replace(PrendiValore(oNode.Text), ",", ".") & ",'" & PrendiDescrizione(oNode.Text) & "')"
                End If
                
        End Select
    Next
    
    Set rs = Nothing
End Sub


Sub GestioneNodo(TipoParte As tpoParte, ByVal Modifica As Boolean, Optional testo As String = vbNullString)
    Dim sMessage As String
    Dim oNodeP As Node
    Dim oNodeIns As Node
    Dim parte As String
    Dim nodeSingle As Node
    Dim nodeChildSingle As Node
    Dim sKeySingle As String
    
    
    'controllo valori
    If Val(txt(0).Text) = 0 Or Val(txt(1).Text) = 0 Then
        MsgBox "I valori inseriti non sono validi", vbExclamation, App.Title
        Exit Sub
    End If
    
    
    If testo <> vbNullString Then
        txt(2).Text = testo
    Else
        'If Not Modifica Then txt(2).Text = ""
    End If
    
    
    If Not oNode Is Nothing Then
        On Error Resume Next
        Select Case Livello
            Case 3
                'livello utennza
                Set oNodeIns = oNode
            Case 4
                'livello PV o PF
                Set oNodeIns = oNode.Parent
            Case 5
                Set oNodeIns = oNode.Parent.Parent
        End Select
        
        If TipoParte = ParteFissa Then
            sMessage = "Attenzione il bonus:" & txt(1) & " è già presente nella parte fissa" + vbCrLf + "per la categoria (" & oNodeIns.Text & ")"
            parte = PF
        Else
            sMessage = "Attenzione il bonus:" & txt(1) & " è già presente nella parte variabile" + vbCrLf + "per la categoria (" & oNodeIns.Text & ")"
            parte = PV
        End If
        
        
        
        If InStr(1, oNodeIns.Key, UD) > 0 Then
            sKeySingle = SEPARATORE & IIf(txt(2).Text = RIDUZIONE_SINGLE, "1", "0")
        End If
        
        Set oNodeP = tvw.Nodes(oNodeIns.Key & "Liv2" & SEPARATORE & parte & "Liv3" & SEPARATORE & txt(1) & sKeySingle)
        If Err = 0 Then
            'se ho trovato il nodo e non sono in modifca o se sono in modifica e ho modifcato la percentuale non
            'deve esistere già
            If Not Modifica Or (oNodeP.Key <> oNode.Key) Then MsgBox sMessage, vbExclamation, App.Title: Exit Sub
        End If
        
        'un nodo riduzione single deve esistere una sola volta
        If txt(2).Text = RIDUZIONE_SINGLE Then
            Set nodeSingle = tvw.Nodes(oNodeIns.Key & "Liv2" & SEPARATORE & parte)
            If nodeSingle.children > 0 And Not Modifica Then
                Set nodeChildSingle = nodeSingle.Child
                
                Do While Not nodeChildSingle Is Nothing
                    If InStr(1, nodeChildSingle.Text, RIDUZIONE_SINGLE) <> 0 Then
                        MsgBox "Attenzione esiste già la riduzione per SINGLE", vbExclamation, App.Title: Exit Sub
                    End If
                    Set nodeChildSingle = nodeChildSingle.Next
                Loop
            End If
        End If
        
        If Modifica Then
            oNode.Text = txt(1) & "%" & SEPARATOREVALORE & txt(0) & SEPARATOREDESCRIZIONE & txt(2)
            oNode.Key = oNodeIns.Key & "Liv2" & SEPARATORE & parte & "Liv3" & SEPARATORE & txt(1) & sKeySingle
        Else
            Err.Clear
            Set oNodeP = tvw.Nodes(oNodeIns.Key & "Liv2" & SEPARATORE & parte)
            If Err > 0 Then
                oNodeIns.ForeColor = vbBlue
                Set oNodeP = tvw.Nodes.Add(oNodeIns.Key, tvwChild, oNodeIns.Key & "Liv2" & SEPARATORE & parte, parte, IIf(parte = PV, ICON_PV, ICON_PF))
            End If
            
            Set oNodeP = tvw.Nodes.Add(oNodeP.Key, tvwChild, oNodeP.Key & "Liv3" & SEPARATORE & txt(1) & sKeySingle, txt(1) & "%" & SEPARATOREVALORE & txt(0) & SEPARATOREDESCRIZIONE & txt(2))
        End If
        oNodeP.EnsureVisible
        
    End If


End Sub

Private Sub txt_LostFocus(index As Integer)
    If index = 1 Then txt(1).Text = CInt(txt(1).Text)
End Sub

Private Sub txt_Validate(index As Integer, Cancel As Boolean)
    If index = 0 Or index = 1 Then
        If Not IsNumeric(txt(index).Text) Then MsgBox "Il valore deve essere numerico", vbExclamation, App.Title: Cancel = True
    End If
    
End Sub
