VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRegistri 
   Caption         =   "Registro di Carico e Scarico"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   8070
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "&Ricerca"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   4080
      TabIndex        =   8
      Top             =   5880
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   5295
      TabIndex        =   3
      Top             =   5520
      Width           =   5295
      Begin VB.CheckBox Check1 
         Caption         =   "Da Stampare"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   10
         Top             =   480
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Stampato"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   9
         Top             =   120
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPickerDal 
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   90439681
         CurrentDate     =   39453
      End
      Begin MSComCtl2.DTPicker DTPickerDal 
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   90439681
         CurrentDate     =   39453
      End
      Begin VB.Label Label1 
         Caption         =   "Al"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   7
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Dal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Stampa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   5400
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwRegistro 
      Height          =   5055
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8916
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nr Reg."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tipo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Registro Del"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "CodCer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Rifiuto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "U.M."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Quantità"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Destinazione"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Stato"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Chiudi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   6840
      TabIndex        =   0
      Top             =   5880
      Width           =   1215
   End
End
Attribute VB_Name = "frmRegistri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private offsetX As Long
Private offsetY As Long

Private pwhere As String


Private Sub caricaRegistro()
    Dim itm As ListItem
    Dim rs As Recordset
    
    lvwRegistro.ListItems.Clear
    
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockBatchOptimistic
    
    pwhere = " Registro.DataOperazione between #" & Format$(DTPickerDal(0).Value, "yyyy-MM-dd") & "# and #" & Format$(DTPickerDal(1).Value, "yyyy-MM-dd") & "#"
    If Not (Check1(0).Value = 1 And Check1(1).Value = 1) Then
        If (Check1(0).Value = 1) Then
            If pwhere <> "" Then pwhere = pwhere & " and "
            pwhere = pwhere & " Registro.flgStampato=true "
        ElseIf (Check1(1).Value = 1) Then
            If pwhere <> "" Then pwhere = pwhere & " and "
            pwhere = pwhere & " Registro.flgStampato=false "
        End If
    End If
    
    rs.Source = DevEnvMDB.rscmdRegistri.Source & " Where " & pwhere & " order by idregistro"
    
    Set rs.ActiveConnection = gDBC
    rs.Open
    Set rs.ActiveConnection = Nothing
    
    With rs
        Do While Not .EOF
            Set itm = lvwRegistro.ListItems.Add(, "|" & .Fields("idREgistro"), .Fields("idREgistro"))
            itm.ListSubItems.Add , , IIf(.Fields("carico") = "X", "Carico", "Scarico")
            itm.ListSubItems.Add , , Format$(.Fields("DataOperazione"), "dd/mm/yyyy")
            itm.ListSubItems.Add , , .Fields("CodRifiuto")
            itm.ListSubItems.Add , , .Fields("DesRifiuto")
            itm.ListSubItems.Add , , .Fields("UnitaMisura")
            itm.ListSubItems.Add , , .Fields("quantita")
            itm.ListSubItems.Add , , IIf(.Fields("Recupe") = "X", nullToBlank(.Fields("RecupeDes")), nullToBlank(.Fields("SmaltDes")))
            itm.ListSubItems.Add , , IIf(.Fields("flgStampato"), "Stampato", "Ancora da Stampare")
            itm.Checked = True
            .MoveNext
        Loop
    End With
    
        
    AllargaLeColonneLVW lvwRegistro, False
End Sub

Private Sub Command1_Click(Index As Integer)
    
    
    Select Case Index
        Case 0
            Unload Me
        Case 1
            stampaRegistro
        Case 2
            caricaRegistro
    End Select
End Sub


Private Sub stampaRegistro()
    Dim rs As ADODB.Recordset
    Dim qry As ADODB.Command
    Dim oStream As Stream
    Dim idRegistro As Long
    Dim bBeginTrans As Boolean
    
    On Error GoTo stampaRegistroERR
    
    gDBC.BeginTrans
    bBeginTrans = True
    gDBC.Execute ("delete from tmpStampaRiferimentiRegistro")
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.LockType = adLockBatchOptimistic
    rs.CursorType = adOpenStatic
    
    rs.Source = Replace("SELECT Registro.idRegistro, RegistroCar.nrRegistro" & _
                " FROM ((Registro INNER JOIN LegameRegistroOperazioneCarico as lop ON Registro.idRegistro = lop.idRegistro )" & _
                " INNER JOIN Registro as RegistroCar ON RegistroCar.idRegistro = lop.idOperazioneCarico )" & _
                " WHERE Registro.flgOperazione='S' AND 1=1 order by Registro.idRegistro", "1=1", pwhere)
    Set rs.ActiveConnection = gDBC
    rs.Open
    
    Set qry = CreateObject("ADODB.Command")
    qry.CommandText = "Insert into tmpStampaRiferimentiRegistro (idRegistro,riferimenti) values (?,?)"
    Set qry.ActiveConnection = gDBC
    
    qry.Parameters.Append qry.CreateParameter(, adInteger, adParamInput)
    qry.Parameters.Append qry.CreateParameter(, adLongVarWChar, adParamInput, 30000)
    
    
    idRegistro = 0
    Do While Not rs.EOF
        If idRegistro <> rs!idRegistro.Value Then
            If Not oStream Is Nothing Then
                If oStream.size <> 0 Then
                    oStream.Position = 0
                    qry.Parameters(0).Value = idRegistro
                    qry.Parameters(1).Value = oStream.ReadText
                    qry.Execute
                    ' chiudo lo stream
                    oStream.Close
                    Set oStream = Nothing
                
                End If
            End If
            'ADODB.Stream
            Set oStream = CreateObject("ADODB.Stream")
            ' apro lo stream
            oStream.Open
            ' charset
            oStream.Charset = "ASCII"
            ' tipo di scrittura: testo
            oStream.Type = adTypeText
            idRegistro = rs!idRegistro.Value
        End If
        
        oStream.WriteText CStr(rs!nrRegistro.Value & ", "), adWriteChar
        rs.MoveNext
    Loop
    
    
    oStream.Position = 0
    qry.Parameters(0).Value = idRegistro
    qry.Parameters(1).AppendChunk oStream.ReadText
    qry.Execute
    ' chiudo lo stream
    oStream.Close
    Set oStream = Nothing
    
    rs.Close
    Set rs = Nothing
    
    
    gDBC.CommitTrans
    bBeginTrans = False
    
    If ReportRegistro.Stampa(pwhere, CBool(gStampaRegistroAnteprima), pfrmRegistri) Then
        gDBC.Execute "update Registro set flgStampato =true where " & pwhere
        caricaRegistro
    End If
    
    Exit Sub
stampaRegistroERR:
    If bBeginTrans Then gDBC.RollbackTrans
End Sub



Private Sub Form_Load()
    Check1(1).Value = 1
    
    offsetX = 10 * Screen.TwipsPerPixelX
    offsetY = 10 * Screen.TwipsPerPixelY
    
    
    DTPickerDal(0).Value = primoNonStampato
    DTPickerDal(1).Value = Now
    
    applicaResizioneUtente
    
    caricaRegistro
End Sub

Private Function primoNonStampato() As Date
    Dim rs As ADODB.Recordset
    
    Set rs = gDBC.Execute("SELECT Registro.dataOperazione From Registro where (((Registro.flgStampato) = False)) ORDER BY Registro.dataOperazione")
    
    primoNonStampato = Now
    
    If Not rs.EOF Then primoNonStampato = rs(0)
End Function

Private Sub Form_Resize()
    Dim i As Integer
    Dim posLeft As Long
    Dim posTop As Long
    
    
    On Error Resume Next
    posTop = Me.ScaleHeight - Command1(i).Height - 50
    posLeft = Me.ScaleWidth - Command1(i).Width - 50
    
    For i = Command1.LBound To Command1.UBound
        If i > 0 Then posLeft = posLeft - offsetX - Command1(i).Width
        
        Command1(i).Move posLeft, posTop
    Next
    
    Picture1.Top = Me.ScaleHeight - Picture1.ScaleHeight
    
    lvwRegistro.Move 50, 50, Me.ScaleWidth - 50, Picture1.Top - offsetY
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set pfrmRegistri = Nothing
    
End Sub


Private Sub applicaResizioneUtente()
    Command1(1).enabled = frmLogin.Livello <> LVL_USER
End Sub

Public Sub RefreshME()
    caricaRegistro
End Sub

