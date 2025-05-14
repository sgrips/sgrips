VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFormulari 
   Caption         =   "Formulario di Uscita Rifiuto"
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
         Format          =   21102593
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
         Format          =   21102593
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
   Begin MSComctlLib.ListView lvwFormulario 
      Height          =   5055
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8916
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
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
         Text            =   "Nr Registrazione"
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
         Text            =   "Unita Misura"
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
Attribute VB_Name = "frmFormulari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private offsetX As Long
Private offsetY As Long

Private pwhere As String


Private Sub caricaFormulario()
    Dim itm As ListItem
    Dim rs As Recordset
    
    lvwFormulario.ListItems.Clear
    
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockBatchOptimistic
    
    pwhere = " Data between #" & Format$(DTPickerDal(0).Value, "yyyy-MM-dd") & "# and #" & Format$(Format$(DTPickerDal(1).Value, "yyyy-MM-dd") & " 23:59:59", "yyyy-MM-dd hh:mm:ss") & "#"
    If Not (Check1(0).Value = 1 And Check1(1).Value = 1) Then
        If (Check1(0).Value = 1) Then
            If pwhere <> "" Then pwhere = pwhere & " and "
            pwhere = pwhere & " flgStampato=true "
        ElseIf (Check1(1).Value = 1) Then
            If pwhere <> "" Then pwhere = pwhere & " and "
            pwhere = pwhere & " flgStampato=false "
        End If
    End If
    
    rs.Source = Replace(DevEnvMDB.rscmdFormulario.Source, "1=1", pwhere) & " order by Data"
    
    Set rs.ActiveConnection = gDBC
    rs.Open
    Set rs.ActiveConnection = Nothing
    
    With rs
        Do While Not .EOF
            Set itm = lvwFormulario.ListItems.Add(, "|" & .Fields("id"), .Fields("id"))
            itm.ListSubItems.Add , , Format$(.Fields("Data"), "dd/mm/yyyy")
            itm.ListSubItems.Add , , .Fields("CodRifiuto")
            itm.ListSubItems.Add , , .Fields("DesRifiuto")
            itm.ListSubItems.Add , , "KG"
            itm.ListSubItems.Add , , .Fields("quantita")
            itm.ListSubItems.Add , , .Fields("DesSmaRec")
            itm.ListSubItems.Add , , IIf(.Fields("flgStampato"), "Stampato", "Ancora da Stampare")
            itm.Checked = True
            .MoveNext
        Loop
    End With
    
        
    AllargaLeColonneLVW lvwFormulario, False
End Sub

Private Sub Command1_Click(Index As Integer)
    
    
    Select Case Index
        Case 0
            Unload Me
        Case 1
            stampaFormulario
        Case 2
            caricaFormulario
    End Select
End Sub


Private Sub stampaFormulario()
    Dim idFormulari As String
    Dim i As Long
    Dim iCount As Long
    Dim mWhere As String
    
    On Error GoTo stampaFormularioERR
    
    mWhere = pwhere
    iCount = lvwFormulario.ListItems.Count
    idFormulari = " id in ( "
    For i = 1 To lvwFormulario.ListItems.Count
    
        If lvwFormulario.ListItems(i).Checked Then
            idFormulari = idFormulari & "'" & lvwFormulario.ListItems(i).Text & "',"
            
            iCount = iCount - 1
        End If
    Next
    
    If iCount > 0 Then
        idFormulari = left$(idFormulari, Len(idFormulari) - 1) & " ) "
        mWhere = pwhere & " and " & idFormulari
    End If
    
    
    If gReportFormulari = "" Then
        If ReportFormulario.Stampa(mWhere, CBool(gStampaFormularioAnteprima), pfrmFormulari) Then
            gDBC.Execute "update FormularioRifiuti set flgStampato =true where " & mWhere
            caricaFormulario
        End If
    ElseIf gReportFormulari = "ReportFormularioSETA" Then
        If ReportFormularioSETA.Stampa(mWhere, CBool(gStampaFormularioAnteprima), pfrmFormulari) Then
            gDBC.Execute "update FormularioRifiuti set flgStampato =true where " & mWhere
            caricaFormulario
        End If
    End If
    
    Exit Sub
stampaFormularioERR:

End Sub



Private Sub Form_Load()
    Check1(1).Value = 1
    
    offsetX = 10 * Screen.TwipsPerPixelX
    offsetY = 10 * Screen.TwipsPerPixelY
    
    
    DTPickerDal(0).Value = primoNonStampato
    DTPickerDal(1).Value = Now
    
    applicaResizioneUtente
    
    caricaFormulario
End Sub

Private Function primoNonStampato() As Date
    Dim rs As ADODB.Recordset
    
    Set rs = gDBC.Execute("SELECT  data From FormularioRifiuti where flgStampato = False ORDER BY data")
    
    primoNonStampato = Now
    
    If Not rs.EOF Then primoNonStampato = rs(0)
End Function

Private Sub Form_Resize()
    Dim i As Integer
    Dim posLeft As Long
    Dim posTop As Long
    
    
    On Error Resume Next
    posTop = Me.ScaleHeight - Command1(i).Height - 50
    posLeft = Me.ScaleWidth - Command1(i).width - 50
    
    For i = Command1.LBound To Command1.ubound
        If i > 0 Then posLeft = posLeft - offsetX - Command1(i).width
        
        Command1(i).Move posLeft, posTop
    Next
    
    Picture1.top = Me.ScaleHeight - Picture1.ScaleHeight
    
    lvwFormulario.Move 50, 50, Me.ScaleWidth - 50, Picture1.top - offsetY
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set pfrmFormulari = Nothing
    
    If gReportFormulari = "" Then
        Unload ReportFormulario
    ElseIf gReportFormulari = "ReportFormularioSETA" Then
        Unload ReportFormularioSETA
    End If

End Sub


Private Sub applicaResizioneUtente()
    Command1(1).enabled = frmLogin.Livello <> LVL_USER
End Sub

Public Sub RefreshME()
    caricaFormulario
End Sub

