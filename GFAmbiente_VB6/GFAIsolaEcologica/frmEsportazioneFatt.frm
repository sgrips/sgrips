VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEsportazioneFatt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Esportazione"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   13335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Ricarica"
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
      Index           =   2
      Left            =   10080
      TabIndex        =   10
      Top             =   6840
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Già Esportati"
      Height          =   195
      Left            =   4560
      TabIndex        =   9
      Top             =   360
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvwMovEsp 
      Height          =   5895
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   10398
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Esporta"
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
      Index           =   1
      Left            =   11160
      TabIndex        =   7
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Chiudi"
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
      Left            =   12240
      TabIndex        =   6
      Top             =   6840
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   6960
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      Format          =   17694721
      CurrentDate     =   39357
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      Format          =   17694721
      CurrentDate     =   39357
   End
   Begin VB.Label lblEsportazione 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Avanzamento Esportazione"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   6720
      Width           =   9780
   End
   Begin VB.Label lblEsportazione 
      AutoSize        =   -1  'True
      Caption         =   "Al"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      Top             =   315
      Width           =   150
   End
   Begin VB.Label lblEsportazione 
      AutoSize        =   -1  'True
      Caption         =   "Dal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   315
      Width           =   285
   End
End
Attribute VB_Name = "frmEsportazioneFatt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mChecked As Long
Private mWhere As String

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
        
            Command1(0).enabled = False
            Command1(1).enabled = False
            Command1(2).enabled = False
                
            Esportazione
                
            Command1(0).enabled = True
            Command1(1).enabled = True
            Command1(2).enabled = True
        
        Case 2
            CaricaMovimenti
    End Select
End Sub

Private Sub Form_Load()
    lvwMovEsp.ColumnHeaders.Clear
    
    
    DTPicker1(0).value = "01-01-" & Year(Date)
    DTPicker1(1).value = Date
    
    
    DTPicker1(0).value = Null
    DTPicker1(1).value = Null
    
    CaricaMovimenti
End Sub


Private Sub Esportazione()
    Dim item As ListItem
    Dim pathEXP As String
    Dim fileEXP As String
    Dim fso As Scripting.FileSystemObject
    Dim txtStream As TextStream
    Dim i As Long
    Dim s As String
    
    On Error GoTo EsportazioneERR
    
    pathEXP = App.Path
    fileEXP = App.Path & "\" & gFattureFile & IIf(gFattureFileType = 0, ".TXT", ".CSV")
    
    If mChecked = 0 Then Exit Sub
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(fileEXP)) Then
        fso.DeleteFile fileEXP, True
    End If
    
    Set txtStream = fso.OpenTextFile(fileEXP, ForAppending, True)
     
    
    ProgressBar1.Min = 0
    ProgressBar1.Max = mChecked
    ProgressBar1.value = 0
    
    
    gDBC.BeginTrans
    For Each item In lvwMovEsp.ListItems
        If (item.Checked) Then
            
            ProgressBar1.value = ProgressBar1.value + 1
            s = GetCurrentComune!codistat & gFattureSeparatore
            s = s & Trim$(item.ListSubItems("PROGRESSIVO_INVIO").Text) & gFattureSeparatore
            s = s & Trim$(item.ListSubItems("COD_ANA").Text) & gFattureSeparatore
            s = s & Trim$(item.ListSubItems("COD_IMMOBILE").Text) & gFattureSeparatore
            s = s & Trim$(item.ListSubItems("COD_RIFIUTO").Text) & gFattureSeparatore
            s = s & Trim$(item.ListSubItems("QUANTITA").Text) & gFattureSeparatore
            s = s & Trim$(Format$(item.ListSubItems("DATA_MOVIMENTO").Text, "dd/MM/yyyy")) & gFattureSeparatore
            s = s & Trim$(item.ListSubItems("NUMERO_FORMULARIO").Text) & gFattureSeparatore
            s = s & Trim$(item.ListSubItems("STATO_INVIO").Text) & gFattureSeparatore
            s = s & Trim$(Format$(CStr(Now), "dd/MM/yyyy")) & gFattureSeparatore
            s = s & Trim$(Format$(CStr(Now), "hh\:nn\:ss")) & gFattureSeparatore
            s = s & Trim$(item.ListSubItems("FLAG").Text) & gFattureSeparatore
            s = s & Trim$(item.ListSubItems("NOTE").Text)
            txtStream.WriteLine s
                
            'modifico lo stato dei Fatturazione
            gDBC.Execute "update Fatturazione set stato='1',DATA_SCARICO=now() where id=" & Mid$(item.key, 2)
                
        End If
    Next
    
    
    If Not fso.FolderExists(gPercorsoEsp) Then
        fso.CreateFolder gPercorsoEsp
    End If
    
    txtStream.Close
    Set txtStream = Nothing
    Dim nomefile As String
    nomefile = gPercorsoEsp & "\" & gFattureFile & Format$(Now, "YYMMDDHHNNSS") & IIf(gFattureFileType = 0, ".TXT", ".CSV")
    fso.MoveFile fileEXP, nomefile
    
    If gFTP_Active Then
        'invio dati via FTP
        Dim FTP As frmFTP
        Set frmFTP = New frmFTP
     
       
        frmFTP.FileDestination = "out/" & fso.GetFileName(nomefile)
        frmFTP.FileSource = nomefile
        frmFTP.Rename = True
        frmFTP.DoWork 1
        Set frmFTP = Nothing
    End If
    
    Set fso = Nothing
    
    gDBC.CommitTrans
    
    CaricaMovimenti
    
    MsgBox "Esportazione avvenuta con successo", vbExclamation + vbOK, App.Title
    
    
    Exit Sub
EsportazioneERR:
    On Error Resume Next
    gDBC.RollbackTrans
End Sub

Private Sub CaricaMovimenti()
    Dim rs As ADODB.Recordset
    Dim item As ListItem
    Dim i As Integer
    
    
    Screen.MousePointer = vbHourglass
    mWhere = ""
    If (Not IsNull(DTPicker1(0).value)) Then
        If mWhere <> "" Then mWhere = mWhere & " AND "
        mWhere = "DATA_MOVIMENTO >= #" & Format$(DTPicker1(0).value, "YYYY-MM-DD 00:00:00") & "#"
    End If

    If (Not IsNull(DTPicker1(1).value)) Then
        If mWhere <> "" Then mWhere = mWhere & " AND "
        mWhere = mWhere & "DATA_MOVIMENTO <= #" & Format$(DTPicker1(1).value, "YYYY-MM-DD 23:59:59") & "#"
    End If

    If (Check1.value = 1) Then
        If mWhere <> "" Then mWhere = mWhere & " AND "
        mWhere = mWhere & "STATO = '1'"
    Else
        If mWhere <> "" Then mWhere = mWhere & " AND "
        mWhere = mWhere & "STATO = '0'"
    End If


    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    
        
    
    rs.Source = "SELECT  id,PROGRESSIVO_INVIO, Fatturazione.COD_ANA, Fatturazione.COD_IMMOBILE, COD_RIFIUTO, QUANTITA, DATA_MOVIMENTO, Fatturazione.NUMERO_FORMULARIO, Fatturazione.STATO_INVIO, DATA_SCARICO, FLAG, NOTE, STATO from Fatturazione " & _
            "WHERE " & mWhere & " and Fatturazione.codComune='" & GetCurrentComune!codistat & "'"
    
    Set rs.ActiveConnection = gDBC
    rs.Open
    Set rs.ActiveConnection = Nothing
    
    If rs.Fields.Count > 0 And lvwMovEsp.ColumnHeaders.Count = 0 Then
        For i = 0 To rs.Fields.Count - 1
            lvwMovEsp.ColumnHeaders.Add , rs.Fields(i).Name, rs.Fields(i).Name
        Next
    End If
    
    mChecked = 0
    lvwMovEsp.ListItems.Clear
    Do While Not rs.EOF
        
        For i = 0 To rs.Fields.Count - 1
            If i = 0 Then
                Set item = lvwMovEsp.ListItems.Add(, "_" & rs.Fields("id").value, rs.Fields(i).value)
            Else
                item.ListSubItems.Add , rs.Fields(i).Name, nullToBlank(rs.Fields(i).value)
            End If
        Next
        
        item.Checked = True
        rs.MoveNext
        
        mChecked = mChecked + 1
    Loop
    
    AllargaLeColonneLVW lvwMovEsp, False
    
    Screen.MousePointer = vbDefault
    If rs.state = adStateOpen Then
        rs.Close
        Set rs = Nothing
    End If
End Sub

Private Sub lvwMovEsp_ItemCheck(ByVal item As MSComctlLib.ListItem)
    If item.Checked = True Then
        mChecked = mChecked + 1
    Else
        mChecked = mChecked - 1
    End If
End Sub
