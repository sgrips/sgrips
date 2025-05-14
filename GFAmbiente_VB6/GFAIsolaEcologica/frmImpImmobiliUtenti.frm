VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImpImmobiliUtenti 
   Caption         =   "Importazione Immobili - Utenti"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4860
   ScaleWidth      =   8070
   Begin VB.CommandButton cmd 
      Caption         =   "&Scarica Da FTP"
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
      Index           =   4
      Left            =   1560
      TabIndex        =   13
      Top             =   4320
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   6255
      TabIndex        =   9
      Top             =   3720
      Width           =   6255
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   0
         X1              =   720
         X2              =   5400
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   720
         X2              =   5400
         Y1              =   150
         Y2              =   150
      End
      Begin VB.Label Label1 
         Caption         =   "Badge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Importa"
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
      Index           =   3
      Left            =   3240
      TabIndex        =   6
      Top             =   4320
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwAnteprima 
      Height          =   1695
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2990
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Carica"
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
      Left            =   4440
      TabIndex        =   2
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
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
      Left            =   6840
      TabIndex        =   1
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Anteprima"
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
      Left            =   5640
      TabIndex        =   0
      Top             =   4320
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwImportazione 
      Height          =   1695
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2990
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   4200
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
   End
   Begin MSComctlLib.ListView lvwAnteprima 
      Height          =   1695
      Index           =   1
      Left            =   1920
      TabIndex        =   7
      Top             =   1680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2990
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwImportazione 
      Height          =   1695
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   1680
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2990
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwAnteprima 
      Height          =   1695
      Index           =   2
      Left            =   6000
      TabIndex        =   11
      Top             =   1680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2990
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwImportazione 
      Height          =   1695
      Index           =   2
      Left            =   4080
      TabIndex        =   12
      Top             =   1680
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2990
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmImpImmobiliUtenti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CMD_CHIUDI = 0
Private Const CMD_ANTEPRIMA = 1
Private Const CMD_CARICA = 2
Private Const CMD_IMPORTA = 3
Private Const CMD_FTP = 4

Private offsetX As Long
Private offsetY As Long
Private psFileName() As String
Private legameBarcodeBadge As Boolean
Private sCampoNonTrovato As String

Private Sub commandEnabled(enable As Boolean)
    Dim i As Integer
    
    For i = cmd.LBound To cmd.UBound
        cmd(i).enabled = enable
    Next
End Sub


Private Sub cmd_Click(Index As Integer)
    Dim enabled As Boolean
    
    Select Case Index
        Case CMD_CHIUDI
            MDIMain.tmrInvioAccessi.enabled = True
            Unload Me
        Case CMD_ANTEPRIMA
            commandEnabled False
            
            lvwAnteprima(0).Visible = True
            lvwAnteprima(1).Visible = True
            lvwAnteprima(2).Visible = legameBarcodeBadge
            
            enabled = Anteprima() And AnteprimaBadge() And AnteprimaLegameBarcodeBadge()
            
            If enabled = False And Me.Visible Then
            
                MsgBox "Attenzione verificare il nome delle intestazione di colonna [" + sCampoNonTrovato + "] nel file CSV e che il formato sia il seguente:" + vbCrLf + "Separatore Campo [" + pImpImmUteSepCol + "]" _
                    + vbCrLf _
                    + vbCrLf + "Separatore Decimale [" + pImpImmUteSepDec + "]" _
                    + vbCrLf + "Separatore Testo [" + pImpImmUteSepText + "]" _
                    + vbCrLf + "Percorso [" + pImpImmUtePath + "]"
            End If
            
            commandEnabled True
            
            cmd(CMD_CARICA).enabled = enabled
            cmd(CMD_IMPORTA).enabled = False
            
        Case CMD_CARICA
            commandEnabled False
            
            lvwAnteprima(0).Visible = True
            lvwAnteprima(1).Visible = True
            lvwAnteprima(2).Visible = legameBarcodeBadge
            
            enabled = False
            If Carica() Then
                If CaricaBadge() Then
                    If CaricaLegameBarcodeBadge() Then
                        enabled = True
                    End If
                End If
            End If
            
            'enabled = Carica() And CaricaBadge() And CaricaLegameBarcodeBadge()
            
            commandEnabled True
            'carica i file ed l'importa in automatico
            cmd(CMD_IMPORTA).enabled = enabled
            
            If enabled Then cmd_Click (CMD_IMPORTA)
            
        Case CMD_IMPORTA
            commandEnabled False
            Dim f As Object
            Dim i As Integer
            Dim operazioneOK As Boolean
            
            
            'cancella utenti, immobili, legami e badge
            gDBC.Execute ("delete from utenti")
            gDBC.Execute ("delete from immobili")
            gDBC.Execute ("delete from badge")
            
            operazioneOK = Importa
            If operazioneOK Then
                operazioneOK = ImportaBadge
                If operazioneOK And legameBarcodeBadge Then
                    operazioneOK = ImportaLegamiBadge
                End If
            End If
            
            If operazioneOK Then
                'cancello i file importati
                Set f = CreateObject("Scripting.FileSystemObject")
    
                For i = LBound(psFileName) To UBound(psFileName)
                    f.DeleteFile (pImpImmUtePath & "\" & psFileName(i))
                Next i
                
                f.DeleteFile (pImpImmUtePath & "\*.zip")
                MsgBox "Importazione terminta con successo!", vbOKOnly, "Importazione"
            End If
            
            commandEnabled True
        Case CMD_FTP
            ScaricaFileDaFTP
    End Select
    
End Sub

Private Sub ScaricaFileDaFTP()
    Dim FileName As String
    Dim frmFTP As frmFTP
    Dim so As FileSystemObject
    
    Set so = New FileSystemObject
    'FileName = GetCurrentComune!codistat & "*_legamebadge_storico.csv"
 
    Set frmFTP = New frmFTP
    FileName = "*.zip"
    frmFTP.FileDestination = gPercorsoImp & "\"
    frmFTP.FileSource = "in/" & FileName
    frmFTP.Rename = True
    frmFTP.DoWork 2
    If frmFTP.FileDestination <> "" Then
        If so.FileExists(frmFTP.FileDestination) Then
            
            On Error Resume Next
            DoEvents
            frmAttesa.Show
            DoEvents
            
            ExecuteAndWait (App.Path & "\unzip.exe -o -j " & Chr$(34) & frmFTP.FileDestination & Chr$(34) & " -d " & gPercorsoImp)
            
            psFileName = findFileToImport
            Unload frmAttesa
            
            cmd_Click (CMD_ANTEPRIMA)
            cmd_Click (CMD_CARICA)
        End If
    End If
    Set frmFTP = Nothing
     
End Sub

Private Sub Form_Load()
    commandEnabled False
    
    MDIMain.tmrInvioAccessi.enabled = False
    
    cmd(CMD_CHIUDI).enabled = True
    cmd(CMD_ANTEPRIMA).enabled = True
    
    pImpImmUtePath = gPercorsoImp
    pImpImmUteSepCol = gUtentiSeparatore
    pImpImmUteSepText = """"
    pImpImmUteSepDec = "."
    
    
    With lvwAnteprima(0)
        .View = lvwReport
        .HideColumnHeaders = False
        .GridLines = True
        .Visible = False
        .FullRowSelect = True
    End With
    
    With lvwAnteprima(1)
        .View = lvwReport
        .HideColumnHeaders = False
        .GridLines = True
        .Visible = False
        .FullRowSelect = True
    End With
    
    With lvwAnteprima(2)
        .View = lvwReport
        .HideColumnHeaders = False
        .GridLines = True
        .Visible = False
        .FullRowSelect = True
    End With
    
    psFileName = findFileToImport
    
    On Error Resume Next
    Me.Caption = Me.Caption & ":" & psFileName(0)
    
    cmd_Click (CMD_ANTEPRIMA)
    
End Sub

Private Sub Form_Resize()
    Dim posLeft As Long
    Dim posTop As Long
    Dim i As Integer
    
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub

    offsetX = 10 * Screen.TwipsPerPixelX
    offsetY = 10 * Screen.TwipsPerPixelX
    
    posTop = Me.ScaleHeight - cmd(i).Height - offsetX
    posLeft = Me.ScaleWidth - cmd(i).width - offsetX
    
    For i = cmd.LBound To cmd.UBound
        If i > 0 Then posLeft = posLeft - offsetX - cmd(i).width
        
        cmd(i).Move posLeft, posTop
    Next
    
    lvwAnteprima(0).Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, cmd(0).top / 2
    lvwImportazione(0).Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, cmd(0).top / 2
    
    Picture1.Move lvwAnteprima(0).left + offsetX, lvwAnteprima(0).top + lvwAnteprima(0).Height + offsetY, lvwAnteprima(0).width - offsetX
    
    posTop = Picture1.top + Picture1.Height + offsetY
    
    lvwAnteprima(1).Move Me.ScaleLeft, posTop, Me.ScaleWidth * IIf(legameBarcodeBadge, 0.5, 1) - offsetX, cmd(0).top - posTop - offsetY
    lvwImportazione(1).Move Me.ScaleLeft, posTop, Me.ScaleWidth * IIf(legameBarcodeBadge, 0.5, 1) - offsetX, cmd(0).top - posTop - offsetY
    
    lvwAnteprima(2).Move lvwAnteprima(1).left + lvwAnteprima(1).width + offsetX, lvwAnteprima(1).top, lvwAnteprima(1).width, lvwAnteprima(1).Height
    lvwImportazione(2).Move lvwAnteprima(1).left + lvwAnteprima(1).width + offsetX, lvwAnteprima(1).top, lvwAnteprima(1).width, lvwAnteprima(1).Height
    
    
    ProgressBar1.Move offsetX, cmd(0).top, cmd(cmd.UBound).left - offsetX, cmd(0).Height
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set pfrmImpImmobiliUtenti = Nothing
End Sub

Private Function findFileToImport() As String()
    Dim s() As String
    Dim sLastDate As Long
    Dim sDate As String
    Dim i As Integer
    Dim sFile As String
    Dim codComune As String
    Dim rs As ADODB.Recordset
    
    ReDim s(2)
    
    On Error GoTo findFileToImportERR
    
    Dim rsComune As ADODB.Recordset
    Set rsComune = GetCurrentComune
    If rsComune Is Nothing Then Exit Function

    codComune = rsComune!codistat
   
    sFile = Dir$(pImpImmUtePath & "\" & codComune & "*_immobili_storico.csv")
    
    Do While sFile <> vbNullString
        i = InStr(7, sFile, "_")
        sDate = Mid$(sFile, i + 1, 8)
        sDate = Right$(sDate, 4) & Mid$(sDate, 3, 2) & left$(sDate, 2)
        If sLastDate < sDate Then sLastDate = sDate
        sFile = Dir$
    Loop
    s(0) = codComune & "_" & Right$(sLastDate, 2) & Mid$(sLastDate, 5, 2) & left$(sLastDate, 4) & "_immobili_storico.csv"
        
    sFile = Dir$(pImpImmUtePath & "\" & codComune & "*_badge_storico.csv")
    Do While sFile <> vbNullString
        i = InStr(7, sFile, "_")
        sDate = Mid$(sFile, i + 1, 8)
        sDate = Right$(sDate, 4) & Mid$(sDate, 3, 2) & left$(sDate, 2)
        If sLastDate < sDate Then sLastDate = sDate
        sFile = Dir$
    Loop
    s(1) = codComune & "_" & Right$(sLastDate, 2) & Mid$(sLastDate, 5, 2) & left$(sLastDate, 4) & "_badge_storico.csv"
        
        
    sFile = Dir$(pImpImmUtePath & "\" & codComune & "*_legamebadge_storico.csv")
    Do While sFile <> vbNullString
        i = InStr(7, sFile, "_")
        sDate = Mid$(sFile, i + 1, 8)
        sDate = Right$(sDate, 4) & Mid$(sDate, 3, 2) & left$(sDate, 2)
        If sLastDate < sDate Then sLastDate = sDate
        sFile = Dir$
    Loop
    s(2) = codComune & "_" & Right$(sLastDate, 2) & Mid$(sLastDate, 5, 2) & left$(sLastDate, 4) & "_legamebadge_storico.csv"
        
        
    legameBarcodeBadge = Dir$(pImpImmUtePath & "\" & s(2)) <> vbNullString
        
    findFileToImport = s
    Exit Function
findFileToImportERR:
MsgBox (Err.Description)
End Function


Private Function Anteprima() As Boolean
    Dim f As Scripting.FileSystemObject
    Dim fText As TextStream
    Dim s As String
    Dim col() As String
    Dim i As Long
    Dim itm  As ListItem
    Dim countRecord As Integer
    Dim colHeader As ColumnHeader
    
    
On Error GoTo AnteprimaERR

    Set f = CreateObject("Scripting.FileSystemObject")

    Anteprima = True
    
    
    lvwAnteprima(0).ListItems.Clear
    lvwAnteprima(0).ColumnHeaders.Clear
    lvwAnteprima(0).Visible = False
    
   
    
    Set fText = f.OpenTextFile(pImpImmUtePath & "\" & psFileName(0), ForReading, False)
    'carico i primi 30 record per verificare il formato del file
    countRecord = 30
    Do While Not fText.AtEndOfStream And countRecord > 0
        s = fText.ReadLine
        col = scomponiRiga(s, pImpImmUteSepCol, pImpImmUteSepDec, pImpImmUteSepText)
        
        
        If UBound(col) > lvwAnteprima(0).ColumnHeaders.Count - 1 Then
           
           For i = 0 To UBound(col)
                Set colHeader = lvwAnteprima(0).ColumnHeaders.Add(, "col" & i, col(i))
                If Not colTracciati.isContained("IMMUTE", col(i)) Then
                    colHeader.Text = colHeader.Text & "*"
                    sCampoNonTrovato = col(i)
                    Anteprima = False
                End If
            
           Next
        Else
            For i = LBound(col) To UBound(col)
                With lvwAnteprima(0)
                    
                    If i = 0 Then
                        Set itm = .ListItems.Add(, , col(i))
                    Else
                        itm.ListSubItems.Add , , col(i)
                    End If
                End With
            Next
            countRecord = countRecord - 1
        End If
    Loop
    
    
    

    fText.Close
    Set fText = Nothing
    Set f = Nothing
    
    
    
    AllargaLeColonneLVW lvwAnteprima(0), False
    
    lvwAnteprima(0).Visible = True
    Exit Function
AnteprimaERR:
    Anteprima = False
    If Me.Visible Then MsgBox "frmImpImmobiliUtenti Anteprima:" & Err.Description
End Function

Private Function AnteprimaBadge() As Boolean
    Dim f As New Scripting.FileSystemObject
    Dim fText As TextStream
    Dim s As String
    Dim col() As String
    Dim i As Long
    Dim itm  As ListItem
    Dim countRecord As Integer
    Dim colHeader As ColumnHeader
    
    
On Error GoTo AnteprimaBadgeERR
    AnteprimaBadge = False
    
    lvwAnteprima(1).ListItems.Clear
    lvwAnteprima(1).ColumnHeaders.Clear
    lvwAnteprima(1).Visible = False
    
    
    AnteprimaBadge = True
    
    Set fText = f.OpenTextFile(pImpImmUtePath & "\" & psFileName(1), ForReading, False)
    'carico i primi 30 record per verificare il formato del file
    countRecord = 30
    Do While Not fText.AtEndOfStream And countRecord > 0
        s = fText.ReadLine
        col = scomponiRiga(s, pImpImmUteSepCol, pImpImmUteSepDec, pImpImmUteSepText)
        
        
        If UBound(col) > lvwAnteprima(1).ColumnHeaders.Count - 1 Then
           
           For i = 0 To UBound(col)
                Set colHeader = lvwAnteprima(1).ColumnHeaders.Add(, "col" & i, col(i))
                If Not colTracciati.isContained("UTEBADGE", col(i)) Then
                    colHeader.Text = colHeader.Text & "*"
                     sCampoNonTrovato = col(i)
                    AnteprimaBadge = False
                End If
            
           Next
        Else
            For i = LBound(col) To UBound(col)
                With lvwAnteprima(1)
                    
                    If i = 0 Then
                        Set itm = .ListItems.Add(, , col(i))
                    Else
                        itm.ListSubItems.Add , , col(i)
                    End If
                End With
            Next
            countRecord = countRecord - 1
        End If
    Loop
    
    

    fText.Close
    Set fText = Nothing
    Set f = Nothing
    
    
    
    AllargaLeColonneLVW lvwAnteprima(1), False
    
    lvwAnteprima(1).Visible = True
    Exit Function
AnteprimaBadgeERR:
    AnteprimaBadge = False
End Function


Private Function AnteprimaLegameBarcodeBadge() As Boolean
    Dim f As New Scripting.FileSystemObject
    Dim fText As TextStream
    Dim s As String
    Dim col() As String
    Dim i As Long
    Dim itm  As ListItem
    Dim countRecord As Integer
    Dim colHeader As ColumnHeader
    
    
On Error GoTo AnteprimaLegameBarcodeBadgeERR
    AnteprimaLegameBarcodeBadge = False
    
    lvwAnteprima(2).ListItems.Clear
    lvwAnteprima(2).ColumnHeaders.Clear
    lvwAnteprima(2).Visible = False
    
    
    AnteprimaLegameBarcodeBadge = True
    
    If Not legameBarcodeBadge Then Exit Function
    
    Set fText = f.OpenTextFile(pImpImmUtePath & "\" & psFileName(2), ForReading, False)
    'carico i primi 30 record per verificare il formato del file
    countRecord = 30
    Do While Not fText.AtEndOfStream And countRecord > 0
        s = fText.ReadLine
        col = scomponiRiga(s, pImpImmUteSepCol, pImpImmUteSepDec, pImpImmUteSepText)
        
        
        If UBound(col) > lvwAnteprima(2).ColumnHeaders.Count - 1 Then
           
           For i = 0 To UBound(col)
                Set colHeader = lvwAnteprima(2).ColumnHeaders.Add(, "col" & i, col(i))
                If Not colTracciati.isContained("LEGBADGE", col(i)) Then
                    colHeader.Text = colHeader.Text & "*"
                    AnteprimaLegameBarcodeBadge = False
                     sCampoNonTrovato = col(i)
                End If
            
           Next
        Else
            For i = LBound(col) To UBound(col)
                With lvwAnteprima(2)
                    
                    If i = 0 Then
                        Set itm = .ListItems.Add(, , col(i))
                    Else
                        itm.ListSubItems.Add , , col(i)
                    End If
                End With
            Next
            countRecord = countRecord - 1
        End If
    Loop
    
    

    fText.Close
    Set fText = Nothing
    Set f = Nothing
    
    
    
    AllargaLeColonneLVW lvwAnteprima(2), False
    
    lvwAnteprima(2).Visible = True
    Exit Function
AnteprimaLegameBarcodeBadgeERR:
    AnteprimaLegameBarcodeBadge = False
End Function

Private Function scomponiRiga(sValue As String, sepCol As String, sepDec As String, sepText As String, Optional Formattazione As String = vbNullString) As String()
    Dim s() As String
    Dim s1 As String
    Dim i As Long
    Dim rexp As New RegExp
    Dim matchCol As MatchCollection
    Dim newSize As Long
    
    If sepText <> "" Then
        rexp.Pattern = sepCol & "(" & sepText & ".*?" & sepText & ")|" & sepCol & "[^" & sepCol & "]*"
    Else
        rexp.Pattern = sepCol & "[^" & sepCol & "]*"
    End If
    rexp.Global = True
    rexp.IgnoreCase = True
    Set matchCol = rexp.Execute(sepCol & sValue)
    
    'se il separatore testo è impostato verifico se il dato è racchiuso tra esso
    'If UBound(s) > 0 And sepText <> "" Then
    '    For i = LBound(s) To UBound(s)
    '        s1 = Replace(s(i), sepText, vbNullString)
    '        If Formattazione <> "" Then s1 = Format$(s1, Formattazione)
    '        s(i) = s1
    '    Next
   ' End If
    Dim c As Match
    For Each c In matchCol
        If sepText <> "" Then
            s1 = Replace(c.value, sepText, vbNullString)
            If Formattazione <> "" Then s1 = Format$(s1, Formattazione)
            On Error Resume Next
            newSize = UBound(s) + 1
            ReDim Preserve s(newSize)
            On Error GoTo 0
            s(newSize) = Mid$(s1, 2)
        End If
    Next
    scomponiRiga = s
End Function

Private Function Carica() As Boolean
    Dim f As New Scripting.FileSystemObject
    Dim fText As TextStream
    Dim s As String
    Dim col() As String
    Dim i As Long
    
    Dim itm  As ListItem
    Dim cmd As ADODB.Command
    Dim byteTot As Long
    Dim fFile As file
    Dim colIdxField As New Collection
    Dim colTracciato As Collection
    Dim pTracciatoMap As TracciatoMapping
    
    Dim rs As ADODB.Recordset
    
On Error GoTo CaricaERR
    Carica = False
    
    
    lvwImportazione(0).ListItems.Clear
    lvwImportazione(0).ColumnHeaders.Clear
    lvwImportazione(0).Visible = False
    
    Set cmd = CreateObject("ADODB.Command")
    cmd.CommandText = "delete * from tmpImmobiliUtenti"
    Set cmd.ActiveConnection = gDBC
    cmd.Execute
    Set cmd.ActiveConnection = Nothing
    
    cmd.CommandText = "INSERT INTO tmpImmobiliUtenti (DataScarico,CodImmobile,Tipo,CodImmobilePrincipale,CodComune,ComuneDes,ComuneCap,ComuneBelfiore, " & _
        "CodVia,Indirizzo,NumCivico,Subalterno,NumInterno,Scala,Piano,Superficie,NucleoFamiliare,CodTariffa,ClasseTariffa, " & _
        "TipoUtenza,CodAnagrafe,TipoAnagrafe,Cognome,Nome,CodFiscale,DataDiNascita,CodComuneDiNascita,RagioneSociale, " & _
        "PartitaIva,CodComuneResidenza,ComuneResidenzaDes,ComuneResidenzaCap,ComuneResidenzaBelfiore,CodViaResidenza, " & _
        "IndirizzoResidenza,NumCivicoResidenza,SubalternoResidenza,InternoResidenza,ScalaResidenza,PianoResidenza,TipoDomicilio,NominativoRecapito, " & _
        "CodComuneRecapito,ComuneRecapitoDes,ComuneRecapitoCap,ComuneRecapitoBelfiore, " & _
        "CodViaRecapito,IndirizzoRecapito,NumCivicoRecapito,SubalternoRecapito,InternoRecapito,ScalaRecapito,PianoRecapito,FlgUtenzaDomesticaResidente, " & _
        "DataInizioValidita,DataFineValidita,FlgInquilino,EnteCreditore,Sesso,Telefono) " & _
        "values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?, ?, ?, ?, ?, ?,?, ?, ?, ?, ?, ?, ?, ?," & _
        "?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ? , ?, ?, ?, ?, ?,  ?, ?, ?, ?, ?, ?,?,?)"

    cmd.Prepared = True
    Set cmd.ActiveConnection = gDBC
    
    
    Set colTracciato = colTracciati.item("IMMUTE")
    For i = 1 To colTracciato("FIELD").Count
        Set pTracciatoMap = colTracciato("FIELD")(i)
        
        s = Split(pTracciatoMap.mappingField, pImpImmUteSepCol)(0)
        If pTracciatoMap.fieldType = FIELD_TYPE_NUM Then
            If pTracciatoMap.fieldSizeLen > 1 And pTracciatoMap.fieldSizePrecision > 0 Then
                Dim p1 As Parameter
                Set p1 = cmd.CreateParameter(s, adDecimal, adParamInput)
                p1.size = 8
                p1.Precision = 2
                cmd.Parameters.Append p1
            Else
                cmd.Parameters.Append cmd.CreateParameter(s, adBigInt, adParamInput)
            End If
            
        ElseIf s = "ClasseTariffa" Then
            cmd.Parameters.Append cmd.CreateParameter(s, adVarChar, adParamInput, 255)
        Else
            cmd.Parameters.Append cmd.CreateParameter(s, adVarChar, adParamInput, 100)
        End If
    Next
    
    cmd.Parameters.Append cmd.CreateParameter("EnteCreditore", adVarChar, adParamInput, 100)
    cmd.Parameters.Append cmd.CreateParameter("Sesso", adVarChar, adParamInput, 100)
    cmd.Parameters.Append cmd.CreateParameter("Telefono", adVarChar, adParamInput, 100)
    
    'lettura file di testo
    Set fText = f.OpenTextFile(pImpImmUtePath & "\" & psFileName(0), ForReading, False)
    
    Set fFile = f.GetFile(pImpImmUtePath & "\" & psFileName(0))
    byteTot = fFile.size
    
    ProgressBar1.Min = 0
    ProgressBar1.Max = byteTot
    ProgressBar1.value = 0
    Me.Caption = "Importazione Immobili - Utenti:" & psFileName(0)

    Set fFile = Nothing
    
    ProgressBar1.Tag = (ProgressBar1.Max / 100)
    
    Do While Not fText.AtEndOfStream
        s = fText.ReadLine
        col = scomponiRiga(s, pImpImmUteSepCol, pImpImmUteSepDec, pImpImmUteSepText)
        ProgressBar1.value = ProgressBar1.value + Len(s)
        Me.Caption = "Importazione Immobili - Utenti:" & psFileName(0) & " - " & FormatNumber(ProgressBar1.value * 100 / ProgressBar1.Max, 2)
        
        'importazione in tabella temporanea
        If colIdxField.Count = 0 Then
            'creo indice per recuperare i campi da importare successivamente
            For i = LBound(col) To UBound(col)
                Set pTracciatoMap = colTracciato("FIELDEXT")(col(i))
                s = Split(pTracciatoMap.mappingField, pImpImmUteSepCol)(0)
                colIdxField.Add i, s
            Next
        Else
        
            For i = 0 To cmd.Parameters.Count - 1
               If cmd.Parameters(i).Type = adChar Or cmd.Parameters(i).Type = adVarChar Or cmd.Parameters(i).Type = adWChar Then
                    cmd.Parameters(i).value = vbNullString
                Else
                    cmd.Parameters(i).value = 0
                End If
            Next
        
            For i = 0 To cmd.Parameters.Count - 1
                On Error Resume Next
                Set pTracciatoMap = colTracciato("FIELD")(cmd.Parameters(i).Name)
                s = Trim$(col(colIdxField(cmd.Parameters(i).Name)))
                On Error GoTo CaricaERR:
                'On Error GoTo 0
              '  Debug.Assert s <> "9792" Or cmd.Parameters(i).Name <> "CodImmobile"
                
                
                Select Case cmd.Parameters(i).Name
                    Case "CodComune", "CodComuneDiNascita", _
                        "CodComuneRecapito", "CodComuneResidenza"
                        s = IIf(Val(s) = 0, vbNullString, Format$(s, "000000"))
                    Case "FlgUtenzaDomesticaResidente"
                        s = IIf(s = "R", -1, 0)
                    Case "DataFineValidita"
                        s = IIf(s = "99/99/9999", "31/12/2050", s)
                End Select
                
                If pTracciatoMap.fieldFormat <> vbNullString And s <> vbNullString Then s = Format$(s, pTracciatoMap.fieldFormat)
                If pTracciatoMap.fieldType = FIELD_TYPE_CHAR Then
                     s = left$(s, pTracciatoMap.fieldSizeLen)
                End If
                
'                Debug.Assert cmd.Parameters(i).Name <> "DataFineValidita"
                
                'controllo tipo dato
                
                If cmd.Parameters(i).Type = adDecimal Then
                    If s <> vbNullString Then s = Replace(s, ".", ",")
                End If
                cmd.Parameters(i).value = IIf(s = vbNullString, Null, s)
                'Debug.Print cmd.Parameters(i).Name & "=" & nullToBlank(cmd.Parameters(i).Value)
                
            Next
            cmd.Parameters("Telefono").value = Null
            cmd.Parameters("EnteCreditore").value = Null
            cmd.Parameters("Sesso").value = Null
            
            
            
            cmd.Execute , , adExecuteNoRecords
            
        End If
        
        
        If ProgressBar1.value > ProgressBar1.Tag Then
            ProgressBar1.Tag = ProgressBar1.Tag + (ProgressBar1.Max / 100)
            DoEvents
        End If
     
    Loop
    

    Me.Caption = "Importazione Immobili - Utenti:" & psFileName(0)
    ProgressBar1.value = 0
    
    Carica = True
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockBatchOptimistic
    rs.Source = "select * from tmpImmobiliUtenti"
    Set rs.ActiveConnection = gDBC
    rs.Open
    Set rs.ActiveConnection = Nothing
    
    ProgressBar1.Min = 0
    ProgressBar1.Max = rs.RecordCount
    ProgressBar1.value = 0
    Me.Caption = "Importazione Immobili - Utenti"
    ProgressBar1.Tag = (ProgressBar1.Max / 100)
    
    
    Do While Not rs.EOF
        
        ProgressBar1.value = ProgressBar1.value + 1
        Me.Caption = "Importazione Immobili - Utenti:" & psFileName(0) & " - " & FormatNumber(ProgressBar1.value * 100 / ProgressBar1.Max, 2)
        
        For i = 0 To rs.Fields.Count - 1
            
            If i > lvwImportazione(0).ColumnHeaders.Count - 1 Then
                lvwImportazione(0).ColumnHeaders.Add , , rs(i).Name
            End If
            
            If i = 0 Then
                Set itm = lvwImportazione(0).ListItems.Add(, , rs(i).value)
            Else
                itm.ListSubItems.Add , , nullToBlank(rs(i).value)
            End If
        Next
        
        If ProgressBar1.value > ProgressBar1.Tag Then
            ProgressBar1.Tag = ProgressBar1.Tag + (ProgressBar1.Max / 100)
            DoEvents
        End If
        
        
        rs.MoveNext
    Loop
    
    lvwAnteprima(0).Visible = False
    
    AllargaLeColonneLVW lvwImportazione(0), False
    lvwImportazione(0).Visible = True
    fText.Close
    Set fText = Nothing
    Set colIdxField = Nothing
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Exit Function
CaricaERR:
    s = vbNullString
    If Not cmd Is Nothing Then
        For i = 0 To cmd.Parameters.Count - 1
            s = s & cmd.Parameters(i).Name & "=" & nullToBlank(cmd.Parameters(i).value) & vbCrLf
        Next
        
    End If
    MsgBox "frmImpImmobiliUtenti Carica" & Err.Description + vbCrLf + s
End Function

Private Function CaricaBadge() As Boolean
    Dim f As New Scripting.FileSystemObject
    Dim fText As TextStream
    Dim s As String
    Dim col() As String
    Dim i As Long
    
    Dim itm  As ListItem
    Dim cmd As ADODB.Command
    Dim byteTot As Long
    Dim fFile As file
    Dim colIdxField As New Collection
    Dim colTracciato As Collection
    Dim pTracciatoMap As TracciatoMapping
    
    Dim rs As ADODB.Recordset
    
    On Error GoTo CaricaBadgeERR
    CaricaBadge = False
    
    
    lvwImportazione(1).ListItems.Clear
    lvwImportazione(1).ColumnHeaders.Clear
    lvwImportazione(1).Visible = False
    
    Set cmd = CreateObject("ADODB.Command")
    cmd.CommandText = "delete * from tmpBadge"
    Set cmd.ActiveConnection = gDBC
    cmd.Execute
    Set cmd.ActiveConnection = Nothing
    
    cmd.CommandText = "INSERT INTO tmpBadge ( CodBadge, CodUtente, TipoAnagrafe, ProgressivoBadge, DataBadgeInizio, DataBadgeFine, " & _
            "CodComune,  TipoUtenza,CodImmobile,progInterno, [default] ) values (?, ?, ?, ?, ?, ?, ?, ?, ?,?, ?)"


    cmd.Prepared = True
    Set cmd.ActiveConnection = gDBC
    
    
    Set colTracciato = colTracciati.item("UTEBADGE")
    For i = 1 To colTracciato("FIELD").Count
        Set pTracciatoMap = colTracciato("FIELD")(i)
        
        s = Split(pTracciatoMap.mappingField, pImpImmUteSepCol)(0)
        If pTracciatoMap.fieldType = FIELD_TYPE_NUM Then
            cmd.Parameters.Append cmd.CreateParameter(s, adBigInt, adParamInput)
        Else
            cmd.Parameters.Append cmd.CreateParameter(s, adVarChar, adParamInput, 100)
        End If
    Next
    
    cmd.Parameters.Append cmd.CreateParameter("default", adBoolean, adParamInput)
    
    'lettura file di testo
    Set fText = f.OpenTextFile(pImpImmUtePath & "\" & psFileName(1), ForReading, False)
    
    Set fFile = f.GetFile(pImpImmUtePath & "\" & psFileName(1))
    byteTot = fFile.size
    
    ProgressBar1.Min = 0
    ProgressBar1.Max = byteTot
    ProgressBar1.value = 0
    Me.Caption = "Importazione Badge:" & psFileName(1)

    Set fFile = Nothing
    
    ProgressBar1.Tag = (ProgressBar1.Max / 100)
    
    Do While Not fText.AtEndOfStream
        s = fText.ReadLine
        col = scomponiRiga(s, pImpImmUteSepCol, pImpImmUteSepDec, pImpImmUteSepText)
        ProgressBar1.value = ProgressBar1.value + Len(s)
        Me.Caption = "Importazione Badge:" & psFileName(1) & " - " & FormatNumber(ProgressBar1.value * 100 / ProgressBar1.Max, 2)
        
        'importazione in tabella temporanea
        If colIdxField.Count = 0 Then
            'creo indice per recuperare i campi da importare successivamente
            For i = LBound(col) To UBound(col)
                Set pTracciatoMap = colTracciato("FIELDEXT")(col(i))
                s = Split(pTracciatoMap.mappingField, pImpImmUteSepCol)(0)
                colIdxField.Add i, s
            Next
        Else
            For i = 0 To cmd.Parameters.Count - 1
                On Error Resume Next
                Set pTracciatoMap = colTracciato("FIELD")(cmd.Parameters(i).Name)
                s = Trim$(col(colIdxField(cmd.Parameters(i).Name)))
                On Error GoTo CaricaBadgeERR:
                'On Error GoTo 0
                'Debug.Assert s <> "3763" Or cmd.Parameters(i).Name <> "CodAnagrafe"
                
                
                Select Case cmd.Parameters(i).Name
                    Case "CodComune"
                        s = IIf(Val(s) = 0, vbNullString, Format(s, "000000"))
                    Case "DataBadgeFine"
                        s = IIf(s = "99/99/9999", "31/12/2050", s)
                    Case "CodImmobile"
                        s = IIf(s = "", "0", s)
                End Select
                
                If pTracciatoMap.fieldFormat <> vbNullString And s <> vbNullString Then s = Format$(s, pTracciatoMap.fieldFormat)
                If pTracciatoMap.fieldType = FIELD_TYPE_CHAR Then
                     s = left$(s, pTracciatoMap.fieldSizeLen)
                End If
                
                'controllo tipo dato
                cmd.Parameters(i).value = IIf(s = vbNullString, Null, s)
            Next
            cmd.Parameters("default").value = cmd.Parameters(9).value = 1
            
            cmd.Execute , , adExecuteNoRecords
            
        End If
        
        
        If ProgressBar1.value > ProgressBar1.Tag Then
            ProgressBar1.Tag = ProgressBar1.Tag + (ProgressBar1.Max / 100)
            DoEvents
        End If
     
    Loop
    

    Me.Caption = "Importazione Badge:" & psFileName(1)
    ProgressBar1.value = 0
    
    CaricaBadge = True
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockBatchOptimistic
    rs.Source = "select * from tmpBadge"
    Set rs.ActiveConnection = gDBC
    rs.Open
    Set rs.ActiveConnection = Nothing
    
    ProgressBar1.Min = 0
    ProgressBar1.Max = rs.RecordCount
    ProgressBar1.value = 0
    Me.Caption = "Importazione Badge:" & psFileName(1)
    ProgressBar1.Tag = (ProgressBar1.Max / 100)
    
    
    Do While Not rs.EOF
        
        ProgressBar1.value = ProgressBar1.value + 1
        Me.Caption = "Importazione Badge:" & FormatNumber(ProgressBar1.value * 100 / ProgressBar1.Max, 2)
        
        For i = 0 To rs.Fields.Count - 1
            
            If i > lvwImportazione(1).ColumnHeaders.Count - 1 Then
                lvwImportazione(1).ColumnHeaders.Add , , rs(i).Name
            End If
            
            If i = 0 Then
                Set itm = lvwImportazione(1).ListItems.Add(, , rs(i).value)
            Else
                itm.ListSubItems.Add , , nullToBlank(rs(i).value)
            End If
        Next
        
        If ProgressBar1.value > ProgressBar1.Tag Then
            ProgressBar1.Tag = ProgressBar1.Tag + (ProgressBar1.Max / 100)
            DoEvents
        End If
        
        
        rs.MoveNext
    Loop
    
    lvwAnteprima(1).Visible = False
    
    AllargaLeColonneLVW lvwImportazione(1), False
    lvwImportazione(1).Visible = True
    fText.Close
    Set fText = Nothing
    Set colIdxField = Nothing
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Exit Function
CaricaBadgeERR:
    If Not cmd Is Nothing Then
        For i = 0 To cmd.Parameters.Count - 1
            s = s & cmd.Parameters(i).Name & "=" & nullToBlank(cmd.Parameters(i).value) & vbCrLf
        Next
        
    End If
    MsgBox "frmImpImmobiliUtenti CaricaBadge" & Error & vbCrLf & s
End Function


Private Function CaricaLegameBarcodeBadge() As Boolean
    Dim f As New Scripting.FileSystemObject
    Dim fText As TextStream
    Dim s As String
    Dim col() As String
    Dim i As Long
    
    Dim itm  As ListItem
    Dim cmd As ADODB.Command
    Dim byteTot As Long
    Dim fFile As file
    Dim colIdxField As New Collection
    Dim colTracciato As Collection
    Dim pTracciatoMap As TracciatoMapping
    
    Dim rs As ADODB.Recordset
    
    On Error GoTo CaricaLegameBarcodeBadgeERR
    CaricaLegameBarcodeBadge = True
    
    If Not legameBarcodeBadge Then Exit Function
    
    lvwImportazione(2).ListItems.Clear
    lvwImportazione(2).ColumnHeaders.Clear
    lvwImportazione(2).Visible = False
    
    Set cmd = CreateObject("ADODB.Command")
    cmd.CommandText = "delete * from tmpLegameBarcodeBadge"
    Set cmd.ActiveConnection = gDBC
    cmd.Execute
    Set cmd.ActiveConnection = Nothing
    
    cmd.CommandText = "INSERT INTO tmpLegameBarcodeBadge ( CodComune, Barcode, CodBadge ) values (?, ?, ?)"


    cmd.Prepared = True
    Set cmd.ActiveConnection = gDBC
    
    
    Set colTracciato = colTracciati.item("LEGBADGE")
    For i = 1 To colTracciato("FIELD").Count
        Set pTracciatoMap = colTracciato("FIELD")(i)
        
        s = Split(pTracciatoMap.mappingField, pImpImmUteSepCol)(0)
        If pTracciatoMap.fieldType = FIELD_TYPE_NUM Then
            cmd.Parameters.Append cmd.CreateParameter(s, adBigInt, adParamInput)
        Else
            cmd.Parameters.Append cmd.CreateParameter(s, adVarChar, adParamInput, 100)
        End If
    Next
    
    
    'lettura file di testo
    Set fText = f.OpenTextFile(pImpImmUtePath & "\" & psFileName(2), ForReading, False)
    Set fFile = f.GetFile(pImpImmUtePath & "\" & psFileName(2))
    
    byteTot = fFile.size
    
    ProgressBar1.Min = 0
    ProgressBar1.Max = byteTot
    ProgressBar1.value = 0
    Me.Caption = "Importazione Legami Badge:" & psFileName(2)

    Set fFile = Nothing
    
    ProgressBar1.Tag = (ProgressBar1.Max / 100)
    
    Do While Not fText.AtEndOfStream
        s = fText.ReadLine
        col = scomponiRiga(s, pImpImmUteSepCol, pImpImmUteSepDec, pImpImmUteSepText)
        ProgressBar1.value = ProgressBar1.value + Len(s)
        Me.Caption = "Importazione Legami Badge:" & psFileName(2) & " - " & FormatNumber(ProgressBar1.value * 100 / ProgressBar1.Max, 2)
        
        'importazione in tabella temporanea
        If colIdxField.Count = 0 Then
            'creo indice per recuperare i campi da importare successivamente
            For i = LBound(col) To UBound(col)
                Set pTracciatoMap = colTracciato("FIELDEXT")(col(i))
                s = Split(pTracciatoMap.mappingField, pImpImmUteSepCol)(0)
                colIdxField.Add i, s
            Next
        Else
            For i = 0 To cmd.Parameters.Count - 1
                On Error Resume Next
                Set pTracciatoMap = colTracciato("FIELD")(cmd.Parameters(i).Name)
                s = Trim$(col(colIdxField(cmd.Parameters(i).Name)))
                On Error GoTo CaricaLegameBarcodeBadgeERR:
                'On Error GoTo 0
                'Debug.Assert s <> "3763" Or cmd.Parameters(i).Name <> "CodAnagrafe"
                
                
                Select Case cmd.Parameters(i).Name
                    Case "CodComune"
                        s = IIf(Val(s) = 0, vbNullString, Format(s, "000000"))
                End Select
                
                If pTracciatoMap.fieldFormat <> vbNullString And s <> vbNullString Then s = Format$(s, pTracciatoMap.fieldFormat)
                If pTracciatoMap.fieldType = FIELD_TYPE_CHAR Then
                     s = left$(s, pTracciatoMap.fieldSizeLen)
                End If
                
                'controllo tipo dato
                cmd.Parameters(i).value = IIf(s = vbNullString, Null, s)
            Next
            
            cmd.Execute , , adExecuteNoRecords
        End If
        
        
        If ProgressBar1.value > ProgressBar1.Tag Then
            ProgressBar1.Tag = ProgressBar1.Tag + (ProgressBar1.Max / 100)
            DoEvents
        End If
     
    Loop
    

    Me.Caption = "Importazione Legami Badge:" & psFileName(2)
    ProgressBar1.value = 0
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockBatchOptimistic
    rs.Source = "select * from tmpLegameBarcodeBadge"
    Set rs.ActiveConnection = gDBC
    rs.Open
    Set rs.ActiveConnection = Nothing
    
    ProgressBar1.Min = 0
    ProgressBar1.Max = rs.RecordCount
    ProgressBar1.value = 0
    Me.Caption = "Importazione Legami Badge:" & psFileName(2)
    ProgressBar1.Tag = (ProgressBar1.Max / 100)
    
    
    Do While Not rs.EOF
        
        ProgressBar1.value = ProgressBar1.value + 1
        Me.Caption = "Importazione Legami Badge:" & FormatNumber(ProgressBar1.value * 100 / ProgressBar1.Max, 2)
        
        For i = 0 To rs.Fields.Count - 1
            
            If i > lvwImportazione(2).ColumnHeaders.Count - 1 Then
                lvwImportazione(2).ColumnHeaders.Add , , rs(i).Name
            End If
            
            If i = 0 Then
                Set itm = lvwImportazione(2).ListItems.Add(, , rs(i).value)
            Else
                itm.ListSubItems.Add , , nullToBlank(rs(i).value)
            End If
        Next
        
        If ProgressBar1.value > ProgressBar1.Tag Then
            ProgressBar1.Tag = ProgressBar1.Tag + (ProgressBar1.Max / 100)
            DoEvents
        End If
        
        
        rs.MoveNext
    Loop
    
    lvwAnteprima(2).Visible = False
    
    AllargaLeColonneLVW lvwImportazione(2), False
    lvwImportazione(2).Visible = True
    fText.Close
    Set fText = Nothing
    Set colIdxField = Nothing
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Exit Function
CaricaLegameBarcodeBadgeERR:
    CaricaLegameBarcodeBadge = False
    If Not cmd Is Nothing Then
        For i = 0 To cmd.Parameters.Count - 1
            s = s & cmd.Parameters(i).Name & "=" & nullToBlank(cmd.Parameters(i).value) & vbCrLf
        Next
        
    End If
    MsgBox "frmImpImmobiliUtenti CaricaLegameiBadge" & Error & vbCrLf & s
End Function



Private Function ImportaBadge() As Boolean
    Dim Command As ADODB.Command
       
    ImportaBadge = False
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ImportaBadgeERR
    
    ProgressBar1.Max = 2
    ProgressBar1.Min = 1
    
    
    'InsBadgeFromTMPBadge
    Set Command = CreateObject("ADODB.Command")
    Command.CommandText = "InsBadgeFromTMPBadge"
    Command.CommandType = adCmdStoredProc
    Set Command.ActiveConnection = gDBC
    Command.Execute , , ADODB.adExecuteNoRecords
    ProgressBar1.value = 1
 
    DoEvents
    
    'UpdBadgeFromTMPBadge
    Set Command = CreateObject("ADODB.Command")
    Command.CommandText = "UpdBadgeFromTMPBadge"
    Command.CommandType = adCmdStoredProc
    Command.Prepared = True
    Set Command.ActiveConnection = gDBC
    Command.CommandTimeout = 0
    Command.Execute , , ADODB.adExecuteNoRecords
    ProgressBar1.value = 2
    
    ImportaBadge = True
    
    Screen.MousePointer = vbDefault
    Exit Function
ImportaBadgeERR:
    Screen.MousePointer = vbDefault
    MsgBox "frmImpImmobiliUtenti ImportaBadge:" & Err.Description

    
End Function

Private Function ImportaLegamiBadge() As Boolean
    Dim Command As ADODB.Command
       
    ImportaLegamiBadge = False
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ImportaLegamiBadgeERR
    
    ProgressBar1.Max = 2
    ProgressBar1.Min = 1
    
    'delLegameBadge
    Set Command = CreateObject("ADODB.Command")
    Command.CommandText = "delete from LegameBarcodeBadge"
    Command.CommandType = adCmdText
    Set Command.ActiveConnection = gDBC
    Command.Execute , , ADODB.adExecuteNoRecords
    ProgressBar1.value = 1
       
    
    'InsLegameBadgeFromTMPLegameBadge
    Set Command = CreateObject("ADODB.Command")
    Command.CommandText = "InsLegameBadgeFromTMPLegameBadge"
    Command.CommandType = adCmdStoredProc
    Set Command.ActiveConnection = gDBC
    Command.Execute , , ADODB.adExecuteNoRecords
    ProgressBar1.value = 2

    Set Command = Nothing
    
    DoEvents
    
    ImportaLegamiBadge = True
    
    Screen.MousePointer = vbDefault
    Exit Function
ImportaLegamiBadgeERR:
    Screen.MousePointer = vbDefault
    MsgBox "frmImpImmobiliUtenti ImportaLegami:" & Err.Description

End Function


Private Function Importa() As Boolean
    Dim Command As ADODB.Command
       
    Importa = False
    Screen.MousePointer = vbHourglass
    
    On Error Resume Next
    gDBC.Execute "drop table tmpUtentiDaAggiornare", , ADODB.adExecuteNoRecords
    
    On Error GoTo ImportaERR
    
    ProgressBar1.Max = 6
    ProgressBar1.Min = 1
    
    'viewUtentiFromTMPImmobiliUtenti
    Set Command = CreateObject("ADODB.Command")
    Command.CommandText = "viewUtentiFromTMPImmobiliUtenti"
    Command.CommandType = adCmdStoredProc
    Set Command.ActiveConnection = gDBC
    Command.Execute , , ADODB.adExecuteNoRecords
    ProgressBar1.value = 1
    
    'InsTariffaFromTMPImmobiliUtenti
    Set Command = CreateObject("ADODB.Command")
    Command.CommandText = "InsTariffaFromTMPImmobiliUtenti"
    Command.CommandType = adCmdStoredProc
    Set Command.ActiveConnection = gDBC
    Command.Execute , , ADODB.adExecuteNoRecords
    ProgressBar1.value = 2
    
    'UpdImmobiliFromTMPImmobiliUtenti
    Set Command = CreateObject("ADODB.Command")
    Command.CommandText = "UpdImmobiliFromTMPImmobiliUtenti"
    Command.CommandType = adCmdStoredProc
    Command.Prepared = True
    Set Command.ActiveConnection = gDBC
    Command.CommandTimeout = 0
    Command.Execute , , ADODB.adExecuteNoRecords
    ProgressBar1.value = 3
    
    'UPDUtentiFromTMPImmobiliUtenti
    Set Command = CreateObject("ADODB.Command")
    Command.CommandText = "UPDUtentiFromTMPImmobiliUtenti"
    Command.CommandType = adCmdStoredProc
    Set Command.ActiveConnection = gDBC
    Command.Execute , , ADODB.adExecuteNoRecords
    ProgressBar1.value = 4
    
    'InsImmobiliFromTMPImmobiliUtenti
    Set Command = CreateObject("ADODB.Command")
    Command.CommandText = "InsImmobiliFromTMPImmobiliUtenti"
    Command.CommandType = adCmdStoredProc
    Set Command.ActiveConnection = gDBC
    Command.Execute , , ADODB.adExecuteNoRecords
    ProgressBar1.value = 5
    
    'InsUtentiFromTMPImmobiliUtenti
    Set Command = CreateObject("ADODB.Command")
    Command.CommandText = "InsUtentiFromTMPImmobiliUtenti"
    Command.CommandType = adCmdStoredProc
    Set Command.ActiveConnection = gDBC
    Command.Execute , , ADODB.adExecuteNoRecords
    ProgressBar1.value = 6
    
    Importa = True
    
    Screen.MousePointer = vbDefault
    Exit Function
ImportaERR:
    MsgBox "frmImpImmobiliUtenti Importa:" & Err.Description
    Screen.MousePointer = vbDefault
End Function



Private Sub Picture1_Resize()
    Line1(1).X2 = Picture1.ScaleWidth - 50
    Line1(0).Y1 = Line1(1).Y1 + 8
    Line1(0).Y2 = Line1(1).Y1 + 8
    Line1(0).X1 = Line1(1).X1 + 12
    Line1(0).X2 = Picture1.ScaleWidth - 50
End Sub


