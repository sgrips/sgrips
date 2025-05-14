VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportazioneXLS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importazione da XLS"
   ClientHeight    =   6105
   ClientLeft      =   6615
   ClientTop       =   4305
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   5520
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvwTabelle 
      Height          =   2175
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "tabelle"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgXLS 
      Left            =   5160
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.xls"
      DialogTitle     =   "Importazione da file XLS"
      FileName        =   "*.xls"
      Filter          =   "*.xls"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Apri"
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Chiudi"
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   2
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Importa"
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   1
      Top             =   5640
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lvwAnni 
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1508
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "anni"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblRigheXLS 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
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
      Left            =   4680
      TabIndex        =   13
      Top             =   4920
      Width           =   585
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Righe Lette"
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
      Index           =   7
      Left            =   3480
      TabIndex        =   12
      Top             =   4920
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Importazione"
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
      Index           =   6
      Left            =   120
      TabIndex        =   9
      Top             =   4920
      Width           =   1110
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Comune"
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
      Index           =   5
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   5160
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Comune"
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
      Index           =   3
      Left            =   1320
      TabIndex        =   7
      Top             =   120
      Width           =   840
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Codice Istat"
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
      TabIndex        =   6
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Anni"
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
      TabIndex        =   5
      Top             =   3600
      Width           =   390
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Tabelle"
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
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "frmImportazioneXLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CMD_OPEN = 0
Private Const CMD_IMPORT = 2
Private Const CMD_CANCEL = 1

Private xlsApp As Object
Private pFileName As String


Private Sub Command1_Click(index As Integer)
    lblRigheXLS.Caption = 0
    Select Case index
        Case CMD_OPEN
            dlgXLS.FileName = "*.xls"
            dlgXLS.Filter = "Microsoft Excel (*.xls)|(*.xls)"
            dlgXLS.FilterIndex = 0
            dlgXLS.CancelError = False
            dlgXLS.ShowOpen
            
            If (dlgXLS.FileName <> "" And dlgXLS.FileName <> "*.xls") Then
                pFileName = dlgXLS.FileName
                readXLS (dlgXLS.FileName)
            End If
        Case CMD_IMPORT
            importXLS (pFileName)
        Case CMD_CANCEL
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    ProgressBar1.Visible = False
    Command1(CMD_IMPORT).Enabled = False
    
    lvwTabelle.Enabled = False
    lvwAnni.Enabled = False
    
    lvwTabelle.ColumnHeaders(1).Width = lvwTabelle.Width - 500
    lvwAnni.ColumnHeaders(1).Width = lvwAnni.Width - 100
    
    Set xlsApp = CreateObject("Excel.Application")
End Sub

Private Sub importXLS(FileName As String)
    On Error GoTo importXLSERR
    Dim wrkBook As Object
    Dim wrkSheet As Object
    Dim i As Long
    
    Set wrkBook = xlsApp.Workbooks.Open(FileName)
    xlsApp.DisplayAlerts = False
    Set wrkSheet = wrkBook.Sheets(1)
    
    Command1(CMD_CANCEL).Enabled = False
    Command1(CMD_IMPORT).Enabled = False
    Command1(CMD_OPEN).Enabled = False
    
    ProgressBar1.Visible = True
    ProgressBar1.Max = lvwTabelle.ListItems.Count
    ProgressBar1.value = 0
    
    For i = 1 To lvwTabelle.ListItems.Count
        If lvwTabelle.ListItems(i).Checked Then
            importTable lvwTabelle.ListItems(i).Text, wrkSheet
        End If
        ProgressBar1.value = i
    Next
    
    Command1(CMD_CANCEL).Enabled = True
    Command1(CMD_IMPORT).Enabled = True
    Command1(CMD_OPEN).Enabled = True
    
    
    MsgBox "Operazione conclusa", vbExclamation + vbOKOnly, "Importazione dati XLS"
    
    xlsApp.Workbooks.Close
    Set wrkBook = Nothing
    Set wrkSheet = Nothing
    
    Exit Sub
importXLSERR:

    Command1(CMD_CANCEL).Enabled = True
    Command1(CMD_IMPORT).Enabled = True
    Command1(CMD_OPEN).Enabled = True

    If Not wrkBook Is Nothing Then xlsApp.Workbooks.Close
    Set wrkBook = Nothing
    Set wrkSheet = Nothing
    
    MsgBox Err.description, vbCritical + vbOKOnly, App.Title
End Sub

Private Sub importTable(nomeTable As String, wrkSheet As Object)
    Dim row As Long
    Dim rowCol As Long
    Dim valueCell As String
    Dim CodComune As String
    Dim primavolta As Boolean
    
    'cerco la tabella da importare
    primavolta = True
    Do While valueCell <> "BonusUtenzeDomestiche"
        rowCol = rowCol + 1
        valueCell = wrkSheet.Range("A" + CStr(rowCol)).Text
        
        If valueCell = nomeTable Then
            row = rowCol
            Do
                CodComune = wrkSheet.Range("B" + CStr(row + 1))
                If CodComune <> "" Then
                    'importo i dati
                    importTableDB wrkSheet, row, valueCell, rowCol, primavolta
                End If
                primavolta = False
                row = row + 1
                lblRigheXLS.Caption = row
            Loop While CodComune <> ""
        End If
        
     Loop
     

End Sub


Private Sub importTableDB(wrkSheet As Object, row As Long, ByVal tabella As String, rowCol As Long, primavolta As Boolean)
    Dim qryINS As ADODB.Command
    Dim qryUPD As ADODB.Command
    Static rs As Recordset
    Dim CodComune As String
    Dim annoComune As Integer
    Dim col As Long
    Dim i As Long
    Dim s As String
    Dim param As ADODB.Parameter
    Dim Insert As Boolean
    Static ParAbitanti As String
    Static ParteFissa As Long
    Static AreaGeografica As String
    
    On Error GoTo importTableDBERR
    
    CodComune = wrkSheet.Range("B" + CStr(row + 1))
    annoComune = wrkSheet.Range("C" + CStr(row + 1))
    
    Select Case tabella
        Case "AnagraficaComunale"
            Set rs = DatiAnagraficiComunali(CodComune, annoComune)
        Case "CostoComunale"
            Set rs = DatiCostiComune(CodComune, annoComune)
        Case "DatiTecniciComune"
            Set rs = DatiTecniciComune(CodComune, annoComune)
        Case "UtenzeDomesticheComunali"
            If primavolta Then Set rs = SelezionaUtenzeDomestiche(CodComune, annoComune)
        Case "CategorieAttivitaComunali"
            If primavolta Then Set rs = SelezionaUtenzeAttivita(CodComune, annoComune)
        Case "ValoriCoefficientiUDComunaliKA", "ValoriCoefficientiUDComunaliKB"
            
            If primavolta Then
                Set rs = DatiAnagraficiComunali(CodComune, annoComune, ParAbitanti)
                
                If Right$(tabella, 2) = "KB" Then
                    AreaGeografica = 0
                    ParAbitanti = 0
                    Set rs = SelezionaCoeffK(CoeffKB, CodComune, annoComune, 0, 0)
                Else
                    AreaGeografica = rs("AreaGeografica")
                    Set rs = SelezionaCoeffK(CoeffKA, CodComune, annoComune, rs("AreaGeografica"), ParAbitanti)
                    
                End If
                
            End If
            ParteFissa = IIf(tabella = "ValoriCoefficientiUDComunaliKA", -1, 0)
            tabella = Left$(tabella, Len(tabella) - 2)
    
        Case "ValoriCoefficientiCAComunaliKC", "ValoriCoefficientiCAComunaliKD"
            
            If primavolta Then
                Set rs = DatiAnagraficiComunali(CodComune, annoComune, ParAbitanti)
                AreaGeografica = rs("AreaGeografica")
                If Right$(tabella, 2) = "KC" Then
                    Set rs = SelezionaCoeffK(CoeffKC, CodComune, annoComune, rs("AreaGeografica"), ParAbitanti)
                Else
                    Set rs = SelezionaCoeffK(CoeffKD, CodComune, annoComune, rs("AreaGeografica"), ParAbitanti)
                End If
                
            End If
            ParteFissa = IIf(tabella = "ValoriCoefficientiCAComunaliKC", -1, 0)
            tabella = Left$(tabella, Len(tabella) - 2)
        Case "ParametriSistema"
            col = 4
            LetParamValue annoComune, CodComune, wrkSheet.Cells(row + 1, col).Text, True, wrkSheet.Cells(row + 1, col + 2).Text
            Exit Sub
        Case "BonusUtenzeAttivita"
            If primavolta Then
                Set rs = SelezionaBonus(BonusUtenzaAttivita, CodComune, annoComune, TariffaNormalizzata)
                
                If Not rs.EOF Then
                    gDBC.Execute ("delete * from BonusUtenzeAttivita where CodComune='" & CodComune & _
                        "' and AnnoRiferimento='" & annoComune & _
                        "' and CodTipoTariffa=" & rs("CodTipoTariffa") & "")

                    rs.Requery

                End If
            End If
        Case "BonusUtenzeDomestiche"
            If primavolta Then
                Set rs = SelezionaBonus(BonusUtenzaDomestica, CodComune, annoComune, TariffaNormalizzata)
                
                If Not rs.EOF Then
                    gDBC.Execute ("delete * from BonusUtenzeDomestiche where CodComune='" & CodComune & _
                        "' and AnnoRiferimento='" & annoComune & _
                        "' and CodTipoTariffa=" & rs("CodTipoTariffa") & "")

                    rs.Requery

                End If
            End If
    End Select
    
    col = 2
    If rs.EOF Then
        Set qryINS = gADOX.Procedures("Ins" + tabella).Command
        qryINS.Parameters.Refresh
        
        For i = 1 To qryINS.Parameters.Count
            
            If ((tabella = "ValoriCoefficientiUDComunali" Or tabella = "ValoriCoefficientiCAComunali") And i > 5) Then
                qryINS.Parameters("pCodParGeografica").value = AreaGeografica
                qryINS.Parameters("pCodRange").value = ParAbitanti
                qryINS.Parameters("pParteFissa").value = ParteFissa
                
                Exit For
            Else
                
                Set param = qryINS.Parameters("p" + wrkSheet.Cells(rowCol, col).Text)
                param.value = FormatFLDFromXLS(wrkSheet.Cells(row + 1, col).value, rs.fields(wrkSheet.Cells(rowCol, col).Text).Type)
            End If
            
            col = col + 1
        Next
        
        If (tabella = "BonusUtenzeAttivita") Or (tabella = "BonusUtenzeDomestiche") Then
            'verifico che il valore sia > 0
            If (qryINS.Parameters("pValore").value = 0) Then
                Set qryINS = Nothing
                Exit Sub
            End If
        End If
        
        qryINS.Execute
        Set qryINS = Nothing
    Else
        Set qryUPD = gADOX.Procedures("Upd" + tabella).Command
        qryUPD.Parameters.Refresh
        
        For i = 1 To qryUPD.Parameters.Count
            
            If ((tabella = "ValoriCoefficientiUDComunali" Or tabella = "ValoriCoefficientiCAComunali") And i > 5) Then
                
                qryUPD.Parameters("pCodParGeografica").value = AreaGeografica
                qryUPD.Parameters("pCodRange").value = ParAbitanti
                qryUPD.Parameters("pParteFissa").value = ParteFissa
                
                Exit For
            Else
                qryUPD.Parameters("p" + wrkSheet.Cells(rowCol, col).Text).value = FormatFLDFromXLS(wrkSheet.Cells(row + 1, col).value, rs.fields(wrkSheet.Cells(rowCol, col).Text).Type)
            End If
            
            col = col + 1
        Next
        
        qryUPD.Execute
        Set qryUPD = Nothing

    End If
    
   
    Exit Sub
importTableDBERR:
    If Not Nothing Is qryINS Then
        If i < qryINS.Parameters.Count Then
            s = "Parametro:" & qryINS.Parameters(i).Name & "[" & wrkSheet.Cells(row, col).Text & "]" & vbCrLf
        End If
    End If
    
    If Not Nothing Is qryUPD Then
        If i < qryUPD.Parameters.Count Then
            s = "Parametro:" & qryUPD.Parameters(i).Name & "[" & wrkSheet.Cells(row, col).Text & "]" & vbCrLf
        End If
    End If
    
    Err.Raise 999, "importTableDB", "Errore nei dai relativi alla Tabella:" & tabella & vbCrLf & s & Err.description
End Sub

Function FormatFLDFromXLS(ByVal s As String, ByVal Tipo As DataTypeEnum) As Variant
    'Formatta il campo a seconda del tipo
    Dim i As Integer

    If (s = vbNullString) Then
        FormatFLDFromXLS = Null
    Else
        Select Case Tipo
            Case adChar, adVarWChar, adVarChar
                FormatFLDFromXLS = s
            Case adDate
                FormatFLDFromXLS = Format$(s, "YYYY-MM-DD")
            Case adBoolean
                s = UCase$(s)
                FormatFLDFromXLS = IIf(s = "TRUE" Or s = "VERO", "-1", "0")
            Case Else
                i = InStr(s, ".")
                If i > 0 Then Mid$(s, i) = ","
                FormatFLDFromXLS = s
        End Select
    
    End If
    
    
End Function


Private Sub readXLS(FileName As String)
    On Error GoTo readXLSERR
    Dim wrkBook As Object
    Dim wrkSheet As Object
    Dim row As Long
    Dim rowAnni As Long
    Dim valueCell As String
    Dim anniCell As String
    Dim xlsOK As Boolean
    Dim Item As ListItem
    Dim codIstat As String
    
    Set wrkBook = xlsApp.Workbooks.Open(FileName)
    xlsApp.DisplayAlerts = False
    
    Set wrkSheet = wrkBook.Sheets(1)
    
    lvwTabelle.ListItems.Clear
    lvwAnni.ListItems.Clear
    
    lvwTabelle.Enabled = False
    lvwAnni.Enabled = False
    Command1(CMD_IMPORT).Enabled = False
    
    
    Screen.MousePointer = vbHourglass
    DoEvents
    Do While valueCell <> "BonusUtenzeDomestiche"
        row = row + 1
        lblRigheXLS.Caption = row
        valueCell = wrkSheet.Range("A" + CStr(row)).Text
        
        If valueCell <> vbNullString Then
            If (valueCell = "AnagraficaComunale") Then
                xlsOK = True
                'leggo gli anni disponibili
                rowAnni = row
                Do
                    rowAnni = rowAnni + 1
                    anniCell = wrkSheet.Range("C" + CStr(rowAnni)).Text
                    If codIstat = "" Then codIstat = wrkSheet.Range("B" + CStr(rowAnni)).Text
                    If anniCell <> "" Then
                        Set Item = lvwAnni.ListItems.Add(, , anniCell)
                        Item.Checked = True
                    End If
                Loop While anniCell <> ""
                
            ElseIf xlsOK = False Then
                MsgBox "File XLS non valido per l'importazione", vbExclamation + vbOKOnly, App.Title
                Exit Do
            End If
            Set Item = lvwTabelle.ListItems.Add(, , valueCell)
            Item.Checked = True
        End If
    Loop
    Screen.MousePointer = vbDefault
    
    If xlsOK Then
        lvwTabelle.Enabled = True
        lvwAnni.Enabled = True
        Command1(CMD_IMPORT).Enabled = True
        lbl(3).Caption = codIstat
        lbl(5).Caption = getDesComune(codIstat)
    End If
    
    xlsApp.Workbooks.Close
    Set wrkBook = Nothing
    Set wrkSheet = Nothing
    
    
    
    Exit Sub
readXLSERR:
    If Not wrkBook Is Nothing Then xlsApp.Workbooks.Close
    Set wrkBook = Nothing
    Set wrkSheet = Nothing
    
    MsgBox Err.description, vbCritical + vbOKOnly, App.Title
End Sub

Private Sub Form_Unload(Cancel As Integer)
    xlsApp.Workbooks.Close
    Set xlsApp = Nothing
End Sub

