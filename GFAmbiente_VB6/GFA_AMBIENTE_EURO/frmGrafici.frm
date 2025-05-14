VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmGrafici 
   Caption         =   "Form1"
   ClientHeight    =   6570
   ClientLeft      =   4605
   ClientTop       =   3510
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   8190
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      ScaleHeight     =   495
      ScaleWidth      =   4875
      TabIndex        =   1
      Top             =   5820
      Width           =   4875
      Begin VB.CommandButton Command1 
         Caption         =   "&Stampa"
         Height          =   375
         Index           =   3
         Left            =   60
         TabIndex        =   5
         Top             =   60
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Impostazioni"
         Height          =   375
         Index           =   2
         Left            =   1260
         TabIndex        =   4
         Top             =   60
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ridisegna"
         Height          =   375
         Index           =   1
         Left            =   2460
         TabIndex        =   3
         Top             =   60
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Chiudi"
         Height          =   375
         Index           =   0
         Left            =   3660
         TabIndex        =   2
         Top             =   60
         Width           =   1095
      End
   End
   Begin MSChart20Lib.MSChart MSChart 
      Height          =   5235
      Left            =   240
      OleObjectBlob   =   "frmGrafici.frx":0000
      TabIndex        =   0
      Top             =   240
      Width           =   7635
   End
   Begin MSAdodcLib.Adodc ADO1 
      Height          =   330
      Left            =   2520
      Top             =   5880
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmGrafici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private pCodComune As String
Private pSuperfici() As Double
Private pDesLabelCoordX As Collection

Private pAnniDiConfronto As String
Private pAnnoRiferimento As Integer

Private oldpart As Long
Private Type Grafico
    chartType As MSChart20Lib.VtChChartType
    FootnoteText As String
    TitleText As String
    AxisTitleX As String
    AxisTitleY As String
    indicatori As Boolean
    serieUnite As Boolean
    legenda As Boolean

End Type
    
Private Grafici(10) As Grafico
Private Const CMD_CHIUDI = 0
Private Const CMD_RIDISEGNA = 1
Private Const CMD_IMPOSTAZIONI = 2
Private Const CMD_STAMPA = 3

Private pIDchart As Integer
Private pDesComune As String

Public Sub ShowChart(idChart As Integer, Optional sCodComune As String = vbNullString, Optional AnnoRiferimento As String)
    Dim i As Integer
    Dim rs As ADODB.Recordset
    
    'se la modalità è tutti i comuni chiedo di quale comune si vanno ad inserire i dati
    If sCodComune = vbNullString Then
        AnnoRiferimento = IIf(AnnoRiferimento = vbNullString, Year(Now), AnnoRiferimento)

        If gComuneAttivo = ALL_COMUNI Then
            frmComuneDaLavorare.Show vbModal
            sCodComune = GetSetting(App.Title, SEZ_ULTIMI_VALORI, KEY_ULTIMI_VALORI_COMUNESEL, "")
        Else
            sCodComune = gComuneAttivo
        End If
    End If
    
    
    pDesComune = getDesComune(sCodComune)
    pAnnoRiferimento = AnnoRiferimento
    pAnniDiConfronto = AnnoRiferimento & "|" & AnnoRiferimento - 1
    
    pCodComune = sCodComune
    
    ReDim pSuperfici(12)
    For i = LBound(pSuperfici) To UBound(pSuperfici)
        pSuperfici(i) = 60 + (i * 20)
    Next
    
    pIDchart = idChart
    defaultSetting
    createChart
        
    Me.Caption = "Grafico:" & pDesComune & " - " & Grafici(pIDchart).TitleText
        

    Me.Show
    

End Sub

Private Sub defaultSetting()
    Dim rs As ADODB.Recordset
    'Grafico Utenza Domestica 1 componente
    
    Set rs = gDBC.Execute("select * from ChartDefault where idChart=" & pIDchart)
    
    Grafici(pIDchart).TitleText = rs("Titolo")
    Grafici(pIDchart).FootnoteText = rs("NoteGrafico")
    Grafici(pIDchart).chartType = rs("TipoGrafico")
    Grafici(pIDchart).AxisTitleX = rs("TitoloAsseX")
    Grafici(pIDchart).AxisTitleY = rs("TitoloAsseY")
    Grafici(pIDchart).indicatori = CBool(rs("Indicatori"))
    Grafici(pIDchart).serieUnite = CBool(rs("SerieUnite"))
    Grafici(pIDchart).legenda = CBool(rs("Legenda"))
End Sub

Private Sub changeSetting()
    Dim marker As Object
    Dim i As Integer
    
    With MSChart
        .TextLengthType = VtTextLengthTypeVirtual
        .RandomFill = False
        
        .FootnoteText = Grafici(pIDchart).FootnoteText & " - Comune di " & pDesComune
        .Footnote.Location.LocationType = VtChLocationTypeCustom
        .Footnote.Location.RECT.Min.Set 260, 50
        .Footnote.Location.RECT.Max.Set 9000, 350
        .Footnote.TextLayout.HorzAlignment = VtHorizontalAlignmentLeft
        
        .TitleText = Grafici(pIDchart).TitleText
       
        .Plot.Axis(VtChAxisIdX, 1).AxisTitle = Grafici(pIDchart).AxisTitleX
        .Plot.Axis(VtChAxisIdY, 1).AxisTitle = Grafici(pIDchart).AxisTitleY
        
        .Stacking = Grafici(pIDchart).serieUnite
        If .DataGrid.ColumnCount > 0 Then
            If .Plot.SeriesCollection.Count > 0 Then
                For i = 1 To .Plot.SeriesCollection.Count
                    With .Plot.SeriesCollection(i).DataPoints(-1).marker
                        .Visible = Grafici(pIDchart).indicatori
                        .size = 2
                    End With
                    .Plot.SeriesCollection(i).Pen.Width = 2
                Next
            End If
        End If
       .chartType = Grafici(pIDchart).chartType
       .ShowLegend = Grafici(pIDchart).legenda
       
    End With
End Sub

Private Sub Command1_Click(index As Integer)
    Dim pOLDAnniDiConfronto As String

    
    Select Case index
        Case CMD_CHIUDI
            SalvaDefault
            Unload Me
        Case CMD_RIDISEGNA
            createChart
        Case CMD_IMPOSTAZIONI
        
            pOLDAnniDiConfronto = pAnniDiConfronto
            
            If Not frmGraficiImpostazioni.showImpostazioni(Grafici(pIDchart).TitleText, _
                Grafici(pIDchart).FootnoteText, _
                Grafici(pIDchart).AxisTitleX, _
                Grafici(pIDchart).AxisTitleY, _
                Grafici(pIDchart).chartType, _
                Grafici(pIDchart).indicatori, _
                Grafici(pIDchart).legenda, _
                Grafici(pIDchart).serieUnite, _
                pOLDAnniDiConfronto, _
                pAnnoRiferimento, _
                pCodComune) Then
                    If pAnniDiConfronto <> pOLDAnniDiConfronto Then
                        pAnniDiConfronto = pOLDAnniDiConfronto
                        createChart
                    Else
                        changeSetting
                    End If
            End If
            
        Case CMD_STAMPA
            PrintChart
    End Select
    
End Sub


Private Sub PrintChart()
    Dim pColor As Long
        
    
    Printer.Orientation = vbPRORLandscape

    pic.Visible = False
    pColor = Me.BackColor
    Me.BackColor = vbWhite
    Me.PrintForm
    Me.BackColor = pColor
    pic.Visible = True
    
End Sub


Private Sub Form_Load()
    
    
    ADO1.ConnectionString = gDBC
   
End Sub

Private Sub Form_Resize()
    If WindowState = vbMinimized Then Exit Sub
    pic.Move Me.ScaleWidth - pic.Width, Me.ScaleHeight - pic.Height
    MSChart.Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight - pic.Height

End Sub


Private Sub createChart()
    Dim i As Integer, ii As Integer
    Dim tariffaAnno As Double
    Dim anni() As String
    Dim rsAnni() As Recordset
    Dim sCommand As String
    
    
    If pAnniDiConfronto = vbNullString Then Exit Sub
    MSChart.Visible = False
    gDBC.Execute "DROP TABLE ChartData"
    
    anni = Split(pAnniDiConfronto, "|")
    sCommand = "CREATE TABLE ChartData (label varchar(50)"
    For i = LBound(anni) To UBound(anni)
        sCommand = sCommand & ",valore" & CStr(i) & " numeric(18,5) default 0"
    Next
    
    gDBC.Execute sCommand & ")"
    ReDim rsAnni(UBound(anni))
    
    Set pDesLabelCoordX = New Collection
    gDBC.BeginTrans
    If pIDchart >= 1 And pIDchart <= 6 Then
        For ii = LBound(anni) To UBound(anni)
           Set rsAnni(ii) = SelezionaUtenzeDomestiche(pCodComune, anni(ii))
           rsAnni(ii).Filter = "IdNumeriComponenti=" & CStr(pIDchart)
        
        
           Do While Not rsAnni(ii).EOF
                For i = LBound(pSuperfici) To UBound(pSuperfici)
    
                    tariffaAnno = (rsAnni(ii)("TariffaNComp") + rsAnni(ii)("TariffaSuperficie") * pSuperfici(i))
                    If ii = 0 Then
                        gDBC.Execute "insert into ChartData (label,valore0) values('" & pSuperfici(i) & "'," & FormatFLD(tariffaAnno, adDecimal) & ")"
                    Else
                        gDBC.Execute "update ChartData set valore" & CStr(ii) & "=" & FormatFLD(tariffaAnno, adDecimal) & " where label='" & pSuperfici(i) & "'"
                    End If
                Next
                rsAnni(ii).MoveNext
           Loop
        Next
     ElseIf pIDchart = 7 Or pIDchart = 8 Then
        For ii = LBound(anni) To UBound(anni)
            Set rsAnni(ii) = SelezionaUtenzeAttivita(pCodComune, anni(ii))
            rsAnni(ii).Filter = IIf(pIDchart = 7, "IdCategoriaAttivita <= 15", "IdCategoriaAttivita > 15")
    
    
            Do While Not rsAnni(ii).EOF
                tariffaAnno = rsAnni(ii)("tariffa")
                If ii = 0 Then
                    gDBC.Execute "insert into ChartData (label,valore0) values('" & rsAnni(ii)("IdCategoriaAttivita") & "'," & FormatFLD(tariffaAnno, adDecimal) & ")"
                    pDesLabelCoordX.Add rsAnni(ii)("Descrizione").Value, "|" & rsAnni(ii)("IdCategoriaAttivita")
                Else
                    gDBC.Execute "update ChartData set valore" & CStr(ii) & "=" & FormatFLD(tariffaAnno, adDecimal) & " where label='" & rsAnni(ii)("IdCategoriaAttivita") & "'"
                End If
                rsAnni(ii).MoveNext
            Loop
        Next
    ElseIf pIDchart = 9 Then
        For ii = LBound(anni) To UBound(anni)
           Set rsAnni(ii) = SelezionaUtenzeDomestiche(pCodComune, anni(ii))
           
           Do While Not rsAnni(ii).EOF
                tariffaAnno = rsAnni(ii)("TariffaSuperficie")
                If ii = 0 Then
                    gDBC.Execute "insert into ChartData (label,valore0) values('" & rsAnni(ii)("IdNumeriComponenti") & "'," & FormatFLD(tariffaAnno, adDecimal) & ")"
                Else
                    gDBC.Execute "update ChartData set valore" & CStr(ii) & "=" & FormatFLD(tariffaAnno, adDecimal) & " where label='" & rsAnni(ii)("IdNumeriComponenti") & "'"
                End If
                rsAnni(ii).MoveNext
           Loop
        Next
    End If
    
    
    gDBC.CommitTrans
    ADO1.RecordSource = "SELECT label"
    For i = LBound(anni) To UBound(anni)
        ADO1.RecordSource = ADO1.RecordSource & ",valore" & CStr(i) & " as " & anni(i)
    Next
    ADO1.RecordSource = ADO1.RecordSource & " from ChartData"
    ADO1.Refresh
    Set MSChart.DataSource = ADO1
    changeSetting
    
    MSChart.Visible = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set pGrafici(pIDchart) = Nothing
End Sub

Private Sub MSChart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim part As Integer
    Dim index1 As Integer
    Dim index2 As Integer
    Dim index3 As Integer
    Dim index4 As Integer
    


    
    MSChart.TwipsToChartPart X, Y, part, index1, index2, index3, index4

    If VtChPartTypePoint = part Then
        MSChart.Column = index1
        MSChart.Row = index2
        If (pIDchart = 7 Or pIDchart = 8) And Not pDesLabelCoordX Is Nothing Then
            MSChart.ToolTipText = pDesLabelCoordX("|" & MSChart.RowLabel) & ":" & MSChart.data
            
        Else
            MSChart.ToolTipText = MSChart.data
        End If
        
        oldpart = part
    End If
    
End Sub

Private Sub MSChart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim part As Integer
    Dim index1 As Integer
    Dim index2 As Integer
    Dim index3 As Integer
    Dim index4 As Integer
    
    On Error Resume Next
    MSChart.TwipsToChartPart X, Y, part, index1, index2, index3, index4
    
    If oldpart <> part Then
        MSChart.ToolTipText = ""
        
    End If
End Sub


Private Sub SalvaDefault()
    gDBC.Execute "update ChartDefault set Titolo=" & FormatFLD(Grafici(pIDchart).TitleText, adChar) _
            & ",TitoloAsseX=" & FormatFLD(Grafici(pIDchart).AxisTitleX, adChar) _
            & ",TitoloAsseY=" & FormatFLD(Grafici(pIDchart).AxisTitleY, adChar) _
            & ",NoteGrafico=" & FormatFLD(Grafici(pIDchart).FootnoteText, adChar) _
            & ",SerieUnite=" & FormatFLD(Grafici(pIDchart).serieUnite, adBoolean) _
            & ",Legenda=" & FormatFLD(Grafici(pIDchart).legenda, adBoolean) _
            & ",Indicatori=" & FormatFLD(Grafici(pIDchart).indicatori, adBoolean) _
            & ",TipoGrafico=" & FormatFLD(Grafici(pIDchart).chartType, adInteger) _
            & " where idChart=" & pIDchart
End Sub
