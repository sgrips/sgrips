VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDettaglioMovimenti 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dettaglio Giacenza"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   10155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPeso 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   525
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0"
      Top             =   6053
      Width           =   2055
   End
   Begin MSComctlLib.ListView lvwDettaglio 
      Height          =   6015
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10610
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
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
      NumItems        =   0
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Conferma"
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
      Left            =   7800
      TabIndex        =   1
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Chiudi"
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
      Left            =   9000
      TabIndex        =   0
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label lblcodCer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Top             =   6120
      Width           =   870
   End
End
Attribute VB_Name = "frmDettaglioMovimenti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private pRecordsetMovimenti As String
Private pcodCategoria As String
Private pcodSottoCategoria As String
Private pcodCer As String

Private pretQuantita As Double
Private pretElencoLotti As String
Private pretCode As Integer


Public Function showMe(ByVal codCer As String, ByVal tipoRifiuto As String, quantita As Double, elencoLotti As String) As Boolean
    
    
    pcodCategoria = Left$(codCer, 2)
    pcodSottoCategoria = Mid$(codCer, 3, 2)
    pcodCer = Right$(codCer, 2)
    pretElencoLotti = elencoLotti
    pretQuantita = quantita
    
    Me.Caption = tipoRifiuto
    
    lblcodCer.Caption = codCer
    
    
    Me.Show vbModal
    
    showMe = pretCode = 1
    quantita = pretQuantita
    elencoLotti = pretElencoLotti
End Function

Private Sub cmd_Click(Index As Integer)
    Dim itm As ListItem
    
    pretCode = Index
    Select Case Index
        Case 0
            pretQuantita = txtPeso.Text
            pretElencoLotti = "-1"
            Unload Me
        Case 1
            pretQuantita = txtPeso.Text
            pretElencoLotti = ""
            For Each itm In lvwDettaglio.ListItems
                If itm.Checked Then
                    pretElencoLotti = pretElencoLotti & Mid$(itm.key, 2) & ","
                End If
            Next
            
            If (pretElencoLotti <> vbNullString) Then
                pretElencoLotti = Left$(pretElencoLotti, Len(pretElencoLotti) - 1)
            End If
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Dim colH As ColumnHeader
    
    pRecordsetMovimenti = DevEnvMDB.rscmdGiacenze.Source

    With lvwDettaglio
        Set colH = .ColumnHeaders.Add(, , "Nr.Giorni Giac.")
        Set colH = .ColumnHeaders.Add(, , "Qty.Residuo")
        Set colH = .ColumnHeaders.Add(, , "Dta Primo Mov.")
        Set colH = .ColumnHeaders.Add(, , "Dta Prev Ritiro")
        Set colH = .ColumnHeaders.Add(, , "Dta Prev Rit Normativa")
        colH.Alignment = lvwColumnRight
    End With
    
    caricaDati
End Sub



Private Sub caricaDati()
    Dim orderBy As String
    Dim iorderBy As Integer
    Dim listItm As ListItem
    Dim keysLotti As Collection
    Dim qty As Double
    
    qty = 0
    If (DevEnvMDB.rscmdGiacenze.state = adStateOpen) Then
        DevEnvMDB.rscmdGiacenze.Close
    End If
    
    iorderBy = InStr(1, UCase(pRecordsetMovimenti), "ORDER BY ")
    orderBy = Mid$(pRecordsetMovimenti, iorderBy)
    
    DevEnvMDB.rscmdGiacenze.Source = Left$(pRecordsetMovimenti, iorderBy - 1) & " WHERE QtaIngresso>0 and Lotti.tipo='I' and Lotti.codCategoria ='" & pcodCategoria & "'" & _
        " and Lotti.codSottoCategoria ='" & pcodSottoCategoria & "'" & _
        " and Lotti.codCer ='" & pcodCer & "' " & orderBy
    
    DevEnvMDB.rscmdGiacenze.Open
    lvwDettaglio.ListItems.Clear
    
    
    Set keysLotti = creaCollectionChiaviLotto(pretElencoLotti)
    
    With DevEnvMDB.rscmdGiacenze
        Do While Not .EOF
            Set listItm = lvwDettaglio.ListItems.Add(, "|" & .Fields("id"), .Fields("nrGiorniInGiacenza"))
            listItm.ListSubItems.Add , .Fields("qtaResidua").Name, FormatNumber(.Fields("qtaResidua"), 2, , , vbFalse)
            listItm.ListSubItems.Add , .Fields("dataIns").Name, .Fields("dataIns")
            listItm.ListSubItems.Add , .Fields("DataPrevistaRitiro").Name, .Fields("DataPrevistaRitiro")
            listItm.ListSubItems.Add , .Fields("DataPrevistaRitiroNormativa").Name, .Fields("DataPrevistaRitiroNormativa")
            
            
            
            
            If (.Fields("DataPrevistaRitiro")) <= Date Or _
                .Fields("DataPrevistaRitiroNormativa") <= Date Then
                    coloraRiga listItm, vbRed
            End If
            
            If isLottiInKeys(keysLotti, .Fields("id")) Or keysLotti.Count = 0 Then
                listItm.Checked = True
                qty = qty + getNumber(.Fields("qtaResidua").Value)
            End If
            .MoveNext
        Loop
        txtPeso.Text = FormatNumber(qty, 2, , , vbFalse)
    End With


    lvwDettaglio.HideColumnHeaders = False
    AllargaLeColonneLVW lvwDettaglio, False
    'lvwDettaglio.HideColumnHeaders = True
    If (DevEnvMDB.rscmdGiacenze.state = adStateOpen) Then
        DevEnvMDB.rscmdGiacenze.Close
        DevEnvMDB.rscmdGiacenze.Source = pRecordsetMovimenti
    End If

End Sub

Private Sub lvwDettaglio_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim qty As Double
    Dim qty1 As Double
    
    If txtPeso.Text <> "" Then qty = getNumber(txtPeso.Text)
    
    If Item.ListSubItems("qtaResidua").Text <> "" Then qty1 = getNumber(Item.ListSubItems("qtaResidua").Text)
    
    txtPeso = FormatNumber(qty + IIf(Item.Checked = True, qty1, -qty1), 2, , , vbFalse)
    
End Sub

Private Function creaCollectionChiaviLotto(keys As String) As Collection
    Dim c As New Collection
    Dim keysLotti() As String
    Dim i As Long
    
    keysLotti = Split(keys, ",")
    For i = LBound(keysLotti) To UBound(keysLotti)
        c.Add keysLotti(i), "|" & keysLotti(i)
    Next
    Set creaCollectionChiaviLotto = c
End Function


Private Function isLottiInKeys(keysLotti As Collection, key As String) As Boolean
    Dim s As String
    isLottiInKeys = False
    
    On Error Resume Next
    s = keysLotti("|" & key)
    isLottiInKeys = s <> vbNullString
    
End Function

