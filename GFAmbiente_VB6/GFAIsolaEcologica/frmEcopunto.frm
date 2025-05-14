VERSION 5.00
Begin VB.Form frmEcopunto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importazione Dati Ecopunto"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Stampa Anomalie"
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
      Left            =   1560
      TabIndex        =   8
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   240
      Top             =   3480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Avvia Lettura"
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
      Left            =   3360
      TabIndex        =   1
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
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
      Left            =   5160
      TabIndex        =   0
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Movimenti Creati:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   3000
      TabIndex        =   7
      Top             =   2880
      Width           =   2460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   4
      Left            =   5460
      TabIndex        =   6
      Top             =   2880
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   3
      Left            =   1920
      TabIndex        =   5
      Top             =   2880
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Num.Pesate:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Lettura Dati in corso ..... "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   1560
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   "Attesa Dati da Ecopunto, collegare la pesa alla seriale [COM], premere il pulsante (Avvio Lettura) per avviare acquisizione."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   5190
   End
End
Attribute VB_Name = "frmEcopunto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mStopLettura As Boolean
Dim comIDpesa As Integer


Private Sub Command1_Click(Index As Integer)
    Dim char As String
    Select Case Index
        Case 0
            Unload Me
        Case 1
        
            If Command1(Index).Caption = "Avvia Lettura" Then
                Label1(0).Visible = False
                Label1(1).Visible = True
                Label1(2).Visible = True
                Label1(3).Visible = True
                Label1(4).Visible = True
                Label1(5).Visible = True
                Timer1.enabled = True
                Command1(Index).Caption = "Termina Lettura"
                mStopLettura = False
                If Not pesa(comIDpesa).COMControl.PortOpen Then
                     pesa(comIDpesa).COMControl.PortOpen = True
                End If
                Label1(1).Caption = "Richiesta Dati in corso ....."
                pesa(comIDpesa).SendMessage Chr$(&H5&)
                
                mStopLettura = False
                Command1(0).enabled = False
                Label1(1).Caption = "Lettura Dati in corso ....."
                If letturaMovimenti(comIDpesa) Then
                    generaMovimentiIsola
                    Label1(1).Caption = "Importazione avvenuta con Successo"
                    Label1(1).Visible = True
                    Command1(Index).Caption = "Avvia Lettura"
                    DoEvents
                End If
                Timer1.enabled = False
                Command1(0).enabled = True
                
                Command1(2).enabled = Label1(3).Caption <> Label1(4).Caption
            Else
                Timer1.enabled = False
                Command1(0).enabled = True
                Command1(Index).Caption = "Avvia Lettura"
                Label1(1).Caption = "Lettura Dati Arrestata ....."
                Label1(2).Visible = False
                Label1(3).Visible = False
                Label1(4).Visible = False
                Label1(5).Visible = False
                
                mStopLettura = True
            End If
        Case 2
            If DevEnvMDB.rstblPesateEcoPunto.state = adStateOpen Then
                DevEnvMDB.rstblPesateEcoPunto.Close
            End If
            DevEnvMDB.rstblPesateEcoPunto.Open
            
            ReportAnomalieEcoPunto.Show vbModal
            
    End Select
End Sub

Private Sub generaMovimentiIsola()
    Dim codUtente As Long
    Dim rs As ADODB.Recordset
    Dim rsImmobili As ADODB.Recordset
    Dim rsCodCer As ADODB.Recordset
    Dim codCategoria As String, _
        codSottoCategoria As String, _
        codCer As String, _
        CodEcoPunto As String, _
        codSubCer As String
    Dim codComune As String
    Dim movId As Long
    
    movId = 0
    'vengono lette tutte le pesata acquisite con la pesa ECOPUNTO, per ogni pesata viene
    'creato un movimento sull'immobile di default del codice utente rilevato.
    Set rs = gDBC.Execute("select * from MovimentiEcoPunto")
    
    codComune = GetCurrentComune!codistat
    
    Do While Not rs.EOF
        codUtente = getCodiceUtenteDaBadge(rs("CodBadge"))
        If rs!peso = 0 Then
            UpdateInfoMov rs!id, "Peso non valido deve essere > 0 "
        Else
        If codUtente <> 0 Then
            Set rsImmobili = gDBC.Execute("SELECT Immobili.* " & _
                "FROM immobili WHERE immobili.FlgUtenzaDomesticaResidente= true and  DATE() between DataInizioValidita and DataFineValidita and   immobili.Tipo<>'Z' AND immobili.[codUtente]=" & codUtente)
    
            If Not rsImmobili.EOF Then
            
                'dal codice link fra Ecopunti e codcer ricavo il rifiuto
                Set rsCodCer = gDBC.Execute("SELECT *  " & _
                                "FROM PeseEcoPuntoConCodCer " & _
                                "WHERE codLink = '" & rs("CodRifiuto").value & "'")
                
                If (Not rsCodCer.EOF And gflgCodPesaECOPunto = 1) Or (rsCodCer.EOF And gflgCodPesaECOPunto = 0) Then
                    If gflgCodPesaECOPunto = 1 Then
                        CodEcoPunto = left$(rs("CodRifiuto").value, 2)
                        codCategoria = rsCodCer("CodCategoria")
                        codSottoCategoria = rsCodCer("codSottoCategoria")
                        codCer = rsCodCer("codCer")
                        codSubCer = rsCodCer("codSubCer")
                    Else
                        'todo ricontrollare cosa fare
                        CodEcoPunto = "00"
                        codCategoria = left$(rs("CodRifiuto"), 2)
                        codSottoCategoria = Mid$(rs("CodRifiuto"), 3, 2)
                        codCer = Right$(rs("CodRifiuto"), 2)
                    End If

                    Set rsCodCer = gDBC.Execute("SELECT *  " & _
                                "FROM CodiciCerComunali " & _
                                "WHERE FlgAbilitatoInRicerca=True  and CodCategoria='" & codCategoria & "' and codSottoCategoria='" & codSottoCategoria & "' and CodCer='" & codCer & "' and CodSubCer='" & codSubCer & "'")
            
                    
                    If Not rsCodCer.EOF Then
                        'crea movimento
                        creaMovimento getNumeratore(MOVIMENTI), CLng(rsImmobili("CodImmobile").value), codUtente, codComune, codCategoria, codSottoCategoria, codCer, codSubCer, _
                            Format$(Now, "yyyy/mm/dd hh:nn:ss"), CDbl(rs("Peso").value), "0", False, frmLogin.codUser, _
                            CDbl(Round(rsCodCer("CorrispettivoBonus").value * rs("Peso").value, 4)), _
                            CDbl(Round(rsCodCer("BonusPunti").value * rs("Peso").value, 4)), _
                            CStr(rsCodCer("UnitaDiMisura").value), "", _
                            "", "", False, "", 0, gCodCausaleMovEcoPuntoIN, gflgMovECOPuntoGeneraGiacenza = 0, CodEcoPunto, 0, vbNullString
                        
                        movId = movId + 1
                        Label1(4).Caption = movId
                    Else
                        UpdateInfoMov rs!id, "Manca CodiceCer " & codCategoria & codSottoCategoria & codCer & ", o non abiliatato in ricerca nella tabella CodiCer "
                    End If
                Else
                    UpdateInfoMov rs!id, "Manca associazione con codice codLink " & rs("CodRifiuto").value & " in tabella PeseEcoPuntoConCodCer "
                End If
            Else
                UpdateInfoMov rs!id, "Manca immobile oppure non impostato come utenza domestica residenziale per utente :" & codUtente & " nella tabella Immobili "
            End If
        Else
            UpdateInfoMov rs!id, "Non trovato Utente associato al badge :" & rs("CodBadge") & " nella tabella Badge "

        End If
        End If
        rs.MoveNext
    Loop
    
    
End Sub

Private Sub UpdateInfoMov(idMov As Long, info As String)
    Dim cmd As New ADODB.Command
    
    cmd.CommandText = "update MovimentiEcoPunto set Errore = ? where id=?"
    Set cmd.ActiveConnection = gDBC
    cmd.Parameters.Append cmd.CreateParameter(, adVarChar, adParamInput, 255)
    cmd.Parameters.Append cmd.CreateParameter(, adInteger, adParamInput)
    cmd.Parameters(0).value = info
    cmd.Parameters(1).value = idMov
    cmd.Execute
End Sub


Private Sub Form_Load()

    Label1(1).Visible = False
    Label1(2).Visible = False
    Label1(3).Visible = False
    Label1(4).Visible = False
    Label1(5).Visible = False
    
    'cerco la pesa per ecopunto
    g_rsPese.Filter = ""
    g_rsPese.Filter = "flgEcopunto=true and active=true"
    
     
    If g_rsPese.EOF Then
        g_rsPese.Filter = ""
        MsgBox "Attenzione non trovata pesa Ecopunto, verficare che sia stata attivata", vbCritical + vbOKOnly, "ECOPUNTO"
        On Error Resume Next
        Unload Me
        Exit Sub
    End If
    
    g_rsPese.Filter = ""
    
    comIDpesa = g_rsPese!ComID
    Label1(0).Caption = Replace(Label1(0).Caption, "[COM]", "COM" & g_rsPese!ComID)
    
    
    Dim rs As ADODB.Recordset
    Set rs = gDBC.Execute("select * from  MovimentiEcoPunto where Errore <>'' and not errore is null")
    Command1(2).enabled = Not rs.EOF
    rs.Close
    
End Sub


Private Function letturaMovimenti(Index As Integer)
    Dim messaggio As String
    Dim msgPesa As messaggioPesa
    Dim cmd As Command
    Dim numRecord As Long
    
    gDBC.BeginTrans
    
    letturaMovimenti = False
    Set cmd = CreateObject("ADODB.Command")
    cmd.CommandType = adCmdText
    cmd.CommandText = "Insert into MovimentiEcoPunto (data,codBadge,codRifiuto,peso) values (?,?,?,?)"
    Set cmd.ActiveConnection = gDBC
    cmd.Parameters.Refresh
    
    gDBC.Execute "delete * from MovimentiEcoPunto"
    
    
    Do While Not mStopLettura
'        messaggio = "$11/04/200817:22:30   1010000002017200102    4.180 kg3D" & vbCr
'        messaggio = "$19/05/200814:58:28   1010002998013200139    1.060 kg35"
        Select Case pesa(Index).ReadMessage(messaggio)
            Case ErroreRicevuto
            Case InLettura
            Case LetturaTerminata
            
                If messaggio = Chr$(&H4&) Then
                    letturaMovimenti = True
                    mStopLettura = True
                    
                Else
                    pesa(comIDpesa).SendMessage Chr$(&H6&)

                    msgPesa = getMessaggioPesataEcoPunto(messaggio)
    
                    cmd.Parameters(0).value = msgPesa.Data
                    cmd.Parameters(1).value = Trim$(msgPesa.codBadge)
                    cmd.Parameters(2).value = Trim$(msgPesa.CodRifiuto)
                    cmd.Parameters(3).value = msgPesa.PesoNetto
                    cmd.Execute
                    
                    numRecord = numRecord + 1
                    Label1(3).Caption = numRecord
                End If
         End Select
         DoEvents
    Loop
    
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    
    
    gDBC.CommitTrans
    
    Exit Function
letturaMovimentiERR:
    gDBC.RollbackTrans
End Function




Private Sub Form_Unload(Cancel As Integer)
    pesa(comIDpesa).COMControl.PortOpen = False
    
    mStopLettura = True
End Sub

Private Sub Timer1_Timer()
    Label1(1).Visible = Not Label1(1).Visible
    DoEvents
End Sub
