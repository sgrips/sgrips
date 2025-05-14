Attribute VB_Name = "GestionePorta"
Option Explicit



Function ComCollegata() As Boolean
    'Verifica se il cavo NULL modem e collegato alla seriale
    ComCollegata = False
    
    If gDRS_CD_RequestMessage = 1 Then
        With FrmMonitoraggio.MSCPortaCom
           If .DSRHolding And .CDHolding Then
                ComCollegata = True
                FrmMonitoraggio.Label1 = "In fase di smistamento"
                frmServizi.CambiaStato gCOM(0), True
           Else
                FrmMonitoraggio.Label1 = "Cavo non collegato su porta COM " & gCOM(0)
                frmServizi.CambiaStato gCOM(0), False
           End If
        
           
        End With
    Else
        ComCollegata = True
        frmServizi.CambiaStato gCOM(0), True
        FrmMonitoraggio.Label1 = "In fase di smistamento"
    End If
End Function

Function InitCom(objCOM As MSComm) As Boolean
    'Inizializzazione porta
    
    On Error GoTo InitComERR
    InitCom = False
    With objCOM
        If .PortOpen = False Then
            .PortOpen = True
            .InputLen = 1
            .InputMode = comInputModeText
            InitCom = True
        End If
    End With
    Exit Function
InitComERR:
    MsgBox "COM" & objCOM.CommPort & " - " & Error$, vbCritical + vbOKOnly, App.Title
End Function

Sub SettingCom(objCOM As MSComm)
    'Settaggio porta
    With objCOM
        .CommPort = gCOM(0)
        .Settings = gBaudeRateCOM(0) & "," & Left$(gBitParita(0), 1) & "," & gDataBits(0) & "," & gBitDiStop(0)
        .InBufferSize = 10240
    End With
End Sub

Static Function LetturaPorta() As String
    'Lettura dal buffer del MSG
    Dim smsg As String
    Dim Esci As Boolean
    Dim sChecksum As Long
    Dim sChar As String
    Dim l As Long
    Dim bStx As Boolean, bEtx As Boolean, bCr As Boolean, bErr As Boolean
    Dim c1 As Long, c2 As Long
    Dim OldBuffer As Long
    
    LetturaPorta = vbNullString
    smsg = vbNullString
    Esci = False
    bErr = True
    
    
    With FrmMonitoraggio.MSCPortaCom
        Do While Not Esci And ComCollegata
            
            If .InBufferCount = 0 And gEsci And Timer - gTimeOut > 5 Then
                Exit Do
            End If
            
            If bErr Then
                Esci = False: bStx = False: bEtx = False
                bErr = False
            End If
            
            If .InBufferCount > 0 Then
                sChar = .Input
                If sChar = STX Then
                    gEsci = False
                    smsg = vbNullString
                    Esci = False: bStx = True
                    bEtx = False: sChecksum = 0
                    c1 = 0: c2 = 0
                    gTimeOut = Timer
                ElseIf sChar = ETX Then
                    If bStx And Not bEtx Then
                        bEtx = True
                    Else
                        bErr = True
                    End If
                ElseIf sChar = CR Then
                    If bStx And bEtx Then
                        Esci = True
                    Else
                        bErr = True
                    End If
                Else
                    If Not bStx Then
                    ElseIf Not bEtx Then
                        smsg = smsg & sChar
                        sChecksum = sChecksum Xor c1
                        c1 = c2: c2 = Asc(sChar)
                    Else
                        bErr = True
                    End If
                End If
                
                If Esci And Not Right$("0" & Hex(sChecksum), 2) = Chr(c1) & Chr(c2) Then
                    Esci = False: bErr = True
                End If
            Else
                DoEvents
            End If
        Loop
    End With
    LetturaPorta = smsg
End Function


Public Static Sub InvioMsg(sBaia As String, err As String, sNumeroProgPOSISORTER As String, SalvaSUDB As Boolean, Optional MSG As String)
    Dim crc As Long, i As Long
    Dim num As Long
    
    If MSG = vbNullString Then
        'Invio messaggio di risposta al POSISORTER
        MSG = "11|" & sNumeroProgPOSISORTER & "|" & sBaia & "|" & err & "|"
        
        crc = 0
        For i = 1 To Len(MSG)
            crc = crc Xor Asc(Mid$(MSG, i, 1))
        Next
        MSG = MSG & Right$("0" & Hex(crc), 2)
    End If
    
    
    With FrmMonitoraggio.MSCPortaCom
'        .Output = STX & MSG & ETX & CR
        If SalvaSUDB Then SalvaDestinazionePacco sNumeroProgPOSISORTER, MSG
    End With
    
    'num = num + 1
    'ScriviLog "reply" & Trim$(Str$(num)), Msg
End Sub


Sub CaricaComboBaudeRate(Cmb As Variant, IdCMB As Long, idPartenza As Long)
    Dim i As Long, ii As Long
    Dim MAXLen As Long
    Dim Trovato As Boolean
    For ii = IdCMB To Cmb.Count - 1
        MAXLen = 0
        With Cmb(ii)
            .AddItem "110"
            .AddItem "300"
            .AddItem "1200"
            .AddItem "2400"
            .AddItem "9600"
            .AddItem "14400"
            .AddItem "19200"
            .AddItem "28800"
            
            For i = 0 To .ListCount - 1
                If .List(i) = gBaudeRateCOM(ii + idPartenza) Then
                    .ListIndex = i
                    Trovato = True
                ElseIf Not Trovato Then .ListIndex = 0
                End If

                If .Parent.TextWidth(.List(i)) > MAXLen And idPartenza <> 0 Then MAXLen = .Parent.TextWidth(.List(i)): .Width = MAXLen + 360
            Next
        End With
    Next
End Sub

Sub CaricaComboBitParita(Cmb As Variant, IdCMB As Long, idPartenza As Long)
    Dim i As Long, ii As Long
    Dim MAXLen As Long
    Dim Trovato As Boolean
    For ii = IdCMB To Cmb.Count - 1
        MAXLen = 0
        With Cmb(ii)
            .AddItem "Even"
            .AddItem "Mark"
            .AddItem "None"
            .AddItem "Odd"
            .AddItem "Space"
            
            For i = 0 To .ListCount - 1
                If .List(i) = gBitParita(ii + idPartenza) Then
                    .ListIndex = i
                    Trovato = True
                ElseIf Not Trovato Then .ListIndex = 0
                End If

                If .Parent.TextWidth(.List(i)) > MAXLen And idPartenza <> 0 Then MAXLen = .Parent.TextWidth(.List(i)): .Width = MAXLen + 360
            Next
        End With
    Next
End Sub

Sub CaricaComboNumeroPorte(Cmb As Variant, IdCMB As Long, idPartenza As Long)
    Dim i As Long
    Dim ii As Long
    Dim MAXLen As Long
    Dim Trovato As Boolean
    
    For ii = IdCMB To Cmb.Count - 1
        MAXLen = 0
        With Cmb(ii)
            .AddItem "COM1"
            .AddItem "COM2"
            .AddItem "COM3"
            .AddItem "COM4"
            .AddItem "COM5"
            .AddItem "COM6"
            .AddItem "COM7"
            .AddItem "COM8"
            .AddItem "COM9"
            .AddItem "COM10"
            .AddItem "COM11"
            .AddItem "COM12"
            .AddItem "COM13"
            For i = 0 To .ListCount - 1
                If .List(i) = "COM" & gCOM(idPartenza + ii) Then
                    .ListIndex = i
                    Trovato = True
                ElseIf Not Trovato Then .ListIndex = 0
                End If
                If .Parent.TextWidth(.List(i)) > MAXLen And idPartenza <> 0 Then MAXLen = .Parent.TextWidth(.List(i)): .Width = MAXLen + 360
            Next
        End With
    Next
    
    
End Sub


Sub CaricaComboBitStop(Cmb As Variant, IdCMB As Long, idPartenza As Long)
    Dim i As Long, ii As Long
    Dim MAXLen As Long
    Dim Trovato As Boolean
    For ii = IdCMB To Cmb.Count - 1
        MAXLen = 0
        With Cmb(ii)
            .AddItem "1"
            .AddItem "1.5"
            .AddItem "2"
            
            For i = 0 To .ListCount - 1
                If .List(i) = gBitDiStop(idPartenza + ii) Then
                    .ListIndex = i
                    Trovato = True
                ElseIf Not Trovato Then .ListIndex = 0
                End If
                If .Parent.TextWidth(.List(i)) > MAXLen And idPartenza <> 0 Then MAXLen = .Parent.TextWidth(.List(i)): .Width = MAXLen + 360
            Next
        End With
    Next
End Sub

Sub CaricaComboDataBits(Cmb As Variant, IdCMB As Long, idPartenza As Long)
    Dim i As Long, ii As Long
    Dim MAXLen As Long
    Dim Trovato As Boolean

    For ii = IdCMB To Cmb.Count - 1
        With Cmb(ii)
            .AddItem "4"
            .AddItem "5"
            .AddItem "6"
            .AddItem "7"
            .AddItem "8"
            For i = 0 To .ListCount - 1
                If .List(i) = gDataBits(idPartenza + ii) Then
                    .ListIndex = i
                    Trovato = True
                ElseIf Not Trovato Then .ListIndex = 0
                End If

                If .Parent.TextWidth(.List(i)) > MAXLen And idPartenza <> 0 Then MAXLen = .Parent.TextWidth(.List(i)): .Width = MAXLen + 550
            Next
            
        End With
    Next
End Sub

Sub CaricaComboTipo(Cmb As Variant, IdCMB As Long)
    Dim i As Long, ii As Long
    Dim MAXLen As Long
    Dim Trovato As Boolean
    
    For ii = IdCMB To Cmb.Count - 1
        MAXLen = 0
        With Cmb(ii)
            .AddItem "ZEBRA8"
            
            For i = 0 To .ListCount - 1
                If .List(i) = gTipoStampanti(ii) Then
                    .ListIndex = i
                    Trovato = True
                ElseIf Not Trovato Then .ListIndex = 0
                End If

                If .Parent.TextWidth(.List(i)) > MAXLen Then MAXLen = .Parent.TextWidth(.List(i)): .Width = MAXLen + 360
            Next
        End With
    Next
End Sub


Sub CaricaComboPosizione(Cmb As Variant, IdCMB As Long)
    Dim i As Long, ii As Long
    Dim MAXLen As Long
    Dim rs As rdoResultset
    Dim Trovato As Boolean
    For ii = IdCMB To Cmb.Count - 1
        MAXLen = 0
        With Cmb(ii)
        
            Set rs = rdoACC.OpenResultset("Select * from PosizioniPistole order by idPistola")
            
            Do While Not rs.EOF
                .AddItem rs("idPistola")
                rs.MoveNext
            Loop
            
            For i = 0 To .ListCount - 1
                If .List(i) = gPosizionePistola(ii) Then
                    .ListIndex = i
                    Trovato = True
                ElseIf Not Trovato Then .ListIndex = 0
                End If

                If .Parent.TextWidth(.List(i)) > MAXLen Then MAXLen = .Parent.TextWidth(.List(i)): .Width = MAXLen + 360
            Next
        End With
    Next
    
    rs.Close
    Set rs = Nothing
End Sub

Static Function ScritturaPorta(ByVal s As String, ByVal MSC As MSComm) As String
    'Lettura dal buffer del MSG
    '~HS stato stampante
    MSC.Output = s
    
End Function

