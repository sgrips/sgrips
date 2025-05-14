Attribute VB_Name = "Module1"
Option Explicit


Private mFileNumber As Long
Public gTimerLettura As Long
Public mDEBUG As Boolean

'tipi log
Enum TIPO_LOG
    TPL_ERRORE = 0
    TPL_INFORAMITIVO = 1
End Enum


Sub main()
    
End Sub


Public Sub TestMessage(tes As String)
    Dim c As COMConnection
    Dim mCOMconnections As New COMConnections

    
    mDEBUG = True
    
    Set c = mCOMconnections.Add("PIPPO")
    

    Form1.MSComm1.PortOpen = True
    Set c.COMControl = Form1.MSComm1
    c.STX = "ESC"
    c.ETX = "STX"
    c.StartChar = ""
    c.message.CarriageReturnON = False
    c.message.CheckSumON = False
    c.message.LineFeedON = False
    c.message.STARCHAR_ON = False
    c.message.STXETX_ON = True
    c.CDHolding = False
    c.DSRHolding = False
    
    c.GetNextChar tes
    Dim esci As Boolean
    Dim s As String
    
    esci = False
    Do While Not esci
        Select Case c.ReadMessage(s)
            Case ErroreRicevuto
                Debug.Print "Errore in lettura"
                esci = True
            Case InLettura
                
            Case LetturaTerminata
                Debug.Print s
                esci = True
        End Select
    Loop
    Unload Form1
End Sub


Public Sub ScriviLog(TipoLog As TIPO_LOG, sApplication As String, sAzione As String, sModulo As String)

    

End Sub
