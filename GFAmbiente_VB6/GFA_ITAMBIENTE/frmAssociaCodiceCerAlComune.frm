VERSION 5.00
Begin VB.Form frmAssociaCodiceCerAlComune 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   13185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   6120
      Picture         =   "frmAssociaCodiceCerAlComune.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   6120
      Picture         =   "frmAssociaCodiceCerAlComune.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   6120
      TabIndex        =   3
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   6120
      TabIndex        =   2
      Top             =   2640
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cer Associato"
      Height          =   7935
      Index           =   1
      Left            =   7080
      TabIndex        =   1
      Top             =   120
      Width           =   5895
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7260
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cer Da Associare"
      Height          =   7935
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7260
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmAssociaCodiceCerAlComune"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
    Dim i As Integer
    Dim cmdq As ADODB.Command
    Select Case Index
        Case 0
            Unload Me
        Case 1
            'associa
            List1(1).AddItem List1(0).List(List1(0).ListIndex) & "*"
        Case 2
            'disassocia
        Case 3
            'salva
            Set cmdq = CreateObject("adodb.command")
            Set cmdq.ActiveConnection = gDBC
            cmdq.CommandText = "INSERT INTO CodiciCerComunali ( CodComune, CodCategoria, CodSottoCategoria, CodCer, " & _
                              "CodSubCer, Descrizione, UnitaDiMisura, flgIngombrante, FlgAbilitatoInRicerca, FlgBonus, " & _
                              "CorrispettivoBonus, flgBonusAPunti, BonusPunti, nomeImmagine ) " & _
                              "values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                              
           cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 6)
           cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
           cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
           cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
           cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
           cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 255)
           cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
           cmdq.Parameters.Append cmdq.CreateParameter(, adBoolean)
           cmdq.Parameters.Append cmdq.CreateParameter(, adBoolean)
           cmdq.Parameters.Append cmdq.CreateParameter(, adBoolean)
           cmdq.Parameters.Append cmdq.CreateParameter(, adDecimal)
           cmdq.Parameters.Append cmdq.CreateParameter(, adBoolean)
           cmdq.Parameters.Append cmdq.CreateParameter(, adDecimal)
           cmdq.Parameters.Append cmdq.CreateParameter(, adBoolean)
           cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 30)
           
           For i = 0 To List1(1).ListCount - 1
                Dim value As String
                
                value = List1(1).List(i)
                If InStr(value, "*") > 0 Then
                    cmdq.Parameters(0).value = GetCurrentComune!codistat
                    cmdq.Parameters(1).value = left(value, 2)
                    cmdq.Parameters(2).value = Mid(value, 3, 2)
                    cmdq.Parameters(3).value = Mid(value, 5, 2)
                    cmdq.Parameters(4).value = "00"
                    value = Mid(value, 10)
                    cmdq.Parameters(5).value = left(value, Len(value) - 1)
                    cmdq.Parameters(6).value = "KG"
                    cmdq.Parameters(7).value = False
                    cmdq.Parameters(8).value = False
                    cmdq.Parameters(9).value = False
                    cmdq.Parameters(10).value = 0
                    cmdq.Parameters(11).value = False
                    cmdq.Parameters(12).value = 0
                    cmdq.Parameters(13).value = False
                    cmdq.Parameters(14).value = ""
                    
                    cmdq.Execute
                End If
                
           Next
           Set cmdq.ActiveConnection = Nothing
           Unload Me
    End Select
    
End Sub


Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    On Error GoTo Form_LoadERR
    
    Set rs = CreateObject("adodb.recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    
    Set rs.ActiveConnection = gDBC
    rs.Source = "select * from CodiciCer where CodCategoria+CodSottoCategoria+CodCer " & _
                "not in (select CodCategoria+CodSottoCategoria+CodCer from CodiciCerComunali where codComune='" & GetCurrentComune!codistat & "')" & _
                "order by CodCategoria,CodSottoCategoria,CodCer,descrizione"
    rs.Open
    Set rs.ActiveConnection = Nothing
    
    Do While Not rs.EOF
        List1(0).AddItem (rs!codCategoria + rs!codSottoCategoria + rs!codCer & " - " & rs!Descrizione)
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs.ActiveConnection = gDBC
    rs.Source = "select * from CodiciCerComunali where codComune='" & GetCurrentComune!codistat & "'" & _
                "order by CodCategoria,CodSottoCategoria,CodCer,descrizione"
    rs.Open
    Set rs.ActiveConnection = Nothing
    
    Do While Not rs.EOF
        List1(1).AddItem (rs!codCategoria + rs!codSottoCategoria + rs!codCer & " - " & rs!Descrizione)
        rs.MoveNext
    Loop
    
Form_LoadERR:
    On Error Resume Next
    If rs.state = adStateOpen Then rs.Close
    
End Sub
