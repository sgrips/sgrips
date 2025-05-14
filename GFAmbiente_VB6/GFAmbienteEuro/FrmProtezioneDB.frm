VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProtezioneDB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imposta Password DB"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   4890
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Chiudi"
      Default         =   -1  'True
      Height          =   375
      Left            =   3780
      TabIndex        =   8
      Top             =   3780
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Salva"
      Height          =   375
      Left            =   2580
      TabIndex        =   7
      Top             =   3780
      Width           =   1035
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2895
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Testo Cryptato"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Testo Decryptato"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      Left            =   3660
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3300
      Width           =   1095
   End
   Begin VB.TextBox txtText 
      Height          =   315
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3300
      Width           =   3375
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "Cancella"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   3780
      Width           =   1035
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "Aggiungi"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3780
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "Password "
      Height          =   195
      Left            =   3660
      TabIndex        =   5
      Top             =   3060
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Testo da cryptare"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   3060
      Width           =   1695
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuImpostaPassword 
         Caption         =   "&Imposta Password"
      End
   End
End
Attribute VB_Name = "frmProtezioneDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private itemSelected As ListItem

'This is a test of a rudimentary encryption/decryption
'algorithm.

'Note: It is possible for the encrypted string to include Chr$(0).
'Visual Basic handles this just fine but Windows (and
'therefore the TextBox control) uses Chr$(0) to signify the end
'of the string. Therefore, programs such as this one that store
'the encrypted data in a text box may end up truncating the data.


'First, add a textbox, called txtText, with its MultiLine property set to true. Then, add another textbox called txtPassword, and two command buttons named cmdEncrypt, and cmdDecrypt. Finally, add the code below.

'This program may be distributed on the condition that it is
'distributed in full and unchanged, and that no fee is charged for
'such distribution with the exception of reasonable shipping and media
'charged. In addition, the code in this program may be incorporated
'into your own programs and the resulting programs may be distributed
'without payment of royalties.
'
'This example program was provided by:
' SoftCircuits Programming
' http://www.softcircuits.com
' P.O. Box 16262
' Irvine, CA 92623


'Set to True to make the password case-sensitive
#Const CASE_SENSITIVE_PASSWORD = False

Private Sub cmdEncrypt_Click()
    Dim Item As ListItem
    ' You can encrypt twice for extra security
    txtText.Tag = txtText
    txtText.Tag = EncryptText((txtText.Tag), txtPassword.Tag)
    txtText.Tag = EncryptText((txtText.Tag), txtPassword.Tag)
    
    Set Item = lvw.ListItems.Add(, , txtText.Tag)
    Item.ListSubItems.Add , , txtText
End Sub

Private Sub cmdDecrypt_Click()
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    lvw.ListItems.Remove (lvw.SelectedItem.Index)
End Sub

'Encrypt text
Private Function EncryptText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String

#If Not CASE_SENSITIVE_PASSWORD Then

    'Convert password to upper case
    'if not case-sensitive
    strPwd = UCase$(strPwd)

#End If

    'Encrypt string
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c + Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    EncryptText = strBuff
End Function

'Decrypt text encrypted with EncryptText
Private Function DecryptText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String

#If Not CASE_SENSITIVE_PASSWORD Then

    'Convert password to upper case
    'if not case-sensitive
    strPwd = UCase$(strPwd)

#End If

    'Decrypt string
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c - Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    DecryptText = strBuff
End Function



Private Sub cmdSave_Click()
    Dim i As Integer
    Dim l As Long
    
    WritePrivateProfileString "PASSWORD", vbNullString, vbNullString, App.Path & "\config.ini"
    For i = 1 To lvw.ListItems.Count
        l = WritePrivateProfileString("PASSWORD", "password" & CStr(i - 1), lvw.ListItems(i).Text, App.Path & "\config.ini")
    Next

End Sub

Private Sub Form_Load()
    'carico lista con testo cryptato salvato su config.ini
    Dim l As Long
    Dim returnString As String
    Dim Item As ListItem
    Dim i As Integer
    Dim password As String

    txtPassword.Tag = "gfambiente"
    txtPassword.Text = EncryptText(EncryptText("gfambiente", txtPassword.Tag), txtPassword.Tag)
    i = 0
    Do
        returnString = Space(255)
        l = GetPrivateProfileString("PASSWORD", "password" & CStr(i), "", returnString, Len(returnString), App.Path & "\config.ini")
        returnString = Left(returnString, l)
        i = i + 1
        If returnString <> "" Then
            Set Item = lvw.ListItems.Add(, , returnString)
            password = DecryptText(DecryptText(returnString, txtPassword.Tag), txtPassword.Tag)
            Item.ListSubItems.Add , password, password
            If password = gPasswordDB Then
                Item.ForeColor = vbBlue
                Item.ListSubItems(1).ForeColor = vbBlue
                Set itemSelected = Item
            End If
        End If
    Loop While returnString <> ""
End Sub


Private Sub lvw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If vbRightButton = Button Then
        Me.PopupMenu mnuPopup
    End If
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub mnuImpostaPassword_Click()
    'imposta la password sul DB
    
    ChiudiConnessione
    ApriConnessione True
    
    gDBC.Execute "ALTER Database Password " & lvw.SelectedItem.ListSubItems(1) & " " & gPasswordDB
    
    itemSelected.ForeColor = vbBlack
    itemSelected.ListSubItems(1).ForeColor = vbBlack
    
    lvw.SelectedItem.ForeColor = vbBlue
    lvw.SelectedItem.ListSubItems(1).ForeColor = vbBlue
    
    Set itemSelected = lvw.SelectedItem
    
    ChiudiConnessione
    ApriConnessione

    cmdSave_Click
End Sub
