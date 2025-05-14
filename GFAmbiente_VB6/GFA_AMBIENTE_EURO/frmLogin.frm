VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selezionare il Comune "
   ClientHeight    =   4044
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5328
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4044
   ScaleWidth      =   5328
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtGiorniProva 
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Text            =   "30"
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   6
      Text            =   "description"
      Top             =   1560
      Width           =   3795
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Text            =   "codistat"
      Top             =   1560
      Width           =   1275
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   1605
      Left            =   60
      TabIndex        =   3
      Top             =   1860
      Width           =   5175
      _ExtentX        =   9123
      _ExtentY        =   2836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton BubbleButton1 
      Cancel          =   -1  'True
      Caption         =   "&Annulla"
      Default         =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   3945
      TabIndex        =   0
      Top             =   3615
      Width           =   1335
   End
   Begin VB.CommandButton BubbleButton1 
      Caption         =   "&Conferma"
      Height          =   375
      Index           =   0
      Left            =   2550
      TabIndex        =   1
      Top             =   3615
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Giorni di Prova"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3660
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ALL COMUNI ATTIVATA BENVENUTO SUPER UTENTE "
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1440
      Left            =   105
      TabIndex        =   4
      Top             =   1995
      Visible         =   0   'False
      Width           =   5130
   End
   Begin VB.Label Label1 
      Caption         =   "scelta comune"
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5070
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsComuni As ADODB.Recordset
Private mColElementChecked As New Collection
Private gGiorniScandenzaDemo As Integer


Private Sub BubbleButton1_Click(Index As Integer)
    Dim i As Integer
    gErrore = True
    Select Case Index
        Case 0
            'tolgo l'eventuale filtro
            txt(0).Text = vbNullString
            txt(1).Text = vbNullString
            gErrore = True
            If Not lbl.Visible Then
                For i = 1 To lvw.ListItems.Count
                     gErrore = False
                     If lvw.ListItems(i).Checked Then
                        If Not AttivaComune(lvw.ListItems(i).Text, lvw.ListItems(i).ListSubItems("Comune")) Then
                            gErrore = True
                            Exit For
                        End If
                    End If
                Next
            
            Else
                If AttivaComune(ALL_COMUNI, ALL_COMUNI) Then
                    gErrore = False
                End If
            End If
            
            If Not gErrore Then
                LetParamValue "1900", vbNullString, PARAM_GIORNIDEMO, False, txtGiorniProva
                gGiorniScandenzaDemo = txtGiorniProva
            End If
            
            ComuneGiaAttivato
            Unload Me
        Case 1
            Unload Me
    End Select
End Sub


Private Sub Form_Load()
        
    
    txtGiorniProva = gGiorniScandenzaDemo
    Screen.MousePointer = vbHourglass
    DoEvents
        
    Label1.Caption = "Selezionare uno o più comuni dalla lista sotto indicata. " & _
                     "Premere Conferma per rendere il comune attivo. Ora la procedura " & _
                     "utilizzera le informazioni relative al comune selezionato. Se si è commesso " & _
                     "un errore, utilizzare disattiva la funzione (Disattiva Comune), " & _
                     "ATTENZIONE questa funzione distruggerà tutti i dati inseriti relativi al comune."
    
    txt(0).Text = vbNullString
    txt(1).Text = vbNullString
    CaricaComuni
    Screen.MousePointer = vbDefault
    
    
End Sub

Private Sub CaricaComuni()
    Set rsComuni = CreateObject("ADODB.Recordset")
    With rsComuni
        .LockType = adLockBatchOptimistic
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        Set .ActiveConnection = gDBC
        .Source = "Select comune,CODISTAT from comuni order by comune"
        .Open
        Set .ActiveConnection = Nothing

        CaricaListView lvw
    End With
End Sub

Private Sub CaricaListView(lvw As ListView)
    Dim Item As ListItem
    Dim obj As Variant
    lvw.ListItems.Clear
    lvw.ColumnHeaders.Add , , "Istat", lvw.Width * 0.2
    lvw.ColumnHeaders.Add , , "Comune", lvw.Width * 0.7
    
    Do While Not rsComuni.EOF
        Set Item = lvw.ListItems.Add(, "|" & rsComuni("CodIstat"), rsComuni("CodIstat"))
        On Error Resume Next
        obj = mColElementChecked("|" & rsComuni("CodIstat").Value)
        If Err = 0 Then Item.Checked = True
        Item.ListSubItems.Add , "Comune", rsComuni("Comune")
        rsComuni.MoveNext
    Loop

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
                    lbl.Visible = True
                    lvw.Visible = False
            End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not rsComuni Is Nothing Then
        If rsComuni.State = adStateOpen Then
            rsComuni.Close
        End If
    End If
End Sub

Private Sub lvw_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    If Item.Checked Then
        mColElementChecked.Add Item.Key, Item.Key
    Else
        mColElementChecked.Remove Item.Key
    End If
End Sub

Private Sub lvw_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub txt_Change(Index As Integer)
    Dim Filter As String
    
    If rsComuni Is Nothing Then Exit Sub
    rsComuni.Filter = vbNullString
    If txt(Index).Text <> vbNullString Then
        Select Case Index
            Case 0
                Filter = "codistat like '" & txt(0) & "%' and "
            Case 1
                Filter = Filter & "Comune like '" & txt(1) & "%' and "
        End Select
        
        Filter = Left$(Filter, Len(Filter) - Len(" and "))
        rsComuni.Filter = Filter
    End If
 
    
    CaricaListView lvw
    
End Sub

Private Sub txtGiorniProva_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub
