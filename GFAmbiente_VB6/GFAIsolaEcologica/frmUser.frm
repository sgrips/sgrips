VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmUser 
   Caption         =   "Users"
   ClientHeight    =   8790
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   13005
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   2
      Left            =   3360
      ScaleHeight     =   3495
      ScaleWidth      =   8895
      TabIndex        =   15
      Top             =   3480
      Width           =   8895
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Cognome"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cod Utente"
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
         Left            =   240
         TabIndex        =   16
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6660
      Index           =   1
      Left            =   1200
      ScaleHeight     =   6630
      ScaleWidth      =   8505
      TabIndex        =   9
      Top             =   960
      Width           =   8535
      Begin VB.ComboBox cmbLivello 
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
         ItemData        =   "frmUser.frx":0000
         Left            =   2280
         List            =   "frmUser.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox txtidUser 
         BackColor       =   &H00E0E0E0&
         DataField       =   "CodUser"
         DataMember      =   "cmdUsers"
         DataSource      =   "DevEnvMDB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   330
         Width           =   3420
      End
      Begin VB.TextBox txtNome 
         DataField       =   "Nome"
         DataMember      =   "cmdUsers"
         DataSource      =   "DevEnvMDB"
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
         Left            =   1695
         MaxLength       =   50
         TabIndex        =   1
         Top             =   825
         Width           =   3420
      End
      Begin VB.TextBox txtCognome 
         DataField       =   "Cognome"
         DataMember      =   "cmdUsers"
         DataSource      =   "DevEnvMDB"
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
         Left            =   1695
         MaxLength       =   16
         TabIndex        =   2
         Top             =   1335
         Width           =   3420
      End
      Begin VB.TextBox txtPassword 
         DataField       =   "Password"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "00000000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
         DataMember      =   "cmdUsers"
         DataSource      =   "DevEnvMDB"
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
         Left            =   1695
         MaxLength       =   11
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1800
         Width           =   3420
      End
      Begin VB.TextBox txtLivello 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Livello"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "00000000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
         DataMember      =   "cmdUsers"
         DataSource      =   "DevEnvMDB"
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
         Left            =   1695
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   5
         Top             =   2280
         Width           =   540
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cod Utente:"
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
         Left            =   630
         TabIndex        =   14
         Top             =   375
         Width           =   1035
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
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
         Left            =   1095
         TabIndex        =   13
         Top             =   870
         Width           =   570
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cognome:"
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
         Left            =   780
         TabIndex        =   12
         Top             =   1380
         Width           =   885
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Password:"
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
         Index           =   3
         Left            =   750
         TabIndex        =   11
         Top             =   1875
         Width           =   915
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Livello:"
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
         Index           =   10
         Left            =   1050
         TabIndex        =   10
         Top             =   2325
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   0
      Left            =   240
      ScaleHeight     =   2895
      ScaleWidth      =   9735
      TabIndex        =   19
      Top             =   4080
      Width           =   9735
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmUser.frx":0032
         Height          =   2415
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4260
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         HeadLines       =   3
         RowHeight       =   19
         TabAction       =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataMember      =   "cmdUsers"
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "CodUser"
            Caption         =   "CodUser"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Nome"
            Caption         =   "Nome"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Cognome"
            Caption         =   "Cognome"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Password"
            Caption         =   "Password"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "\*\*\*\*\*\*\*\*\*\*\*\*"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Livello"
            Caption         =   "Livello"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1365,165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2085,166
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2085,166
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
               ColumnWidth     =   1365,165
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   734,74
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ricerca"
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
      Left            =   10320
      TabIndex        =   25
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Chiudi"
      Default         =   -1  'True
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
      Left            =   11640
      TabIndex        =   24
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Salva"
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
      Left            =   6360
      TabIndex        =   23
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Nuovo"
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
      Index           =   3
      Left            =   5040
      TabIndex        =   22
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancella"
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
      Index           =   4
      Left            =   3720
      TabIndex        =   21
      Top             =   7680
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2415
      Left            =   6000
      TabIndex        =   26
      Top             =   600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4260
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Lista"
            Object.ToolTipText     =   "Elenco dei trasportatori"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dettaglio"
            Object.ToolTipText     =   "Modifica dettaglio trasportatore"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Filtro"
            Object.ToolTipText     =   "Imposta un filtro per la ricerca"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private offsetX As Single
Private offsetY As Single

Const TXT_CODICEUSER = 0
Const TXT_COGNOME = 1
Const TXT_NOME = 2

Const CMD_CHIUDI = 0
Const CMD_RICERCA = 1
Const CMD_SALVA = 2
Const CMD_NUOVO = 3
Const CMD_CANCELLA = 4

Private pErrore As Boolean
Private pNoEvent As Boolean


Private Sub cmbLivello_Click()
    txtLivello.SetFocus
     txtLivello.Text = cmbLivello.ListIndex
     txtLivello.DataChanged = True
     
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim Des As String
    
    
    
    Select Case Index
        Case CMD_CHIUDI
            cancelCurrentOperation
            Unload Me
        Case CMD_RICERCA
            ricaricaLista
        Case CMD_SALVA
            txtidUser.Locked = True
            txtidUser.BackColor = 14737632
        
            pErrore = DevEnvMDB.rscmdUsers.EditMode <> adEditNone
        
            If DevEnvMDB.rscmdUsers.EditMode <> adEditNone Then
               If txtPassword.Tag <> txtPassword.Text Then
                    txtPassword.Text = getTextEncrypt(txtPassword.Text)
                    txtPassword.Tag = txtPassword.Text
               End If
               DevEnvMDB.rscmdUsers.UpdateBatch adAffectAllChapters
               
               DataGrid1.Refresh
               
               If gDisabilitaNotifiche = 0 Then MsgBox "Operazione eseguita con successo!!", vbExclamation + vbOKOnly, App.Title
               
               pErrore = False
            End If
            
        Case CMD_NUOVO
            
            TabStrip1.Tabs(2).Selected = True
            TabStrip1_Click
            
            DevEnvMDB.rscmdUsers.AddNew
            
            txtidUser.Locked = False
            txtidUser.BackColor = txtCognome.BackColor
            txtidUser.SetFocus
            
            txtLivello.Text = 2
            cmbLivello.Text = "User"
            

            
        Case CMD_CANCELLA
            If (MsgBox("Attenzione, se sicuro di voler cancellare l'utente: " & DevEnvMDB.rscmdUsers("Nome") & " " & DevEnvMDB.rscmdUsers("Cognome"), vbQuestion + vbYesNo) = vbYes) Then
                
                
                'cancella
                DevEnvMDB.rscmdUsers.Delete
                DevEnvMDB.rscmdUsers.UpdateBatch adAffectAllChapters
            End If
        
    End Select
End Sub


Private Function formatValue(Item As Column) As Variant
    On Error Resume Next
    Dim v As Variant
    
    v = Item.Value
    If Err.Number > 0 Then v = Item.Text
    
    If (Item.NumberFormat <> "") Then
        formatValue = toZero(v)
    Else
        formatValue = "=" & Chr$(34) & nullToBlank(v) & Chr$(34)
    End If
End Function


Private Sub ricaricaLista()
    Dim search As String
    
    Screen.MousePointer = vbHourglass
    
    search = vbNullString
    
    If txt(TXT_CODICEUSER) <> vbNullString Then
        search = " codUser like '" & txt(TXT_CODICEUSER) & "%'"
    End If
    

    If txt(TXT_COGNOME) <> vbNullString Then
        If search <> vbNullString Then search = search & " AND "
        search = search & " Cognome like '" & txt(TXT_COGNOME) & "%'"
    End If
    
    If txt(TXT_NOME) <> vbNullString Then
        If search <> vbNullString Then search = search & " AND "
        search = search & " Nome like '" & txt(TXT_NOME) & "%'"
    End If
    
    
    DevEnvMDB.rscmdUsers.Filter = search
    If search = "" Then
        DevEnvMDB.rscmdUsers.Filter = search & " Cognome <> '123'"
    End If
    
    
    
    TabStrip1.Tabs(1).Selected = True

    Screen.MousePointer = vbDefault

End Sub


Private Sub DataGrid1_DblClick()
    If DataGrid1.Row >= 0 Then
        TabStrip1.Tabs(2).Selected = True
        TabStrip1_Click
    End If
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cmd_Click (3)
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If DataGrid1.Row >= 0 Then
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo LoadERR
    
    If DevEnvMDB.rscmdUsers.state = adStateOpen Then
        DevEnvMDB.rscmdUsers.Close
        DevEnvMDB.cmdUsers
    End If
    
    Dim c As Control
    For Each c In Me.Controls
        If TypeName(c) = "TextBox" Then
            Dim t As TextBox
            Set t = c
            If t.DataMember = "cmdUsers" Then
                Set t.DataSource = Nothing
                Set t.DataSource = DevEnvMDB
            End If
        End If
    Next
    
    
    DataGrid1.AllowAddNew = False
    DataGrid1.AllowDelete = False
    DataGrid1.AllowUpdate = False
    
    
    Screen.MousePointer = vbHourglass
    
    DoEvents
    Picture1(0).Visible = False
    Picture1(1).Visible = False
    Picture1(2).Visible = False
    
    TabStrip1.Tabs(1).Selected = True
    
    TabStrip1_Click
    Screen.MousePointer = vbDefault
    
    Set DataGrid1.DataSource = DevEnvMDB
        
    DataGrid1_RowColChange -1, -1

    Exit Sub
LoadERR:
    MsgBox "Load frmUser:" & Err.Description, vbCritical, App.Title
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub

    offsetX = 10 * Screen.TwipsPerPixelX
    offsetY = 10 * Screen.TwipsPerPixelX
    
    resizeControl

    DoEvents
End Sub

Private Sub resizeControl()
    Dim i As Integer
    Dim posLeft As Long
    Dim posTop As Long
    
    
    On Error Resume Next
    posTop = Me.ScaleHeight - cmd(i).Height - offsetX
    posLeft = Me.ScaleWidth - cmd(i).width - offsetX
    
    For i = cmd.LBound To cmd.UBound
        If i > 0 Then posLeft = posLeft - offsetX - cmd(i).width
        
        cmd(i).Move posLeft, posTop
    Next
    
    TabStrip1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - cmd(0).Height - 2 * offsetX
    
    For i = Picture1.LBound To Picture1.UBound
        Picture1(i).Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
    Next
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If DevEnvMDB.rscmdUsers.state = adStateOpen Then
        DevEnvMDB.rscmdUsers.Close
    End If
    DevEnvMDB.rscmdUsers.Filter = vbNullString
    
    Set pfrmUser = Nothing
End Sub

Private Sub Picture1_Resize(Index As Integer)
    Select Case Index
        Case 0
            DataGrid1.Move 0, 0, Picture1(0).ScaleWidth, Picture1(0).ScaleHeight
        Case 1
    '        cmbLivello.Move txtLivello.left, txtLivello.top, txtLivello.width
        Case 2
            
    End Select
    
    
End Sub

Private Sub cancelCurrentOperation()
    If DevEnvMDB.rscmdUsers.EOF Or DevEnvMDB.rscmdUsers.BOF Then Exit Sub

    If DevEnvMDB.rscmdUsers.EditMode = adEditAdd Then
        DevEnvMDB.rscmdUsers.CancelBatch
        If (Not DevEnvMDB.rscmdUsers.EOF And Not DevEnvMDB.rscmdUsers.BOF) Then
            DevEnvMDB.rscmdUsers.MoveFirst
        End If
    ElseIf DevEnvMDB.rscmdUsers.EditMode = adEditInProgress Then
        DevEnvMDB.rscmdUsers.CancelUpdate
    End If
End Sub


Private Sub TabStrip1_Click()
    
    If pNoEvent Then Exit Sub
    
    If (Not DevEnvMDB.rscmdUsers.EOF And Not DevEnvMDB.rscmdUsers.BOF) Then
        If DevEnvMDB.rscmdUsers.EditMode <> adEditNone Then
            If MsgBox("Attenzione i dati sono stati cambiati, vuoi salvarli prima di procedere?", vbQuestion + vbYesNo) = vbYes Then
                cmd_Click 2
                
                If pErrore Then
                    pNoEvent = True
                    TabStrip1.Tabs(2).Selected = True
                    pNoEvent = False
                End If
            Else
                cancelCurrentOperation
            End If
        End If
    End If
    
    
    If TabStrip1.SelectedItem Is Nothing Then Exit Sub
    
    cmd(CMD_RICERCA).enabled = TabStrip1.SelectedItem.Index = 3
    cmd(CMD_RICERCA).default = TabStrip1.SelectedItem.Index = 3
    cmd(CMD_SALVA).enabled = TabStrip1.SelectedItem.Index = 2 And frmLogin.Livello <> LVL_USER
    cmd(CMD_SALVA).default = TabStrip1.SelectedItem.Index = 2 And frmLogin.Livello <> LVL_USER
    cmd(CMD_CHIUDI).default = TabStrip1.SelectedItem.Index = 1
    cmd(CMD_NUOVO).enabled = frmLogin.Livello <> LVL_USER
    cmd(CMD_CANCELLA).enabled = frmLogin.Livello <> LVL_USER
    
    If (TabStrip1.Tag <> "") Then
        Picture1(TabStrip1.Tag).Visible = False
    End If
        
    Picture1(TabStrip1.SelectedItem.Index - 1).Visible = True
    TabStrip1.Tag = TabStrip1.SelectedItem.Index - 1

    If (TabStrip1.SelectedItem.Index = 2) Then
        txtNome.SetFocus
        cmbLivello.ListIndex = txtLivello.Text
        txtPassword.Tag = txtPassword.Text
    End If

    If (TabStrip1.SelectedItem.Index = 3) Then
        txt(0).SetFocus
        
    End If


End Sub


Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    
        Case 1
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyCancel And KeyAscii <> vbKeyDelete Then
                KeyAscii = 0
            End If
    
    End Select
End Sub


