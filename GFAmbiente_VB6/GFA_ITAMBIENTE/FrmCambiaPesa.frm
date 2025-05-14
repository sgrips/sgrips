VERSION 5.00
Begin VB.Form frmCambiaPesa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambia Pesa"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdConferma 
      Caption         =   "Conferma"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   240
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Pesa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmCambiaPesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdConferma_Click()
    gPesaIsolaCod = Trim$(left$(Combo1.Text, 3))
    
    SaveSetting App.Title, "IMPOSTAZIONI", "CODPESAISOLA", gPesaIsolaCod
    
    InvioComandoAllaBilancia cmd_CambioCanale, True
    
    Unload Me
End Sub

Private Sub Form_Load()
    g_rsPese.MoveFirst
    g_rsPese.Sort = "ComID"
    
    Do While Not g_rsPese.EOF
        If nullToBlank(g_rsPese!PCNAME) = vbNullString And g_rsPese!active Then
            Combo1.AddItem g_rsPese!codPesa & " - " & g_rsPese!Descrizione
        End If
        
        If g_rsPese!codPesa = gPesaIsolaCod Then
             Combo1.ListIndex = Combo1.ListCount - 1
        End If
        
        g_rsPese.MoveNext
    Loop
    
    
End Sub
