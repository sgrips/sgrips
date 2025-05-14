VERSION 5.00
Begin VB.Form frmSimuCalcolo 
   Caption         =   "frmSimCalcolo"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   12630
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1815
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   2880
      Width           =   11895
      Begin VB.CommandButton Command1 
         Caption         =   "Salva Tariffa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   9960
         TabIndex        =   23
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   5640
         TabIndex        =   21
         Text            =   "0"
         Top             =   480
         Width           =   1125
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   120
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "0"
         Top             =   1320
         Width           =   1200
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "0"
         Top             =   480
         Width           =   1200
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   3360
         TabIndex        =   10
         Text            =   "0"
         Top             =   1320
         Width           =   1050
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   3360
         TabIndex        =   9
         Text            =   "0"
         Top             =   480
         Width           =   1050
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ricalcola Tariffa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   8160
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "% Ripart. Canone UD:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   37
         Left            =   5640
         TabIndex        =   22
         Top             =   120
         Width           =   2355
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   19
         Left            =   1440
         TabIndex        =   20
         Top             =   480
         Width           =   210
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   20
         Left            =   1440
         TabIndex        =   19
         Top             =   1320
         Width           =   210
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Recupero Produttività:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   26
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   2355
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Inflazione Programmata:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   27
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   2595
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "CU Attività:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   45
         Left            =   3360
         TabIndex        =   16
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "CU Domestico:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   46
         Left            =   3360
         TabIndex        =   15
         Top             =   120
         Width           =   1605
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   47
         Left            =   4560
         TabIndex        =   14
         Top             =   1320
         Width           =   210
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   48
         Left            =   4560
         TabIndex        =   13
         Top             =   480
         Width           =   210
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   600
      TabIndex        =   5
      Top             =   4800
      Width           =   7455
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   6975
      Begin VB.CommandButton Command1 
         Caption         =   "Salva Tariffa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   5160
         TabIndex        =   24
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ricalcola Tariffa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   3360
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   4245
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "0"
         Top             =   360
         Visible         =   0   'False
         Width           =   2280
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Detrazione spazzamento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   6600
         TabIndex        =   3
         Top             =   405
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "()"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   1
         Left            =   2910
         TabIndex        =   2
         Top             =   375
         Visible         =   0   'False
         Width           =   150
      End
   End
End
Attribute VB_Name = "frmSimuCalcolo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pLastParent As Long
Private noEvent  As Boolean
Private mCodComune As String
Private mAnnoRiferimento As String
Private mTipoCalcolo As tpoTariffa


Private Const MINMAXPERCDETRAZIONE_ID = 1
Private Const PAG1_PERCDETRASPAZZAMENTO_ID = 0
Private Const PAG1_N_IP_ID = 1
Private Const PAG1_N_RP_ID = 2
Private Const PAG1_N_PERCINIUD_ID = 3
Private Const PAG1_N_CUD_ID = 4
Private Const PAG1_N_CUA_ID = 5
Private frm As FrmCalcoloTariffa


Private OldTxtValue As String

Private mMAXPERCDETRAZIONE As Double, mMINPERCDETRAZIONE As Double

Private Sub Command1_Click(index As Integer)
    Select Case index
        Case 0
            frm.ImpostaCampi507 Txt(PAG1_PERCDETRASPAZZAMENTO_ID).Text
            frm.VaiAllaFine
        Case 1
            frm.ImpostaCampiNormalizzato Txt(PAG1_N_IP_ID).Text, Txt(PAG1_N_RP_ID).Text, Txt(PAG1_N_PERCINIUD_ID).Text, Txt(PAG1_N_CUA_ID).Text, Txt(PAG1_N_CUD_ID).Text
            frm.VaiAllaFine
        Case 2, 3
            frm.SalvaTariffa
    End Select
End Sub

Private Sub Form_Load()
    noEvent = True
End Sub

Private Sub Form_Resize()
    Dim i  As Integer
    
    If Me.WindowState = vbMinimized Or noEvent Then Exit Sub
    Frame1(0).Move 35, 35, Me.ScaleWidth - 35
    Frame1(1).Move 35, 35, Me.ScaleWidth - 35
    Frame2.Move 0, Frame1(1).Height + 50, Me.Width - 100, Me.ScaleHeight - Frame1(1).Height
    
    
    With frm
        .Move 0, 0, Frame2.Width, Frame2.Height
        
    
      
    
    
        For i = .Picture1.LBound To .Picture1.UBound
            .Picture1(i).Height = frm.ScaleHeight
            .Picture1(i).Width = frm.ScaleWidth
        Next
    
        For i = .picture2.LBound To .picture2.UBound
            .picture2(i).Height = frm.ScaleHeight
            .picture2(i).Width = frm.ScaleWidth
        Next
    
        For i = .Picture3.LBound To .Picture3.UBound
            .Picture3(i).Height = frm.ScaleHeight
            .Picture3(i).Width = frm.ScaleWidth
        Next
        
        For i = .Picture4.LBound To .Picture4.UBound
            .Picture4(i).Height = frm.ScaleHeight
            .Picture4(i).Width = frm.ScaleWidth
        Next
        
        For i = frm.Frame1.LBound To frm.Frame1.UBound
            .Frame1(i).Height = frm.ScaleHeight - 200
            .Frame1(i).Width = frm.ScaleWidth - 200
        Next
        
        .lvw9.Height = .Frame1(11).Height - .lvw9.Top - 200
        .lvw9.Width = .Frame1(11).Width - .lvw9.Left - 200
        .lvw10.Height = .Frame1(11).Height - .lvw10.Top - 200
        .lvw10.Width = .Frame1(11).Width - .lvw10.Left - 200

    End With
    
    Command1(0).Move Frame1(0).Width - Command1(0).Width - 50, Frame1(0).Height - Command1(0).Height - 50
    Command1(3).Move Command1(0).Left - Command1(1).Width - 50, Frame1(0).Height - Command1(0).Height - 50
    
    Command1(1).Move Frame1(1).Width - Command1(1).Width - 50, Frame1(1).Height - Command1(1).Height - 50
    Command1(2).Move Command1(1).Left - Command1(1).Width - 50, Frame1(1).Height - Command1(1).Height - 50
End Sub

Public Sub Carica(CodComune As String, AnnoRiferimento As String, tipoCalcolo As tpoTariffa)
    
    Load Me
        
    Set frm = New FrmCalcoloTariffa
        
    Frame1(0).Visible = tipoCalcolo = Tariffa507
    Frame1(1).Visible = tipoCalcolo = TariffaNormalizzata Or tipoCalcolo = TariffaCostiIndivisibili
    mCodComune = CodComune
    mAnnoRiferimento = AnnoRiferimento
    mTipoCalcolo = tipoCalcolo
    
        
    TitleBar frm, False
    
    pLastParent = GetParent(frm.hwnd)
    
    frm.Hide
    
    DoEvents
    
    SetParent frm.hwnd, Frame2.hwnd
   
  
    frm.Avvio mCodComune, mAnnoRiferimento, tipoCalcolo
     
    If tipoCalcolo = Tariffa507 Then
        Txt(PAG1_PERCDETRASPAZZAMENTO_ID) = GetParamValue(mAnnoRiferimento, mCodComune, "PERCDETRAZIONE", mMAXPERCDETRAZIONE, mMINPERCDETRAZIONE)
        lbl(MINMAXPERCDETRAZIONE_ID) = "(" & Format$(mMINPERCDETRAZIONE, "##0.0") & " -  " & Format$(mMAXPERCDETRAZIONE, "##0.0") & " )"
    Else
        Txt(PAG1_N_IP_ID) = GetParamValue(mAnnoRiferimento, mCodComune, "INFLAZIO")
        Txt(PAG1_N_IP_ID).Tag = Txt(PAG1_N_IP_ID)
        Txt(PAG1_N_RP_ID) = GetParamValue(mAnnoRiferimento, mCodComune, "COEFFRIP")
        Txt(PAG1_N_RP_ID).Tag = Txt(PAG1_N_RP_ID)
            

        Txt(PAG1_N_CUA_ID) = GetParamValue(mAnnoRiferimento, mCodComune, "PERCCUA")
        Txt(PAG1_N_CUD_ID) = GetParamValue(mAnnoRiferimento, mCodComune, "PERCCUD")

        Txt(PAG1_N_PERCINIUD_ID) = GetParamValue(mAnnoRiferimento, mCodComune, "PERCINIUD")

        
    End If
    
    Command1_Click (1)
    noEvent = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frm.Visible = True
    
    SetParent frm.hwnd, pLastParent
    
    Unload frm
End Sub


Private Sub txt_GotFocus(index As Integer)
    If Txt(index).BackColor = BACKCOLOR_CAMPI_CALC Then Exit Sub
    
    'Imposto il numero MIN di caratteri in EDIT MODE
    If Txt(index).Text <> vbNullString Then Txt(index) = CDbl(Txt(index))
    'se Percentuale
    Select Case index
        Case PAG1_N_IP_ID, PAG1_N_RP_ID
                
                Txt(index).MaxLength = 4
        Case Else
                Txt(index).MaxLength = 15
    End Select
    
    DoEvents
    Txt(index).SelStart = 0
    Txt(index).SelLength = Len(Txt(index))
    OldTxtValue = Txt(index)
End Sub

Private Sub txt_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
End Sub

Private Sub txt_LostFocus(index As Integer)
    If Txt(index).Locked Then Exit Sub
    
    'Imposto il numero MIN di caratteri in EDIT MODE
    Txt(index) = CDbl(Txt(index))
    Txt(index).MaxLength = 15
    

    'se Percentuale
    Select Case index
        Case PAG1_N_IP_ID, PAG1_N_RP_ID, _
              PAG1_N_PERCINIUD_ID, _
              PAG1_N_CUD_ID, _
              PAG1_N_CUA_ID

              
            
            If index = PAG1_N_PERCINIUD_ID Then
                LetParamValue mAnnoRiferimento, mCodComune, "PERCINIUD", True, Txt(index).Text
                Txt(index) = Format$(CDbl(Txt(index)), "##0.0000")
            End If
              
            
            If index = PAG1_N_IP_ID Then
                LetParamValue mAnnoRiferimento, mCodComune, "INFLAZIO", False, Txt(index)
            End If
        
            If index = PAG1_N_RP_ID Then
                LetParamValue mAnnoRiferimento, mCodComune, "COEFFRIP", True, Txt(index)
            End If

        Case Else
        
            Txt(index) = Format$(CDbl(Txt(index)), FORMATVALUTA)
    End Select
    
    
    'Calcolo
    If Txt(index).BackColor <> BACKCOLOR_CAMPI_CALC Then
        Txt(PAG1_N_PERCINIUD_ID).Enabled = False


        Select Case index
            Case PAG1_PERCDETRASPAZZAMENTO_ID
                frm.allineaCampoCollegato 49, Txt(index)
            Case PAG1_N_IP_ID
                frm.allineaCampoCollegato 17, Txt(index)
            Case PAG1_N_RP_ID
                frm.allineaCampoCollegato 16, Txt(index)
            Case PAG1_N_PERCINIUD_ID
                frm.allineaCampoCollegato 28, Txt(index)
            Case PAG1_N_CUD_ID
                frm.allineaCampoCollegato 37, Txt(index)
                Txt(PAG1_N_CUA_ID) = frm.getValueTxt(36)
            Case PAG1_N_CUA_ID
                frm.allineaCampoCollegato 36, Txt(index)
                Txt(PAG1_N_CUD_ID) = frm.getValueTxt(37)
        End Select
        
        
            
        
        Txt(PAG1_N_PERCINIUD_ID).Enabled = True
    
    End If
    
    If index = PAG1_N_CUD_ID Or index = PAG1_N_CUA_ID Then
        LetParamValue mAnnoRiferimento, mCodComune, "PERCCUD", True, Txt(PAG1_N_CUD_ID).Text
        Txt(PAG1_N_CUD_ID) = Format$(CDbl(Txt(PAG1_N_CUD_ID)), FORMAT4DEC)
    
        LetParamValue mAnnoRiferimento, mCodComune, "PERCCUA", True, Txt(PAG1_N_CUA_ID).Text
        Txt(PAG1_N_CUA_ID) = Format$(CDbl(Txt(PAG1_N_CUA_ID)), FORMAT4DEC)
    End If
    

End Sub

Private Sub txt_Validate(index As Integer, Cancel As Boolean)
    If Trim$(Txt(index)) = vbNullString Then Exit Sub
    
   'Il valore deve essere numerico
    If Not IsNumeric(Txt(index)) Then MsgBox "Attenzione il valore deve essere numerico", vbOKOnly + vbExclamation, App.Title: Cancel = True: Txt(index) = OldTxtValue: Exit Sub
    
    Select Case index
         
         'Il valore deve essere una Percentuale tra 1 - 99
         Case PAG1_N_PERCINIUD_ID
                
            If Txt(index) > 99 Or Txt(index) <= 0 Then
                MsgBox "Attenzione il valore deve essere una percentuale", vbInformation + vbOKOnly, App.Title
                Cancel = True
            End If
                
                 
         Case PAG1_N_IP_ID, PAG1_N_RP_ID, _
               PAG1_PERCDETRASPAZZAMENTO_ID
            If Txt(index) > 100 Or Txt(index) < 0 Then
                MsgBox "Attenzione il valore deve essere una percentuale", vbInformation + vbOKOnly, App.Title
                Cancel = True
            End If
            
            If index = PAG1_PERCDETRASPAZZAMENTO_ID And (CDbl(Txt(PAG1_PERCDETRASPAZZAMENTO_ID)) > mMAXPERCDETRAZIONE Or CDbl(Txt(PAG1_PERCDETRASPAZZAMENTO_ID)) < mMINPERCDETRAZIONE) Then
                MsgBox "Attenzione il valore deve essere compreso fra " & mMINPERCDETRAZIONE & " e " & mMAXPERCDETRAZIONE, vbInformation + vbOKOnly, App.Title
                Cancel = True
            End If
            
            
    End Select
End Sub
