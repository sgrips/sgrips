VERSION 5.00
Begin VB.Form frmDettaglioCostiRifiuto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dettaglio"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Calcola"
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   14
      Top             =   3930
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Height          =   375
      Index           =   6
      Left            =   2955
      TabIndex        =   13
      Text            =   "Costo5"
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox Txt 
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
      Height          =   375
      Index           =   5
      Left            =   2970
      TabIndex        =   12
      Text            =   "Costo4"
      Top             =   2040
      Width           =   3495
   End
   Begin VB.TextBox Txt 
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
      Height          =   375
      Index           =   4
      Left            =   2970
      TabIndex        =   11
      Text            =   "Costo3"
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox Txt 
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
      Height          =   375
      Index           =   3
      Left            =   2970
      TabIndex        =   10
      Text            =   "Costo2"
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox Txt 
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
      Height          =   375
      Index           =   2
      Left            =   2970
      TabIndex        =   9
      Text            =   "Costo1"
      Top             =   600
      Width           =   3495
   End
   Begin VB.TextBox Txt 
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
      Height          =   375
      Index           =   0
      Left            =   2970
      TabIndex        =   8
      Text            =   "Costo Totale"
      Top             =   150
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Conferma"
      Height          =   375
      Index           =   0
      Left            =   5265
      TabIndex        =   7
      Top             =   3930
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Height          =   375
      Index           =   1
      Left            =   2955
      TabIndex        =   6
      Text            =   "Totale"
      Top             =   3330
      Width           =   3495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000016&
      X1              =   1695
      X2              =   6450
      Y1              =   3210
      Y2              =   3210
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      X1              =   1710
      X2              =   6480
      Y1              =   3195
      Y2              =   3195
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Costo 5:"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   5
      Top             =   2610
      Width           =   585
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Costo 4:"
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   4
      Top             =   2130
      Width           =   585
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Costo 3:"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   3
      Top             =   1650
      Width           =   585
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Costo 2:"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   2
      Top             =   1170
      Width           =   585
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Costo 1:"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   690
      Width           =   585
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Costo Totale:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   210
      Width           =   945
   End
End
Attribute VB_Name = "frmDettaglioCostiRifiuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Enum TipoDettaglio
    CostoRaccoltaeTraportoCRT
    CostoTrattamentoeSmaltimentoCTS
    CostoRaccoltaDifferenziataCRD
    CostoTrattamentoeRicicloCTR
    CostoSpazzamentoCSL
End Enum

Private mTipoDettaglio
Private mFrm As frmEditCostiDelComune
Private mTotaleRifiuto As Double
Private mCostoTotale As Double
Private OldTxtValue As String

Private Const CostoTotaleID = 0
Private Const TotaleID = 1

Private Const CostoPersonaleID = 2
Private Const CostoPercPersDirigenteID = 3
Private Const CostoAmmortamentiID = 4
Private Const CostoPercAmmStruttureID = 5
Private Const CostoAccantonamentiID = 6


Public Function CalcoloCostoRifiuto(frm As frmEditCostiDelComune, CostoTotale As Double, ByVal tpDettaglio As TipoDettaglio) As Double

    mTipoDettaglio = tpDettaglio
    Set mFrm = frm
    
    mCostoTotale = CostoTotale
    
    Me.Show vbModal
    
    CostoTotale = IIf(mCostoTotale < 0, 0, mCostoTotale)
    CalcoloCostoRifiuto = IIf(mTotaleRifiuto < 0, 0, mTotaleRifiuto)
End Function
    


Private Sub Command1_Click(Index As Integer)
    txt_LostFocus CostoTotaleID
    
    DoEvents
    
    If Index = 0 Then Set mFrm = Nothing: Unload Me
End Sub

Private Sub Form_Activate()
    Txt(CostoTotaleID).SetFocus
End Sub

Private Sub Form_Load()
    Dim s As String
    
    'Calcolo CRT - CRD - CTS - CSL a seconda di TipoDettaglio
    Select Case mTipoDettaglio
        Case CostoSpazzamentoCSL
            s = "CSL"
        Case CostoRaccoltaeTraportoCRT
            s = "CRT"
        Case CostoTrattamentoeRicicloCTR
            s = "CTR"
        Case CostoTrattamentoeSmaltimentoCTS
            s = "CTS"
        Case CostoRaccoltaDifferenziataCRD
            s = "CRD"
    End Select
    
    Txt(1) = 0: Txt(2) = 0
    Txt(3) = 0: Txt(4) = 0
    Txt(5) = 0: Txt(6) = 0
    
    Txt(CostoTotaleID) = mCostoTotale
    
    Me.Caption = "Dettaglio - costo " & s
    
    
    lbl(CostoTotaleID).Caption = "Costo Totale " & s
    lbl(CostoPersonaleID).Caption = "Costo Personale " & s
    lbl(CostoPercPersDirigenteID).Caption = "% Costo Personale Dirigente"
    lbl(CostoAmmortamentiID).Caption = "Ammortamenti " & s
    lbl(CostoPercAmmStruttureID).Caption = "% Ammortamenti Strutture Comuni"
    lbl(CostoAccantonamentiID).Caption = "Accantonamenti "
    
    
    Txt(TotaleID).BackColor = BACKCOLOR_CAMPI_CALC
    Txt(TotaleID).Locked = True
    Txt(CostoPersonaleID).BackColor = BACKCOLOR_CAMPI_CALC
    Txt(CostoPersonaleID).Locked = True
    Txt(CostoPercPersDirigenteID).BackColor = BACKCOLOR_CAMPI_CALC
    Txt(CostoPercPersDirigenteID).Locked = True
    Txt(CostoAmmortamentiID).BackColor = BACKCOLOR_CAMPI_CALC
    Txt(CostoAmmortamentiID).Locked = True
    Txt(CostoPercAmmStruttureID).BackColor = BACKCOLOR_CAMPI_CALC
    Txt(CostoPercAmmStruttureID).Locked = True
    Txt(CostoAccantonamentiID).BackColor = BACKCOLOR_CAMPI_CALC
    Txt(CostoAccantonamentiID).Locked = True
    
    txt_LostFocus (CostoTotaleID)

End Sub


Private Sub txt_GotFocus(Index As Integer)
    If Txt(Index).BackColor = BACKCOLOR_CAMPI_CALC Then Exit Sub

    Txt(Index).MaxLength = 12
    Txt(Index).Text = CDbl(Txt(Index).Text)
    Txt(Index).SelStart = 0
    Txt(Index).SelLength = Len(Txt(Index))
    OldTxtValue = Txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Dim CostoPersonale As Double
    Dim CostoPercPersDirigente  As Double
    Dim CostoAmmortamenti As Double
    Dim CostoPercAmmStrutture As Double
    Dim CostoAccantonamenti As Double
    Dim i As Integer
    
    Txt(Index).MaxLength = 15
    With Txt
        If Index = CostoTotaleID Then
            Select Case mTipoDettaglio
                Case CostoSpazzamentoCSL
                    .Item(TotaleID).Text = mFrm.CostiSpazzamento(CDbl(.Item(CostoTotaleID).Text), CostoPersonale, CostoPercPersDirigente, CostoAmmortamenti, CostoPercAmmStrutture)
                Case CostoRaccoltaeTraportoCRT
                    .Item(TotaleID).Text = mFrm.CostiRaccoltaETrasporto(CDbl(.Item(CostoTotaleID).Text), CostoPersonale, CostoPercPersDirigente, CostoAmmortamenti, CostoPercAmmStrutture)
                Case CostoTrattamentoeSmaltimentoCTS
                    .Item(TotaleID) = mFrm.CostiSmaltimentoETrattamento(CDbl(.Item(CostoTotaleID)), CostoPersonale, CostoPercPersDirigente, CostoAmmortamenti, CostoPercAmmStrutture, CostoAccantonamenti)
                Case CostoRaccoltaDifferenziataCRD
                    .Item(TotaleID) = mFrm.CostiRaccoltaDifferenziataMateriale(CDbl(.Item(CostoTotaleID)), CostoPersonale, CostoPercPersDirigente, CostoAmmortamenti, CostoPercAmmStrutture)
                Case CostoTrattamentoeRicicloCTR
                    .Item(TotaleID) = mFrm.CostiTrattamentoRiciclo(CDbl(.Item(CostoTotaleID)), CostoPersonale, CostoPercPersDirigente, CostoAmmortamenti, CostoPercAmmStrutture, CostoAccantonamenti)
            End Select
        

    
            mTotaleRifiuto = CDbl(.Item(TotaleID))
            .Item(CostoPersonaleID) = CostoPersonale
            .Item(CostoPercPersDirigenteID) = CostoPercPersDirigente
            .Item(CostoAmmortamentiID) = CostoAmmortamenti
            .Item(CostoPercAmmStruttureID) = CostoPercAmmStrutture
            .Item(CostoAccantonamentiID) = CostoAccantonamenti
            
            mCostoTotale = CDbl(Txt(CostoTotaleID))
            
            For i = 0 To .Count - 1
                .Item(i) = Format$(CDbl(.Item(i)), FORMATVALUTA)
            Next
        End If
    End With

End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
        'Il valore deve essere numerico
        If Not IsNumeric(Txt(Index)) Then MsgBox "Attenzione il valore deve essere numerico", vbOKOnly + vbExclamation, App.Title: Cancel = True: Txt(Index) = OldTxtValue: Exit Sub

End Sub
