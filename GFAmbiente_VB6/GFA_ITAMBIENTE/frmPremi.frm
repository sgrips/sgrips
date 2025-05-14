VERSION 5.00
Object = "{5C8CED40-8909-11D0-9483-00A0C91110ED}#1.0#0"; "MSDATREP.OCX"
Begin VB.Form frmPremi 
   Caption         =   "Premi"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   15240
   Begin MSDataRepeaterLib.DataRepeater DTRepeter 
      Height          =   3270
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   5768
      _StreamID       =   -1412567295
      _Version        =   393216
      BorderStyle     =   0
      CaptionStyle    =   0
      Caption         =   ""
      BeginProperty RepeatedControlName {21FC0FC0-1E5C-11D1-A327-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         Name            =   "ITAmbienteUCPremi.UCPremi"
      EndProperty
      RepeaterBindings=   4
      BeginProperty RepeaterBinding(0) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "CodePremio"
         DataField       =   "idPremio"
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
      BeginProperty RepeaterBinding(1) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "Descrizione"
         DataField       =   "Descrizione"
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
      BeginProperty RepeaterBinding(2) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "Foto"
         DataField       =   "Immagine"
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
      BeginProperty RepeaterBinding(3) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "PremioDa"
         DataField       =   "DaPunti"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Rimuovi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1680
      TabIndex        =   8
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salva"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   12120
      TabIndex        =   5
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Chiudi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   13680
      TabIndex        =   4
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aggiungi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Descrizione"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   2535
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   0
      X2              =   15120
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      Caption         =   "Foto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   12600
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Con Punti"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   10680
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Premio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmPremi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            DevEnvMDB.rscmdPremi.CancelBatch adAffectAllChapters
            Unload Me
        Case 1
            If DevEnvMDB.rscmdPremi.EOF And Not DevEnvMDB.rscmdPremi.BOF Then
                If DTRepeter.ActiveRow < DTRepeter.RecordCount Then
                    DevEnvMDB.rscmdPremi.MoveNext
                    DevEnvMDB.rscmdPremi.MovePrevious
                ElseIf DTRepeter.ActiveRow = DTRepeter.RecordCount Then
                    DevEnvMDB.rscmdPremi.MovePrevious
                    DevEnvMDB.rscmdPremi.MoveNext
                End If
            End If
            DevEnvMDB.rscmdPremi.UpdateBatch adAffectAllChapters
        Case 3
            DevEnvMDB.rscmdPremi.AddNew
        Case 2
            If Not DevEnvMDB.rscmdPremi.EOF And Not DevEnvMDB.rscmdPremi.BOF Then
                DevEnvMDB.rscmdPremi.Delete adAffectCurrent
            End If
    End Select
        
End Sub



Private Sub Form_Load()
    If DevEnvMDB.rscmdPremi.state = adStateOpen Then
        DevEnvMDB.rscmdPremi.Close
    End If
    
    DevEnvMDB.cmdPremi
    Set DTRepeter.DataSource = DevEnvMDB.rscmdPremi
End Sub

Private Sub Form_Resize()
    DTRepeter.Height = Command1(0).top - DTRepeter.top
    DTRepeter.width = Me.ScaleWidth - DTRepeter.left - 150
    
    Command1(0).Move Me.ScaleWidth - Command1(0).width - 150, Me.ScaleHeight - Command1(0).Height - 150
    Command1(1).Move Command1(0).left - Command1(0).width - 150, Command1(0).top
    
    Command1(3).Move 150, Me.ScaleHeight - Command1(3).Height - 150
    Command1(2).Move Command1(3).left + Command1(3).width + 150, Command1(3).top
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set pfrmPremi = Nothing
End Sub


Public Sub showMe()
    Load Me
    
   
    If gStartMode = 1 Then
        Me.WindowState = vbMaximized
        Me.Show
    Else
        Me.Show
    End If
End Sub
