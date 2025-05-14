VERSION 5.00
Begin VB.Form frmBackupDB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup \ Restore DataBase"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Restore"
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
      Left            =   2160
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Backup"
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
      Left            =   3360
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Chiudi"
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
      Left            =   4560
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmBackupDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            BackupRestore True, True
        Case 2
            BackupRestore False, True
    End Select
    
End Sub

Private Sub Form_Load()
    Label1.Caption = "La funzione di backup permette di salvare lo stato del database in un " & _
            "certo istante, questo permette in succesivo momento di recuperare (attraverso la funzione di Restore) dopo un crash DB lo stato del DataBase " & _
            "almeno fino all'ultimo backup." & vbCrLf
    Label1.Caption = Label1.Caption & "La funzione di Restore permette di recuperare lo stato del" & _
            "database fino all'ultimo backup fatto, una volta fatto un Restore non è più possibile tornare alla versione precedente di Batabase" & vbCrLf

End Sub


Public Sub BackupRestore(flgBackup As Boolean, msgConfirm As Boolean)
    On Error GoTo BackupERR
    
    
    
    If Forms.Count > 4 Then
        MsgBox "Prima di utilizzare questa funzione occorre uscire da tutte le videate aperte", vbExclamation + vbOKOnly
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    For i = 1 To DevEnvMDB.Connections.Count
        DevEnvMDB.Connections(i).Close
    Next
    
    DoEvents
    
    ChiudiConnessione
    
    
    Dim f As FileSystemObject
    Set f = New FileSystemObject
    
    
    If f.DriveExists(gPercorsoBackup) = True Or (f.DriveExists(gPercorsoBackup) = False And f.FolderExists(gPercorsoBackup)) Then
        Dim fOld As String
        Dim fbak As String
        
        fbak = f.BuildPath(gPercorsoBackup, "IsolaEcologica.bak")
        If flgBackup Then
            fOld = f.BuildPath(gPercorsoBackup, "IsolaEcologica.old")
            
            'cancello ultimo backup
            If f.FileExists(fOld) Then f.DeleteFile fOld, True
            'salvo ultimo backup
            If f.FileExists(fbak) Then f.MoveFile fbak, fOld
            'faccio backup
            f.CopyFile gPathDB, fbak, True
        Else
            fOld = f.BuildPath(Replace(gPathDB, f.GetFileName(gPathDB), vbNullString), "IsolaEcologica.old")
            
            'cancello ultimo save
            If f.FileExists(fOld) Then f.DeleteFile fOld, True
            'salvo ultimo MDB
            If f.FileExists(gPathDB) Then f.MoveFile gPathDB, fOld
            'faccio restore
            f.CopyFile fbak, gPathDB, True
        End If
    Else
        MsgBox "Controllare percorso backup " & gPercorsoBackup & " che risulta non valido"
    End If
    
    
    ApriConnessione
    
    For i = 1 To DevEnvMDB.Connections.Count
        DevEnvMDB.Connections(i).Open
    Next
    
    
    Screen.MousePointer = vbDefault
    DoEvents
    
    If msgConfirm Then
    
        MsgBox IIf(flgBackup, "Backup", "Restore") & " DataBase effettuato con successo!!", vbExclamation + vbOKOnly, App.Title
    
    End If
  
    
    Exit Sub
BackupERR:
    MsgBox Error$, vbCritical, App.Title
    On Error GoTo 0
    On Error Resume Next
    
    For i = 1 To DevEnvMDB.Connections.Count
        DevEnvMDB.Connections(i).Open
    Next
    
    ApriConnessione
    Screen.MousePointer = vbDefault
End Sub



