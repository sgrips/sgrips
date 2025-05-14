VERSION 5.00
Begin {78E93846-85FD-11D0-8487-00A0C90DC8A9} prvoa 
   Bindings        =   "Registro1.dsx":0000
   Caption         =   "Registro"
   ClientHeight    =   13395
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   12030
   StartUpPosition =   1  'CenterOwner
   _ExtentX        =   21220
   _ExtentY        =   23627
   _Version        =   393216
   _DesignerVersion=   100688210
   ReportWidth     =   11684
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   GridX           =   5
   GridY           =   5
   LeftMargin      =   100
   TopMargin       =   500
   _Settings       =   31
   DataMember      =   "cmdRegistro"
   NumSections     =   1
   SectionCode0    =   4
   BeginProperty Section0 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "cmdRegistro_Dettaglio"
      Object.Height          =   4309
      NumControls     =   2
      ItemType0       =   4
      BeginProperty Item0 {1C13A8E2-A0B6-11D0-848E-00A0C90DC8A9} 
         _Version        =   393216
         Name            =   "TextBox1"
         Object.Left            =   794
         Object.Top             =   1020
         Object.Width           =   3741
         Object.Height          =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataField       =   "riferimenti"
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
      ItemType1       =   4
      BeginProperty Item1 {1C13A8E2-A0B6-11D0-848E-00A0C90DC8A9} 
         _Version        =   393216
         Name            =   "TextBox2"
         Object.Left            =   794
         Object.Top             =   794
         Object.Width           =   1440
         Object.Height          =   240
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataField       =   "Scarico"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
         DataMember      =   "cmdRegistro"
      EndProperty
   EndProperty
End
Attribute VB_Name = "prvoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Stampa(where As String, preview As Boolean)
    Dim defaultPrinter As String

    On Error GoTo StampaERR
    defaultPrinter = Printer.DeviceName
    
    SetWinDefPrinter gStampaRegistroStampante
    
    Me.orientation = rptOrientPortrait
    If DevEnvMDB.rscmdStampaPesata.State = adStateOpen Then
        DevEnvMDB.rscmdStampaPesata.Close
    End If
    
    
    If DevEnvMDB.rscmdRegistro.State = adStateOpen Then
        DevEnvMDB.rscmdRegistro.Close
    End If
    
    If (allPrint) Then where = ""
    
    DevEnvMDB.rscmdRegistro.Source = Replace(DevEnvMDB.rscmdRegistro.Source, "1=1", where)
    
    DevEnvMDB.cmdRegistro
    
    
    
    If (DevEnvMDB.rscmdRegistro.EOF = False) Then
    
        If preview Then
            Me.Show
        Else
            Me.PrintReport
        End If
    End If
    
    SetWinDefPrinter defaultPrinter
    
    Exit Sub
StampaERR:
    
    SetWinDefPrinter defaultPrinter
End Sub

