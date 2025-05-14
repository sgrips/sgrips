VERSION 5.00
Begin VB.Form frmAttesa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operazioni di Chiusura"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Caption         =   "Attendere, chiusura in corso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6495
   End
End
Attribute VB_Name = "frmAttesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
