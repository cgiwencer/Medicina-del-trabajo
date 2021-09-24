VERSION 5.00
Begin VB.Form Cerrado 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2500
      Left            =   0
      Top             =   45
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sesión cerrada"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   690
      Left            =   90
      TabIndex        =   1
      Top             =   1350
      Width           =   4740
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Acceso autorizado"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   690
      Left            =   -405
      TabIndex        =   0
      Top             =   630
      Width           =   5730
   End
   Begin VB.Image Image1 
      Height          =   3570
      Left            =   0
      Picture         =   "Cerrado.frx":0000
      Stretch         =   -1  'True
      Top             =   -45
      Width           =   5400
   End
End
Attribute VB_Name = "Cerrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
If Timer1.Interval = 2500 Then
    Unload Cerrado
    Set Cerrado = Nothing
End If
End Sub
