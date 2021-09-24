VERSION 5.00
Begin VB.Form autorizado 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1050
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1050
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2500
      Left            =   4770
      Top             =   225
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Acceso autorizado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   690
      Left            =   -135
      TabIndex        =   0
      Top             =   315
      Width           =   5730
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   0
      Picture         =   "autorizado.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5325
   End
End
Attribute VB_Name = "autorizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Timer1_Timer()
If Timer1.Interval = 2500 Then
    Unload autorizado
    Set autorizado = Nothing
End If

End Sub
