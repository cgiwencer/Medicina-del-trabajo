VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Direcciones 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10650
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10650
   ScaleWidth      =   11535
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame SSFrame1 
      Height          =   3930
      Left            =   225
      TabIndex        =   3
      Top             =   855
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   6932
      _Version        =   196608
      BackColor       =   16777215
      BackStyle       =   1
      Caption         =   "DIRECCION DE EXÁMEN NORMAL"
      ClipControls    =   0   'False
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFC0&
         Height          =   1095
         Left            =   405
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2475
         Width           =   9960
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFC0&
         Height          =   1095
         Left            =   360
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   9960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Laboratorio Clínico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   450
         TabIndex        =   7
         Top             =   2160
         Width           =   1950
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Rayos X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   405
         TabIndex        =   6
         Top             =   405
         Width           =   1635
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   1095
      Left            =   3825
      TabIndex        =   0
      Top             =   9225
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   1931
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin Threed.SSCommand SSCommand1 
         Height          =   465
         Left            =   2940
         TabIndex        =   13
         Top             =   270
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   820
         _Version        =   196608
         ForeColor       =   16777215
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Guardar"
         ButtonStyle     =   4
         BevelWidth      =   0
      End
      Begin VB.Image Image11 
         Height          =   330
         Left            =   2475
         Picture         =   "Direccones.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5445
         TabIndex        =   1
         Top             =   405
         Width           =   1710
      End
      Begin VB.Image Image5 
         Height          =   285
         Left            =   5355
         Picture         =   "Direccones.frx":1D7B
         Stretch         =   -1  'True
         Top             =   360
         Width           =   285
      End
      Begin VB.Image Image8 
         Height          =   465
         Left            =   5220
         Picture         =   "Direccones.frx":44F2
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   270
         Width           =   1905
      End
      Begin VB.Image Image12 
         Height          =   555
         Left            =   2295
         Picture         =   "Direccones.frx":6C33
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   225
         Width           =   2175
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   3930
      Left            =   225
      TabIndex        =   8
      Top             =   4995
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   6932
      _Version        =   196608
      BackStyle       =   1
      Caption         =   "DIRECCION DE EXÁMEN DE URGENCIA"
      ClipControls    =   0   'False
      Begin VB.TextBox Text4 
         BackColor       =   &H0080FFFF&
         Height          =   1095
         Left            =   450
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   2475
         Width           =   9960
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H0080FFFF&
         Height          =   1095
         Left            =   450
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   855
         Width           =   9960
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Rayos X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   495
         TabIndex        =   12
         Top             =   540
         Width           =   1950
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Laboratorio Clínico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   495
         TabIndex        =   11
         Top             =   2160
         Width           =   1950
      End
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "EXAMENES"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9840
      TabIndex        =   2
      Top             =   210
      Width           =   1365
   End
   Begin VB.Image Image22 
      Height          =   600
      Left            =   45
      Picture         =   "Direccones.frx":962D
      Stretch         =   -1  'True
      Top             =   45
      Width           =   600
   End
   Begin VB.Image Image14 
      Height          =   690
      Left            =   0
      Picture         =   "Direccones.frx":1408A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11625
   End
End
Attribute VB_Name = "Direcciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim Cn As New ADODB.Connection
Dim rsdi As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open

rsdi.CursorType = adOpenKeyset
rsdi.LockType = adLockOptimistic
rsdi.ActiveConnection = Cn
rsdi.Source = "Select * from direcciones"
rsdi.Open

If Not rsdi.EOF Then
    Text1.Text = rsdi!DirRxN
    Text2.Text = rsdi!DirLabN
    Text3.Text = rsdi!DirRxU
    Text4.Text = rsdi!DirLabU
End If
Cn.Close
End Sub

Private Sub Label10_Click()
Unload Direcciones
Set Direcciones = Nothing
End Sub

Private Sub SSCommand1_Click()
Dim Cn As New ADODB.Connection
Dim rsdi As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open

rsdi.CursorType = adOpenKeyset
rsdi.LockType = adLockOptimistic
rsdi.ActiveConnection = Cn
rsdi.Source = "Select * from direcciones"
rsdi.Open

vDirRxN = Text1.Text
vDirLabN = Text2.Text
vDirRxU = Text3.Text
vDirLabU = Text4.Text

If Not rsdi.EOF Then
    corrige = "UPDATE direcciones set DirRxN = " & "'" & vDirRxN & "',DirLabN = " & "'" & vDirLabN & "', DirRxU = " & "'" & vDirRxU & "',DirLabU = " & "'" & vDirLabU & "'"
    Cn.Execute corrige
Else
    nuevo = "INSERT INTO direcciones set DirRxN = " & "'" & vDirRxN & "',DirLabN = " & "'" & vDirLabN & "', DirRxU = " & "'" & vDirRxU & "',DirLabU = " & "'" & vDirLabU & "'"
    Cn.Execute nuevo
End If
Cn.Close
MsgBox "Direcciones registradas"
Label10_Click
End Sub

Private Sub Text1_LostFocus()
Text1.Text = UCase(Text1.Text)
End Sub
Private Sub Text2_LostFocus()
Text2.Text = UCase(Text2.Text)
End Sub

Private Sub Text3_LostFocus()
Text3.Text = UCase(Text3.Text)
End Sub

Private Sub Text4_LostFocus()
Text4.Text = UCase(Text4.Text)
End Sub
