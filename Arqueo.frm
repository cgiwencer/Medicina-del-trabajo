VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Arqueo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3375
   ClientLeft      =   4980
   ClientTop       =   3255
   ClientWidth     =   4785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1800
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   2265
      Left            =   180
      TabIndex        =   2
      Top             =   930
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   3995
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo2 
         Height          =   375
         Left            =   2145
         TabIndex        =   1
         Top             =   615
         Width           =   1545
         _Version        =   65537
         _ExtentX        =   2725
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483633
      End
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo1 
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   615
         Width           =   1635
         _Version        =   65537
         _ExtentX        =   2884
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483633
      End
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo3 
         Height          =   375
         Left            =   270
         TabIndex        =   8
         Top             =   1485
         Width           =   1545
         _Version        =   65537
         _ExtentX        =   2725
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483633
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Envío"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   330
         TabIndex        =   9
         Top             =   1260
         Width           =   1380
      End
      Begin VB.Image Image5 
         Height          =   285
         Left            =   2250
         Picture         =   "Arqueo.frx":0000
         Stretch         =   -1  'True
         Top             =   1725
         Width           =   345
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
         Height          =   240
         Left            =   2295
         TabIndex        =   7
         Top             =   1725
         Width           =   1920
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   2310
         TabIndex        =   6
         Top             =   1245
         Width           =   1815
      End
      Begin VB.Image Image6 
         Height          =   285
         Left            =   2250
         Picture         =   "Arqueo.frx":2777
         Stretch         =   -1  'True
         Top             =   1245
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2340
         TabIndex        =   4
         Top             =   390
         Width           =   1005
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Inicio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   390
         Width           =   1335
      End
      Begin VB.Image Image7 
         Height          =   465
         Left            =   2115
         Picture         =   "Arqueo.frx":57D9
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   1155
         Width           =   1965
      End
      Begin VB.Image Image8 
         Height          =   465
         Left            =   2115
         Picture         =   "Arqueo.frx":7D8E
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   1635
         Width           =   1965
      End
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "LISTADO DIARIO"
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
      Left            =   2880
      TabIndex        =   5
      Top             =   225
      Width           =   2055
   End
   Begin VB.Image Image22 
      Height          =   600
      Left            =   45
      Picture         =   "Arqueo.frx":A4CF
      Stretch         =   -1  'True
      Top             =   45
      Width           =   600
   End
   Begin VB.Image Image14 
      Height          =   690
      Left            =   0
      Picture         =   "Arqueo.frx":14F2C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4920
   End
End
Attribute VB_Name = "Arqueo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
KeyPreview = True

End Sub
Sub ChgEnterToTab(KeyCode As Integer)
If KeyCode = 13 Then
   KeyCode = 0
   SendKeys "{TAB}"
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   SendKeys "{TAB}"
   KeyAscii = 0
End If
End Sub

Private Sub Image1_Click()

End Sub

Private Sub SSCommand1_Click()
vfechaini = SSDateCombo1.Text
vfechafin = SSDateCombo2.Text

CrystalReport1.ReportFileName = App.Path & "\arqueo.rpt"
CrystalReport1.Formulas(0) = "a = " & "'" & vfechaini & "'"
CrystalReport1.Formulas(1) = "de = " & "'" & vfechafin & "'"
CrystalReport1.Action = 1

End Sub

Private Sub Label1_Click()
Dim Cn As New ADODB.Connection
Dim rsar As New ADODB.Recordset ' Recordset de arqueo
Cn.ConnectionString = Cadena
Cn.Open
vfechai = SSDateCombo1.Date
vfechaf = SSDateCombo2.Date
vfechai = Format(vfechai, "YYYY-MM-dd")
vfechaf = Format(vfechaf, "yyyy-mm-dd")

borrat = "DELETE FROM implisd"
Cn.Execute borrat

grabalista = "INSERT INTO implisd select * from nuevos where FecRev BETWEEN " & "'" & vfechai & "'" & " AND " & "'" & vfechaf & "' AND NueEst = " & 6
Cn.Execute grabalista

vfechaini = SSDateCombo1.Text
vfechafin = SSDateCombo2.Text
vfechaenv = SSDateCombo3.Text


CrystalReport1.ReportFileName = App.Path & "\listadod.rpt"
CrystalReport1.Formulas(0) = "del = " & "'" & vfechaini & "'"
CrystalReport1.Formulas(1) = "al = " & "'" & vfechafin & "'"
CrystalReport1.Formulas(2) = "envio = " & "'" & vfechaenv & "'"
CrystalReport1.Action = 1

End Sub


Private Sub SSOleDBGrid2_DblClick()
vegr_id = Adodc1.Recordset.Fields("cegr_id")
Dim Cn As New ADODB.Connection
Dim rsve As New ADODB.Recordset   ' Recordset de det de venta
Cn.ConnectionString = Cadena
Cn.Open
    
rsve.CursorType = adOpenKeyset
rsve.LockType = adLockOptimistic
rsve.ActiveConnection = Cn
rsve.Source = "Select * from vventa1 Where cegr_id = " & vegr_id
rsve.Open

If Not rsve.EOF Then
    Load Detventa
    Detventa.SSDateCombo1.Text = rsve!pag_fec
    Detventa.Label8.Caption = rsve!pag_nit
    Detventa.Label1.Caption = rsve!pag_ras
    Detventa.Label2.Caption = rsve!pag_NFa
    Detventa.Adodc1.RecordSource = "Select * from vventa1 where cegr_id = " & vegr_id
    Detventa.Adodc1.Refresh
    Detventa.Show
End If
End Sub

Private Sub Label10_Click()
Principal.Enabled = True
Principal.SSFrame3.Visible = False
Unload Arqueo
Set Arqueo = Nothing
End Sub
