VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form Asignacion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11175
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   14715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11175
   ScaleWidth      =   14715
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame SSFrame1 
      Height          =   3480
      Left            =   2610
      TabIndex        =   4
      Top             =   6705
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   6138
      _Version        =   196608
      BackStyle       =   1
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Última ficha asignada"
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
         Left            =   225
         TabIndex        =   7
         Top             =   2250
         Width           =   1905
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   780
         Left            =   90
         TabIndex        =   6
         Top             =   2475
         Width           =   2040
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Height          =   870
         Left            =   1395
         TabIndex        =   5
         Top             =   405
         Width           =   780
      End
      Begin VB.Image Image1 
         Height          =   1005
         Left            =   180
         Picture         =   "Asignacion.frx":0000
         Stretch         =   -1  'True
         Top             =   315
         Width           =   870
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3240
      Top             =   4275
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=MedicinaT"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "MedicinaT"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "cagisa"
      RecordSource    =   "Select * from nuevos Where NueEst = 1 AND  ProEsLa = -1"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid1 
      Bindings        =   "Asignacion.frx":31A9
      Height          =   4800
      Left            =   315
      TabIndex        =   0
      Top             =   945
      Width           =   12090
      _Version        =   196616
      BevelColorFrame =   14737632
      BevelColorHighlight=   14737632
      BevelColorFace  =   14737632
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      ForeColorEven   =   0
      BackColorEven   =   12632256
      BackColorOdd    =   12632256
      RowHeight       =   423
      Columns.Count   =   7
      Columns(0).Width=   1905
      Columns(0).Caption=   "Fecha Exa."
      Columns(0).Name =   "NueFeE"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "NueFeP"
      Columns(0).DataType=   7
      Columns(0).NumberFormat=   "dd/mm/yyyy"
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   6244
      Columns(1).Caption=   "Nombre"
      Columns(1).Name =   "NueNom"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "NueNom"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   3200
      Columns(2).Caption=   "Teléfonos"
      Columns(2).Name =   "NueTel"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "NueTeI"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   5953
      Columns(3).Caption=   "Empresa"
      Columns(3).Name =   "empdes"
      Columns(3).CaptionAlignment=   0
      Columns(3).DataField=   "empdes"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   1032
      Columns(4).Caption=   "Emb"
      Columns(4).Name =   "NueEmb"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   1
      Columns(4).DataField=   "NueEmb"
      Columns(4).DataType=   3
      Columns(4).FieldLen=   256
      Columns(4).Style=   2
      Columns(5).Width=   1005
      Columns(5).Caption=   "Rx"
      Columns(5).Name =   "ProEsRx"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "ProEsRx"
      Columns(5).DataType=   3
      Columns(5).FieldLen=   256
      Columns(5).Style=   2
      Columns(6).Width=   979
      Columns(6).Caption=   "Lab"
      Columns(6).Name =   "ProEsLa"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "ProEsLa"
      Columns(6).DataType=   3
      Columns(6).FieldLen=   256
      Columns(6).Style=   2
      _ExtentX        =   21325
      _ExtentY        =   8467
      _StockProps     =   79
      Caption         =   "SSOleDBGrid1"
      ForeColor       =   0
      BackColor       =   14737632
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   3480
      Left            =   225
      TabIndex        =   8
      Top             =   6705
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   6138
      _Version        =   196608
      BackStyle       =   1
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   420
         Left            =   405
         TabIndex        =   28
         Top             =   1620
         Width           =   1635
      End
      Begin VB.Image Image3 
         Height          =   1005
         Left            =   180
         Picture         =   "Asignacion.frx":31BE
         Stretch         =   -1  'True
         Top             =   315
         Width           =   870
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Height          =   870
         Left            =   1395
         TabIndex        =   11
         Top             =   405
         Width           =   780
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   780
         Left            =   135
         TabIndex        =   10
         Top             =   2475
         Width           =   2040
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Última ficha asignada"
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
         Left            =   270
         TabIndex        =   9
         Top             =   2250
         Width           =   1905
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   3480
      Left            =   9765
      TabIndex        =   12
      Top             =   6705
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   6138
      _Version        =   196608
      BackStyle       =   1
      Begin VB.Image Image6 
         Height          =   1005
         Left            =   180
         Picture         =   "Asignacion.frx":6367
         Stretch         =   -1  'True
         Top             =   315
         Width           =   870
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Height          =   870
         Left            =   1395
         TabIndex        =   15
         Top             =   405
         Width           =   780
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   780
         Left            =   135
         TabIndex        =   14
         Top             =   2475
         Width           =   2040
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Última ficha asignada"
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
         Left            =   270
         TabIndex        =   13
         Top             =   2250
         Width           =   1905
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   3480
      Left            =   7380
      TabIndex        =   16
      Top             =   6705
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   6138
      _Version        =   196608
      BackStyle       =   1
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Última ficha asignada"
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
         Left            =   270
         TabIndex        =   19
         Top             =   2250
         Width           =   1905
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   780
         Left            =   135
         TabIndex        =   18
         Top             =   2475
         Width           =   2040
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Height          =   870
         Left            =   1395
         TabIndex        =   17
         Top             =   405
         Width           =   780
      End
      Begin VB.Image Image7 
         Height          =   1005
         Left            =   180
         Picture         =   "Asignacion.frx":2B9A3
         Stretch         =   -1  'True
         Top             =   315
         Width           =   870
      End
   End
   Begin Threed.SSFrame SSFrame5 
      Height          =   3480
      Left            =   12150
      TabIndex        =   20
      Top             =   6705
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   6138
      _Version        =   196608
      BackStyle       =   1
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Última ficha asignada"
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
         Left            =   270
         TabIndex        =   23
         Top             =   2205
         Width           =   1905
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   780
         Left            =   135
         TabIndex        =   22
         Top             =   2430
         Width           =   2040
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Height          =   870
         Left            =   1395
         TabIndex        =   21
         Top             =   405
         Width           =   780
      End
      Begin VB.Image Image9 
         Height          =   1005
         Left            =   180
         Picture         =   "Asignacion.frx":50FDF
         Stretch         =   -1  'True
         Top             =   315
         Width           =   870
      End
   End
   Begin Threed.SSFrame SSFrame6 
      Height          =   3480
      Left            =   4995
      TabIndex        =   24
      Top             =   6705
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   6138
      _Version        =   196608
      BackStyle       =   1
      Begin VB.Image Image10 
         Height          =   1005
         Left            =   180
         Picture         =   "Asignacion.frx":7661B
         Stretch         =   -1  'True
         Top             =   315
         Width           =   870
      End
      Begin VB.Label Label21 
         BorderStyle     =   1  'Fixed Single
         Height          =   870
         Left            =   1395
         TabIndex        =   27
         Top             =   405
         Width           =   780
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   780
         Left            =   90
         TabIndex        =   26
         Top             =   2475
         Width           =   2040
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Última ficha asignada"
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
         Left            =   225
         TabIndex        =   25
         Top             =   2250
         Width           =   1905
      End
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   12690
      Picture         =   "Asignacion.frx":797C4
      Stretch         =   -1  'True
      Top             =   5265
      Width           =   285
   End
   Begin VB.Image Image4 
      Height          =   285
      Left            =   12690
      Picture         =   "Asignacion.frx":7BF3B
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   285
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   13095
      TabIndex        =   3
      Top             =   5310
      Width           =   1185
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Asignar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   13095
      TabIndex        =   2
      Top             =   2205
      Width           =   1230
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "ASIGNACION DE TURNO PARA EXAMEN MÉDICO"
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
      Left            =   9180
      TabIndex        =   1
      Top             =   225
      Width           =   5235
   End
   Begin VB.Image Image22 
      Height          =   600
      Left            =   45
      Picture         =   "Asignacion.frx":7DE22
      Stretch         =   -1  'True
      Top             =   45
      Width           =   600
   End
   Begin VB.Image Image8 
      Height          =   465
      Left            =   12555
      Picture         =   "Asignacion.frx":8887F
      Stretch         =   -1  'True
      ToolTipText     =   "Nuevo"
      Top             =   5175
      Width           =   1905
   End
   Begin VB.Image Image2 
      Height          =   465
      Left            =   12555
      Picture         =   "Asignacion.frx":8AFC0
      Stretch         =   -1  'True
      ToolTipText     =   "Nuevo"
      Top             =   2070
      Width           =   1905
   End
   Begin VB.Image Image14 
      Height          =   690
      Left            =   0
      Picture         =   "Asignacion.frx":8D575
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16035
   End
End
Attribute VB_Name = "Asignacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
KeyPreview = True
SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " personas registradas"
Adodc1.Refresh
fichasmedicos
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
Private Sub Label10_Click()
Unload Asignacion
Set Asignacion = Nothing
End Sub

Private Function fichasmedicos()
Dim Cn As New ADODB.Connection
Dim rsf As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open

rsf.CursorType = adOpenKeyset
rsf.LockType = adLockOptimistic
rsf.ActiveConnection = Cn
rsf.Source = "Select * from medicos"
rsf.Open

If Not rsf.EOF Then
    Do While Not rsf.EOF
        
    Loop
End If
End Function
