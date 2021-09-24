VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form Asignacion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10500
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   20325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10500
   ScaleWidth      =   20325
   StartUpPosition =   2  'CenterScreen
   Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo1 
      Height          =   330
      Left            =   16065
      TabIndex        =   49
      Top             =   7830
      Width           =   1725
      _Version        =   65537
      _ExtentX        =   3043
      _ExtentY        =   582
      _StockProps     =   93
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4110
      Left            =   2670
      TabIndex        =   3
      Top             =   6345
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   7250
      _Version        =   196608
      BackStyle       =   1
      Enabled         =   0   'False
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Última ficha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   450
         TabIndex        =   44
         Top             =   1710
         Width           =   1275
      End
      Begin VB.Image Image27 
         Height          =   285
         Left            =   315
         Picture         =   "Asignacion.frx":0000
         Stretch         =   -1  'True
         Top             =   3510
         Width           =   285
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "    Anular"
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
         Left            =   315
         TabIndex        =   38
         Top             =   3555
         Width           =   1725
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   780
         Left            =   90
         TabIndex        =   29
         Top             =   1980
         Width           =   2040
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "    Asignar"
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
         Left            =   270
         TabIndex        =   24
         Top             =   2955
         Width           =   1830
      End
      Begin VB.Image Image13 
         Height          =   285
         Left            =   315
         Picture         =   "Asignacion.frx":309F
         Stretch         =   -1  'True
         Top             =   2955
         Width           =   285
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
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
         Left            =   330
         TabIndex        =   18
         Top             =   1230
         Width           =   1635
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Height          =   870
         Left            =   1395
         TabIndex        =   4
         Top             =   195
         Width           =   780
      End
      Begin VB.Image Image1 
         Height          =   915
         Left            =   180
         Picture         =   "Asignacion.frx":4F86
         Stretch         =   -1  'True
         Top             =   165
         Width           =   810
      End
      Begin VB.Image Image15 
         Height          =   465
         Left            =   180
         Picture         =   "Asignacion.frx":812F
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   2865
         Width           =   1905
      End
      Begin VB.Image Image28 
         Height          =   465
         Left            =   180
         Picture         =   "Asignacion.frx":A6E4
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   3420
         Width           =   1905
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
      RecordSource    =   "Select * from nuevos Where  ProEsLa = -1 AND NueEst = 6 ORDER BY NueFem DESC"
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
      Bindings        =   "Asignacion.frx":D1B3
      Height          =   4800
      Left            =   180
      TabIndex        =   0
      Top             =   720
      Width           =   19980
      _Version        =   196616
      BevelColorFrame =   14737632
      BevelColorHighlight=   14737632
      BevelColorFace  =   14737632
      AllowUpdate     =   0   'False
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
      Columns.Count   =   13
      Columns(0).Width=   1984
      Columns(0).Caption=   "Fec. Exa"
      Columns(0).Name =   "NueFeM"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   1
      Columns(0).DataField=   "NueFeM"
      Columns(0).DataType=   7
      Columns(0).NumberFormat=   "dd/mm/yyyy"
      Columns(0).FieldLen=   256
      Columns(1).Width=   2117
      Columns(1).Caption=   "Fec. Rev."
      Columns(1).Name =   "FecRev"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "FecRev"
      Columns(1).DataType=   7
      Columns(1).NumberFormat=   "dd/mm/yyyy"
      Columns(1).FieldLen=   256
      Columns(2).Width=   6191
      Columns(2).Caption=   "Nombre"
      Columns(2).Name =   "NueNom"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "NueNom"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   1984
      Columns(3).Caption=   "Teléfonos"
      Columns(3).Name =   "NueTel"
      Columns(3).CaptionAlignment=   0
      Columns(3).DataField=   "NueTeI"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   5953
      Columns(4).Caption=   "Empresa"
      Columns(4).Name =   "empdes"
      Columns(4).CaptionAlignment=   0
      Columns(4).DataField=   "empdes"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   1032
      Columns(5).Caption=   "Emb"
      Columns(5).Name =   "NueEmb"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   1
      Columns(5).DataField=   "NueEmb"
      Columns(5).DataType=   3
      Columns(5).FieldLen=   256
      Columns(5).Style=   2
      Columns(6).Width=   1005
      Columns(6).Caption=   "Rx"
      Columns(6).Name =   "ProEsRx"
      Columns(6).Alignment=   1
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "ProEsRx"
      Columns(6).DataType=   3
      Columns(6).FieldLen=   256
      Columns(6).Style=   2
      Columns(7).Width=   2011
      Columns(7).Caption=   "Fec.Rx"
      Columns(7).Name =   "FecRx"
      Columns(7).Alignment=   1
      Columns(7).CaptionAlignment=   1
      Columns(7).DataField=   "FecRx"
      Columns(7).DataType=   7
      Columns(7).NumberFormat=   "dd/mm/yyyy"
      Columns(7).FieldLen=   256
      Columns(8).Width=   979
      Columns(8).Caption=   "Lab"
      Columns(8).Name =   "ProEsLa"
      Columns(8).Alignment=   1
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "ProEsLa"
      Columns(8).DataType=   3
      Columns(8).FieldLen=   256
      Columns(8).Style=   2
      Columns(9).Width=   2593
      Columns(9).Caption=   "Fec.Lab"
      Columns(9).Name =   "FecLab"
      Columns(9).Alignment=   1
      Columns(9).CaptionAlignment=   1
      Columns(9).DataField=   "FecLab"
      Columns(9).DataType=   7
      Columns(9).NumberFormat=   "dd/mm/yyyy"
      Columns(9).FieldLen=   256
      Columns(10).Width=   4339
      Columns(10).Caption=   "Médico"
      Columns(10).Name=   "MedNom"
      Columns(10).CaptionAlignment=   0
      Columns(10).DataField=   "MedNom"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   1005
      Columns(11).Caption=   "Ficha"
      Columns(11).Name=   "NueFic"
      Columns(11).Alignment=   1
      Columns(11).CaptionAlignment=   1
      Columns(11).DataField=   "NueFic"
      Columns(11).DataType=   3
      Columns(11).FieldLen=   256
      Columns(12).Width=   3016
      Columns(12).Caption=   "Tipo"
      Columns(12).Name=   "NueTip"
      Columns(12).CaptionAlignment=   0
      Columns(12).DataField=   "NueTip"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      _ExtentX        =   35242
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
      Height          =   4110
      Left            =   300
      TabIndex        =   5
      Top             =   6345
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   7250
      _Version        =   196608
      BackStyle       =   1
      Begin VB.Image Image25 
         Height          =   285
         Left            =   345
         Picture         =   "Asignacion.frx":D1C8
         Stretch         =   -1  'True
         Top             =   3510
         Width           =   285
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "   Anular"
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
         Left            =   390
         TabIndex        =   37
         Top             =   3555
         Width           =   1725
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "     Asignar"
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
         Left            =   315
         TabIndex        =   23
         Top             =   2925
         Width           =   1800
      End
      Begin VB.Image Image11 
         Height          =   285
         Left            =   360
         Picture         =   "Asignacion.frx":10267
         Stretch         =   -1  'True
         Top             =   2955
         Width           =   285
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
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
         TabIndex        =   17
         Top             =   1230
         Width           =   1635
      End
      Begin VB.Image Image3 
         Height          =   915
         Left            =   180
         Picture         =   "Asignacion.frx":1214E
         Stretch         =   -1  'True
         Top             =   165
         Width           =   810
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Height          =   870
         Left            =   1395
         TabIndex        =   8
         Top             =   195
         Width           =   780
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   780
         Left            =   135
         TabIndex        =   7
         Top             =   1965
         Width           =   2040
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Última ficha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   495
         TabIndex        =   6
         Top             =   1695
         Width           =   1275
      End
      Begin VB.Image Image12 
         Height          =   465
         Left            =   225
         Picture         =   "Asignacion.frx":152F7
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   2865
         Width           =   1905
      End
      Begin VB.Image Image26 
         Height          =   465
         Left            =   210
         Picture         =   "Asignacion.frx":178AC
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   3420
         Width           =   1905
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   4110
      Left            =   9825
      TabIndex        =   9
      Top             =   6345
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   7250
      _Version        =   196608
      BackStyle       =   1
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Última ficha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   630
         TabIndex        =   47
         Top             =   1710
         Width           =   1275
      End
      Begin VB.Image Image33 
         Height          =   285
         Left            =   345
         Picture         =   "Asignacion.frx":1A37B
         Stretch         =   -1  'True
         Top             =   3510
         Width           =   285
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Anular"
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
         Left            =   390
         TabIndex        =   41
         Top             =   3555
         Width           =   1815
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   780
         Left            =   120
         TabIndex        =   32
         Top             =   1980
         Width           =   2040
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "      Asignar"
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
         Left            =   315
         TabIndex        =   27
         Top             =   2955
         Width           =   1800
      End
      Begin VB.Image Image20 
         Height          =   285
         Left            =   360
         Picture         =   "Asignacion.frx":1D41A
         Stretch         =   -1  'True
         Top             =   2955
         Width           =   285
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
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
         Left            =   375
         TabIndex        =   21
         Top             =   1230
         Width           =   1635
      End
      Begin VB.Image Image6 
         Height          =   915
         Left            =   180
         Picture         =   "Asignacion.frx":1F301
         Stretch         =   -1  'True
         Top             =   165
         Width           =   810
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Height          =   870
         Left            =   1395
         TabIndex        =   10
         Top             =   195
         Width           =   780
      End
      Begin VB.Image Image21 
         Height          =   465
         Left            =   225
         Picture         =   "Asignacion.frx":4493D
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   2865
         Width           =   1905
      End
      Begin VB.Image Image34 
         Height          =   465
         Left            =   210
         Picture         =   "Asignacion.frx":46EF2
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   3420
         Width           =   1905
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   4110
      Left            =   7425
      TabIndex        =   11
      Top             =   6345
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   7250
      _Version        =   196608
      BackStyle       =   1
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Última ficha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   540
         TabIndex        =   46
         Top             =   1710
         Width           =   1275
      End
      Begin VB.Image Image31 
         Height          =   285
         Left            =   375
         Picture         =   "Asignacion.frx":499C1
         Stretch         =   -1  'True
         Top             =   3510
         Width           =   285
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "    Anular"
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
         Left            =   330
         TabIndex        =   40
         Top             =   3555
         Width           =   1860
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   780
         Left            =   150
         TabIndex        =   31
         Top             =   1980
         Width           =   2040
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "      Asignar"
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
         Left            =   315
         TabIndex        =   26
         Top             =   2955
         Width           =   1770
      End
      Begin VB.Image Image18 
         Height          =   285
         Left            =   360
         Picture         =   "Asignacion.frx":4CA60
         Stretch         =   -1  'True
         Top             =   2955
         Width           =   285
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
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
         Left            =   390
         TabIndex        =   20
         Top             =   1230
         Width           =   1635
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Height          =   870
         Left            =   1395
         TabIndex        =   12
         Top             =   195
         Width           =   780
      End
      Begin VB.Image Image7 
         Height          =   915
         Left            =   180
         Picture         =   "Asignacion.frx":4E947
         Stretch         =   -1  'True
         Top             =   165
         Width           =   810
      End
      Begin VB.Image Image19 
         Height          =   465
         Left            =   225
         Picture         =   "Asignacion.frx":73F83
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   2865
         Width           =   1905
      End
      Begin VB.Image Image32 
         Height          =   465
         Left            =   240
         Picture         =   "Asignacion.frx":76538
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   3420
         Width           =   1905
      End
   End
   Begin Threed.SSFrame SSFrame5 
      Height          =   4110
      Left            =   12210
      TabIndex        =   13
      Top             =   6345
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   7250
      _Version        =   196608
      BackStyle       =   1
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Última ficha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   585
         TabIndex        =   48
         Top             =   1710
         Width           =   1275
      End
      Begin VB.Image Image35 
         Height          =   285
         Left            =   345
         Picture         =   "Asignacion.frx":79007
         Stretch         =   -1  'True
         Top             =   3510
         Width           =   285
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "    Anular"
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
         Left            =   345
         TabIndex        =   42
         Top             =   3555
         Width           =   1725
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   780
         Left            =   120
         TabIndex        =   33
         Top             =   1980
         Width           =   2040
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "       Asignar"
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
         Left            =   240
         TabIndex        =   28
         Top             =   2955
         Width           =   1830
      End
      Begin VB.Image Image23 
         Height          =   285
         Left            =   360
         Picture         =   "Asignacion.frx":7C0A6
         Stretch         =   -1  'True
         Top             =   2955
         Width           =   285
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
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
         Left            =   360
         TabIndex        =   22
         Top             =   1230
         Width           =   1635
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Height          =   870
         Left            =   1395
         TabIndex        =   14
         Top             =   195
         Width           =   780
      End
      Begin VB.Image Image9 
         Height          =   915
         Left            =   180
         Picture         =   "Asignacion.frx":7DF8D
         Stretch         =   -1  'True
         Top             =   165
         Width           =   810
      End
      Begin VB.Image Image24 
         Height          =   465
         Left            =   225
         Picture         =   "Asignacion.frx":A35C9
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   2865
         Width           =   1905
      End
      Begin VB.Image Image36 
         Height          =   465
         Left            =   210
         Picture         =   "Asignacion.frx":A5B7E
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   3420
         Width           =   1905
      End
   End
   Begin Threed.SSFrame SSFrame6 
      Height          =   4110
      Left            =   5055
      TabIndex        =   15
      Top             =   6345
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   7250
      _Version        =   196608
      BackStyle       =   1
      Enabled         =   0   'False
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Última ficha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   540
         TabIndex        =   45
         Top             =   1710
         Width           =   1275
      End
      Begin VB.Image Image29 
         Height          =   285
         Left            =   315
         Picture         =   "Asignacion.frx":A864D
         Stretch         =   -1  'True
         Top             =   3510
         Width           =   285
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "   Anular"
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
         Left            =   270
         TabIndex        =   39
         Top             =   3555
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   780
         Left            =   120
         TabIndex        =   30
         Top             =   1980
         Width           =   2040
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "       Asignar"
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
         Left            =   210
         TabIndex        =   25
         Top             =   2925
         Width           =   1830
      End
      Begin VB.Image Image16 
         Height          =   285
         Left            =   315
         Picture         =   "Asignacion.frx":AB6EC
         Stretch         =   -1  'True
         Top             =   2955
         Width           =   285
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
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
         Left            =   330
         TabIndex        =   19
         Top             =   1230
         Width           =   1635
      End
      Begin VB.Image Image10 
         Height          =   915
         Left            =   180
         Picture         =   "Asignacion.frx":AD5D3
         Stretch         =   -1  'True
         Top             =   165
         Width           =   810
      End
      Begin VB.Label Label21 
         BorderStyle     =   1  'Fixed Single
         Height          =   870
         Left            =   1395
         TabIndex        =   16
         Top             =   195
         Width           =   780
      End
      Begin VB.Image Image17 
         Height          =   465
         Left            =   180
         Picture         =   "Asignacion.frx":B077C
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   2865
         Width           =   1905
      End
      Begin VB.Image Image30 
         Height          =   465
         Left            =   180
         Picture         =   "Asignacion.frx":B2D31
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   3420
         Width           =   1905
      End
   End
   Begin Threed.SSFrame SSFrame7 
      Height          =   735
      Left            =   270
      TabIndex        =   34
      Top             =   5580
      Width           =   14250
      _ExtentX        =   25135
      _ExtentY        =   1296
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   90
         TabIndex        =   35
         Top             =   180
         Width           =   7080
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "    Buscar"
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
         Height          =   330
         Left            =   7470
         TabIndex        =   36
         Top             =   225
         Width           =   1770
      End
      Begin VB.Image Image2 
         Height          =   285
         Left            =   7470
         Picture         =   "Asignacion.frx":B5800
         Stretch         =   -1  'True
         Top             =   225
         Width           =   285
      End
      Begin VB.Image Image4 
         Height          =   465
         Left            =   7335
         Picture         =   "Asignacion.frx":C5069
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   135
         Width           =   1905
      End
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Trabajo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   330
      Left            =   15975
      TabIndex        =   50
      Top             =   7515
      Width           =   1905
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "   Reiniciar   Fichas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   465
      Left            =   16290
      TabIndex        =   43
      Top             =   6705
      Width           =   1635
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image37 
      Height          =   285
      Left            =   16110
      Picture         =   "Asignacion.frx":C7AA1
      Stretch         =   -1  'True
      Top             =   6795
      Width           =   285
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   16155
      Picture         =   "Asignacion.frx":CB183
      Stretch         =   -1  'True
      Top             =   9135
      Width           =   285
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "    Salir"
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
      Left            =   16110
      TabIndex        =   2
      Top             =   9180
      Width           =   1770
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
      Left            =   13680
      TabIndex        =   1
      Top             =   225
      Width           =   5235
   End
   Begin VB.Image Image22 
      Height          =   600
      Left            =   45
      Picture         =   "Asignacion.frx":CD8FA
      Stretch         =   -1  'True
      Top             =   45
      Width           =   600
   End
   Begin VB.Image Image8 
      Height          =   510
      Left            =   16020
      Picture         =   "Asignacion.frx":D8357
      Stretch         =   -1  'True
      ToolTipText     =   "Nuevo"
      Top             =   9045
      Width           =   1905
   End
   Begin VB.Image Image14 
      Height          =   690
      Left            =   0
      Picture         =   "Asignacion.frx":DAA98
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20325
   End
   Begin VB.Image Image38 
      Height          =   555
      Left            =   15975
      Picture         =   "Asignacion.frx":F7C02
      Stretch         =   -1  'True
      ToolTipText     =   "Nuevo"
      Top             =   6660
      Width           =   1905
   End
End
Attribute VB_Name = "Asignacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vFecRev As String
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

Private Sub Image11_Click()
Label28_Click
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
    SSFrame2.Caption = rsf!MedNom
    Label22.Caption = "C - " & rsf!MedCon
    If rsf!MedCol = "VERDE" Then
        Label5.BackColor = &H8000&      'VERDE
        Label13.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label13.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "ROSADO" Then
        Label5.BackColor = &HFF80FF     'ROSADO
        Label2.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label2.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "CELESTE" Then
        Label5.BackColor = &HFFFF80     'CELESTE
        Label20.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label20.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "ROJO" Then
        Label5.BackColor = &HC0&        'ROJO
        Label7.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label7.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "NARANJA" Then
        Label5.BackColor = &H80FF&      'NARANJA
        Label4.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label4.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "AMARILLO" Then
        Label5.BackColor = &HFFFF&          'AMARILLO
        Label6.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label6.ForeColor = &HC0&
        End If
    End If
    rsf.MoveNext
    
    SSFrame1.Caption = rsf!MedNom
    Label23.Caption = "C - " & rsf!MedCon
    If rsf!MedCol = "VERDE" Then
        Label1.BackColor = &H8000&      'VERDE
        Label13.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label13.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "ROSADO" Then
        Label1.BackColor = &HFF80FF     'ROSADO
        Label2.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label2.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "CELESTE" Then
        Label1.BackColor = &HFFFF80     'CELESTE
        Label20.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label20.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "ROJO" Then
        Label1.BackColor = &HC0&        'ROJO
        Label7.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label7.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "NARANJA" Then
        Label1.BackColor = &H80FF&      'NARANJA
        Label4.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label4.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "AMARILLO" Then
        Label1.BackColor = &HFFFF&          'AMARILLO
        Label6.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label6.ForeColor = &HC0&
        End If
    End If
    rsf.MoveNext
    SSFrame6.Caption = rsf!MedNom
    Label24.Caption = "C - " & rsf!MedCon
    If rsf!MedCol = "VERDE" Then
        Label21.BackColor = &H8000&      'VERDE
        Label13.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label13.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "ROSADO" Then
        Label21.BackColor = &HFF80FF     'ROSADO
        Label2.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label2.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "CELESTE" Then
        Label21.BackColor = &HFFFF80     'CELESTE
        Label20.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label20.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "ROJO" Then
        Label21.BackColor = &HC0&        'ROJO
        Label7.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label7.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "NARANJA" Then
        Label21.BackColor = &H80FF&      'NARANJA
        Label4.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label4.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "AMARILLO" Then
        Label21.BackColor = &HFFFF&          'AMARILLO
        Label6.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label6.ForeColor = &HC0&
        End If
    End If
    rsf.MoveNext
    SSFrame4.Caption = rsf!MedNom
    Label25.Caption = "C - " & rsf!MedCon
    If rsf!MedCol = "VERDE" Then
        Label12.BackColor = &H8000&      'VERDE
        Label13.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label13.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "ROSADO" Then
        Label12.BackColor = &HFF80FF     'ROSADO
        Label2.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label2.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "CELESTE" Then
        Label12.BackColor = &HFFFF80     'CELESTE
        Label20.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label20.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "ROJO" Then
        Label12.BackColor = &HC0&        'ROJO
        Label7.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label7.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "NARANJA" Then
        Label12.BackColor = &H80FF&      'NARANJA
        Label4.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label4.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "AMARILLO" Then
        Label12.BackColor = &HFFFF&          'AMARILLO
        Label6.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label6.ForeColor = &HC0&
        End If
    End If
    rsf.MoveNext
    
    SSFrame3.Caption = rsf!MedNom
    Label26.Caption = "C - " & rsf!MedCon
    If rsf!MedCol = "VERDE" Then
        Label11.BackColor = &H8000&      'VERDE
        Label13.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label13.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "ROSADO" Then
        Label11.BackColor = &HFF80FF     'ROSADO
        Label2.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label2.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "CELESTE" Then
        Label11.BackColor = &HFFFF80     'CELESTE
        Label20.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label20.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "ROJO" Then
        Label11.BackColor = &HC0&        'ROJO
        Label7.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label7.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "NARANJA" Then
        Label11.BackColor = &H80FF&      'NARANJA
        Label4.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label4.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "AMARILLO" Then
        Label11.BackColor = &HFFFF&          'AMARILLO
        Label16.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label16.ForeColor = &HC0&
        End If
    End If
    rsf.MoveNext
    SSFrame5.Caption = rsf!MedNom
    Label27.Caption = "C - " & rsf!MedCon
    If rsf!MedCol = "VERDE" Then
        Label15.BackColor = &H8000&      'VERDE
        Label13.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label13.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "ROSADO" Then
        Label15.BackColor = &HFF80FF     'ROSADO
        Label2.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label2.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "CELESTE" Then
        Label15.BackColor = &HFFFF80     'CELESTE
        Label20.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label20.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "ROJO" Then
        Label15.BackColor = &HC0&        'ROJO
        Label7.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label7.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "NARANJA" Then
        Label15.BackColor = &H80FF&      'NARANJA
        Label4.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label4.ForeColor = &HC0&
        End If
    ElseIf rsf!MedCol = "AMARILLO" Then
        Label15.BackColor = &HFFFF&          'AMARILLO
        Label6.Caption = rsf!MedFiA
        If rsf!MedFiA > 11 Then
            Label6.ForeColor = &HC0&
        End If
    End If
    rsf.MoveNext
End If
End Function

Private Sub Label28_Click()
If seleccion = 1 Then
    vFecRev = Format(SSDateCombo1.Text, "yyyy-mm-dd")
    Dim Cn As New ADODB.Connection
    Dim rsf As New ADODB.Recordset
    Cn.ConnectionString = Cadena
    Cn.Open
    If Val(Label4.Caption) < 11 Then
        If Adodc1.Recordset.Fields("NueFic") = 0 Then
            vMedNom = SSFrame2.Caption
            vnueid = Adodc1.Recordset.Fields("NueId")
            Label4.Caption = Val(Label4.Caption) + 1
            vNueFic = Val(Label4.Caption)
            
            'Actuliza tabla nuevos
            otorga = "UPDATE nuevos SET MedNom = " & "'" & vMedNom & "', NueFic = " & vNueFic & ", FecRev = " & "'" & vFecRev & "', NueEst = " & 6 & " WHERE NueId = " & vnueid
            Cn.Execute otorga
            
            'Actualiza Tabla Médicos
            actmed = "UPDATE medicos SET MedFiA = " & vNueFic & " WHERE MedNom = " & "'" & vMedNom & "'"
            Cn.Execute actmed
            
            Adodc1.Refresh
            seleccion = 0
        Else
            MsgBox "Persona con ficha ya asignada", vbInformation, empresa
        End If
    Else
        If MsgBox("No existen fichas disponibles. Desea asignar una ficha mas?", vbYesNo, empresa) = vbYes Then
            If Adodc1.Recordset.Fields("NueFic") = 0 Then
                vMedNom = SSFrame2.Caption
                vnueid = Adodc1.Recordset.Fields("NueId")
                Label4.Caption = Val(Label4.Caption) + 1
                vNueFic = Val(Label4.Caption)
                
                'Actuliza tabla nuevos
                otorga = "UPDATE nuevos SET MedNom = " & "'" & vMedNom & "', NueFic = " & vNueFic & ", FecRev = " & "'" & vFecRev & "', NueEst = " & 6 & " WHERE NueId = " & vnueid
                Cn.Execute otorga
                
                'Actualiza Tabla Médicos
                actmed = "UPDATE medicos SET MedFiA = " & vNueFic & " WHERE MedNom = " & "'" & vMedNom & "'"
                Cn.Execute actmed
                Label4.ForeColor = &H40C0&
                Adodc1.Refresh
                seleccion = 0
            Else
                MsgBox "Persona con ficha ya asignada", vbInformation, empresa
            End If
        End If
    End If
Else
    MsgBox "Debe seleccionar una persona de la lista", vbInformation, empresa
End If

End Sub

Private Sub Label29_Click()
If seleccion = 1 Then
    vFecRev = Format(SSDateCombo1.Text, "yyyy-mm-dd")
    Dim Cn As New ADODB.Connection
    Dim rsf As New ADODB.Recordset
    Cn.ConnectionString = Cadena
    Cn.Open
    If Val(Label2.Caption) < 11 Then
        If Adodc1.Recordset.Fields("NueFic") = 0 Then
            vMedNom = SSFrame1.Caption
            vnueid = Adodc1.Recordset.Fields("NueId")
            Label2.Caption = Val(Label2.Caption) + 1
            vNueFic = Val(Label2.Caption)
            
            'Actuliza tabla nuevos
            otorga = "UPDATE nuevos SET MedNom = " & "'" & vMedNom & "', NueFic = " & vNueFic & ", FecRev = " & "'" & vFecRev & "', NueEst = " & 6 & " WHERE NueId = " & vnueid
            Cn.Execute otorga
            
            'Actualiza Tabla Médicos
            actmed = "UPDATE medicos SET MedFiA = " & vNueFic & " WHERE MedNom = " & "'" & vMedNom & "'"
            Cn.Execute actmed
            
            Adodc1.Refresh
            seleccion = 0
        Else
            MsgBox "Persona con ficha ya asignada", vbInformation, empresa
        End If
    Else
        If MsgBox("No existen fichas disponibles. Desea asignar una ficha mas?", vbYesNo, empresa) = vbYes Then
            If Adodc1.Recordset.Fields("NueFic") = 0 Then
                vMedNom = SSFrame1.Caption
                vnueid = Adodc1.Recordset.Fields("NueId")
                Label2.Caption = Val(Label2.Caption) + 1
                vNueFic = Val(Label2.Caption)
                
                'Actuliza tabla nuevos
                otorga = "UPDATE nuevos SET MedNom = " & "'" & vMedNom & "', NueFic = " & vNueFic & ", FecRev = " & "'" & vFecRev & "', NueEst = " & 6 & " WHERE NueId = " & vnueid
                Cn.Execute otorga
                
                'Actualiza Tabla Médicos
                actmed = "UPDATE medicos SET MedFiA = " & vNueFic & " WHERE MedNom = " & "'" & vMedNom & "'"
                Cn.Execute actmed
                Label2.ForeColor = &H40C0&
                Adodc1.Refresh
                seleccion = 0
            Else
                MsgBox "Persona con ficha ya asignada", vbInformation, empresa
            End If
        End If
    End If
Else
    MsgBox "Debe seleccionar una persona de la lista", vbInformation, empresa
End If


End Sub

Private Sub Label30_Click()
If seleccion = 1 Then
    vFecRev = Format(SSDateCombo1.Text, "yyyy-mm-dd")
    Dim Cn As New ADODB.Connection
    Dim rsf As New ADODB.Recordset
    Cn.ConnectionString = Cadena
    Cn.Open
    If Val(Label7.Caption) < 11 Then
        If Adodc1.Recordset.Fields("NueFic") = 0 Then
            vMedNom = SSFrame6.Caption
            vnueid = Adodc1.Recordset.Fields("NueId")
            Label7.Caption = Val(Label7.Caption) + 1
            vNueFic = Val(Label7.Caption)
            
            'Actuliza tabla nuevos
            otorga = "UPDATE nuevos SET MedNom = " & "'" & vMedNom & "', NueFic = " & vNueFic & ", FecRev = " & "'" & vFecRev & "', NueEst = " & 6 & " WHERE NueId = " & vnueid
            Cn.Execute otorga
            
            'Actualiza Tabla Médicos
            actmed = "UPDATE medicos SET MedFiA = " & vNueFic & " WHERE MedNom = " & "'" & vMedNom & "'"
            Cn.Execute actmed
            
            Adodc1.Refresh
            seleccion = 0
        Else
            MsgBox "Persona con ficha ya asignada", vbInformation, empresa
        End If
    Else
        If MsgBox("No existen fichas disponibles. Desea asignar una ficha mas?", vbYesNo, empresa) = vbYes Then
            If Adodc1.Recordset.Fields("NueFic") = 0 Then
                vMedNom = SSFrame6.Caption
                vnueid = Adodc1.Recordset.Fields("NueId")
                Label7.Caption = Val(Label7.Caption) + 1
                vNueFic = Val(Label7.Caption)
                
                'Actuliza tabla nuevos
                otorga = "UPDATE nuevos SET MedNom = " & "'" & vMedNom & "', NueFic = " & vNueFic & ", FecRev = " & "'" & vFecRev & "', NueEst = " & 6 & " WHERE NueId = " & vnueid
                Cn.Execute otorga
                
                'Actualiza Tabla Médicos
                actmed = "UPDATE medicos SET MedFiA = " & vNueFic & " WHERE MedNom = " & "'" & vMedNom & "'"
                Cn.Execute actmed
                Label7.ForeColor = &H40C0&
                Adodc1.Refresh
                seleccion = 0
            Else
                MsgBox "Persona con ficha ya asignada", vbInformation, empresa
            End If
        End If
    End If
Else
    MsgBox "Debe seleccionar una persona de la lista", vbInformation, empresa
End If


End Sub

Private Sub Label31_Click()
If seleccion = 1 Then
    vFecRev = Format(SSDateCombo1.Text, "yyyy-mm-dd")
    Dim Cn As New ADODB.Connection
    Dim rsf As New ADODB.Recordset
    Cn.ConnectionString = Cadena
    Cn.Open
    If Val(Label13.Caption) < 11 Then
        If Adodc1.Recordset.Fields("NueFic") = 0 Then
            vMedNom = SSFrame4.Caption
            vnueid = Adodc1.Recordset.Fields("NueId")
            Label13.Caption = Val(Label13.Caption) + 1
            vNueFic = Val(Label13.Caption)
            
            'Actuliza tabla nuevos
            otorga = "UPDATE nuevos SET MedNom = " & "'" & vMedNom & "', NueFic = " & vNueFic & ", FecRev = " & "'" & vFecRev & "', NueEst = " & 6 & " WHERE NueId = " & vnueid
            Cn.Execute otorga
            
            'Actualiza Tabla Médicos
            actmed = "UPDATE medicos SET MedFiA = " & vNueFic & " WHERE MedNom = " & "'" & vMedNom & "'"
            Cn.Execute actmed
            
            Adodc1.Refresh
            seleccion = 0
        Else
            MsgBox "Persona con ficha ya asignada", vbInformation, empresa
        End If
    Else
        If MsgBox("No existen fichas disponibles. Desea asignar una ficha mas?", vbYesNo, empresa) = vbYes Then
            If Adodc1.Recordset.Fields("NueFic") = 0 Then
                vMedNom = SSFrame4.Caption
                vnueid = Adodc1.Recordset.Fields("NueId")
                Label13.Caption = Val(Label13.Caption) + 1
                vNueFic = Val(Label13.Caption)
                
                'Actuliza tabla nuevos
                otorga = "UPDATE nuevos SET MedNom = " & "'" & vMedNom & "', NueFic = " & vNueFic & ", FecRev = " & "'" & vFecRev & "', NueEst = " & 6 & " WHERE NueId = " & vnueid
                Cn.Execute otorga
                
                'Actualiza Tabla Médicos
                actmed = "UPDATE medicos SET MedFiA = " & vNueFic & " WHERE MedNom = " & "'" & vMedNom & "'"
                Cn.Execute actmed
                Label13.ForeColor = &H40C0&
                Adodc1.Refresh
                seleccion = 0
            Else
                MsgBox "Persona con ficha ya asignada", vbInformation, empresa
            End If
        End If
    End If
Else
    MsgBox "Debe seleccionar una persona de la lista", vbInformation, empresa
End If

End Sub

Private Sub Label32_Click()
If seleccion = 1 Then
    vFecRev = Format(SSDateCombo1.Text, "yyyy-mm-dd")
    Dim Cn As New ADODB.Connection
    Dim rsf As New ADODB.Recordset
    Cn.ConnectionString = Cadena
    Cn.Open
    If Val(Label16.Caption) < 11 Then
        If Adodc1.Recordset.Fields("NueFic") = 0 Then
            vMedNom = SSFrame3.Caption
            vnueid = Adodc1.Recordset.Fields("NueId")
            Label16.Caption = Val(Label16.Caption) + 1
            vNueFic = Val(Label16.Caption)
            
            'Actuliza tabla nuevos
            otorga = "UPDATE nuevos SET MedNom = " & "'" & vMedNom & "', NueFic = " & vNueFic & ", FecRev = " & "'" & vFecRev & "', NueEst = " & 6 & " WHERE NueId = " & vnueid
            Cn.Execute otorga
            
            'Actualiza Tabla Médicos
            actmed = "UPDATE medicos SET MedFiA = " & vNueFic & " WHERE MedNom = " & "'" & vMedNom & "'"
            Cn.Execute actmed
            
            Adodc1.Refresh
            seleccion = 0
        Else
            MsgBox "Persona con ficha ya asignada", vbInformation, empresa
        End If
    Else
        If MsgBox("No existen fichas disponibles. Desea asignar una ficha mas?", vbYesNo, empresa) = vbYes Then
            If Adodc1.Recordset.Fields("NueFic") = 0 Then
                vMedNom = SSFrame3.Caption
                vnueid = Adodc1.Recordset.Fields("NueId")
                Label16.Caption = Val(Label16.Caption) + 1
                vNueFic = Val(Label16.Caption)
                
                'Actuliza tabla nuevos
                otorga = "UPDATE nuevos SET MedNom = " & "'" & vMedNom & "', NueFic = " & vNueFic & ", FecRev = " & "'" & vFecRev & "', NueEst = " & 6 & " WHERE NueId = " & vnueid
                Cn.Execute otorga
                
                'Actualiza Tabla Médicos
                actmed = "UPDATE medicos SET MedFiA = " & vNueFic & " WHERE MedNom = " & "'" & vMedNom & "'"
                Cn.Execute actmed
                Label16.ForeColor = &H40C0&
                Adodc1.Refresh
                seleccion = 0
            Else
                MsgBox "Persona con ficha ya asignada", vbInformation, empresa
            End If
        End If
    End If
Else
    MsgBox "Debe seleccionar una persona de la lista", vbInformation, empresa
End If


End Sub

Private Sub Label33_Click()
If seleccion = 1 Then
    vFecRev = Format(SSDateCombo1.Text, "yyyy-mm-dd")
    Dim Cn As New ADODB.Connection
    Dim rsf As New ADODB.Recordset
    Cn.ConnectionString = Cadena
    Cn.Open
    If Val(Label20.Caption) < 11 Then
        If Adodc1.Recordset.Fields("NueFic") = 0 Then
            vMedNom = SSFrame5.Caption
            vnueid = Adodc1.Recordset.Fields("NueId")
            Label20.Caption = Val(Label20.Caption) + 1
            vNueFic = Val(Label20.Caption)
            
            'Actuliza tabla nuevos
            otorga = "UPDATE nuevos SET MedNom = " & "'" & vMedNom & "', NueFic = " & vNueFic & ", FecRev = " & "'" & vFecRev & "', NueEst = " & 6 & " WHERE NueId = " & vnueid
            Cn.Execute otorga
            
            'Actualiza Tabla Médicos
            actmed = "UPDATE medicos SET MedFiA = " & vNueFic & " WHERE MedNom = " & "'" & vMedNom & "'"
            Cn.Execute actmed
            
            Adodc1.Refresh
            seleccion = 0
        Else
            MsgBox "Persona con ficha ya asignada", vbInformation, empresa
        End If
    Else
        If MsgBox("No existen fichas disponibles. Desea asignar una ficha mas?", vbYesNo, empresa) = vbYes Then
            If Adodc1.Recordset.Fields("NueFic") = 0 Then
                vMedNom = SSFrame5.Caption
                vnueid = Adodc1.Recordset.Fields("NueId")
                Label20.Caption = Val(Label20.Caption) + 1
                vNueFic = Val(Label20.Caption)
                
                'Actuliza tabla nuevos
                otorga = "UPDATE nuevos SET MedNom = " & "'" & vMedNom & "', NueFic = " & vNueFic & ", FecRev = " & "'" & vFecRev & "', NueEst = " & 6 & " WHERE NueId = " & vnueid
                Cn.Execute otorga
                
                'Actualiza Tabla Médicos
                actmed = "UPDATE medicos SET MedFiA = " & vNueFic & " WHERE MedNom = " & "'" & vMedNom & "'"
                Cn.Execute actmed
                Label20.ForeColor = &H40C0&
                Adodc1.Refresh
                seleccion = 0
            Else
                MsgBox "Persona con ficha ya asignada", vbInformation, empresa
            End If
        End If
    End If
Else
    MsgBox "Debe seleccionar una persona de la lista", vbInformation, empresa
End If

End Sub

Private Sub Label34_Click()
If seleccion = 1 Then
    If Adodc1.Recordset.Fields("NueFic") > 0 Then
        Dim Cn As New ADODB.Connection
        Dim rsf As New ADODB.Recordset
        Cn.ConnectionString = Cadena
        Cn.Open
    
        vMedNom = SSFrame2.Caption
        vnueid = Adodc1.Recordset.Fields("NueId")
        vNueFic = Val(Label4.Caption)
        
        'Actuliza tabla nuevos
        otorga = "UPDATE nuevos SET MedNom = '', NueFic = " & 0 & ", NueEst = " & 5 & " WHERE NueId = " & vnueid
        Cn.Execute otorga
        
        'Actualiza Tabla Médicos
        If vNueFic = Adodc1.Recordset.Fields("NueFic") Then
            Label4.Caption = Val(Label4.Caption) - 1
            vNueFicAct = Label4.Caption
            actmed = "UPDATE medicos SET MedFiA = " & vNueFicAct & " WHERE MedNom = " & "'" & vMedNom & "', FecRev = '0000-00-00'"
            Cn.Execute actmed
            'Label20.Caption = Val(Label20.Caption) - 1
        End If
        Adodc1.Refresh
        seleccion = 0
    Else
        MsgBox "Persona sin ficha asignada", vbInformation, empresa
    End If
Else
    MsgBox "Debe seleccionar una persona de la lista", vbInformation, empresa
End If
End Sub

Private Sub Label35_Click()
If seleccion = 1 Then
    If Adodc1.Recordset.Fields("NueFic") > 0 Then
        Dim Cn As New ADODB.Connection
        Dim rsf As New ADODB.Recordset
        Cn.ConnectionString = Cadena
        Cn.Open
    
        vMedNom = SSFrame1.Caption
        vnueid = Adodc1.Recordset.Fields("NueId")
        vNueFic = Val(Label2.Caption)
        
        'Actuliza tabla nuevos
        otorga = "UPDATE nuevos SET MedNom = '', NueFic = " & 0 & ", NueEst = " & 5 & " WHERE NueId = " & vnueid
        Cn.Execute otorga
        
        'Actualiza Tabla Médicos
        If vNueFic = Adodc1.Recordset.Fields("NueFic") Then
            Label2.Caption = Val(Label2.Caption) - 1
            vNueFicAct = Label2.Caption
            actmed = "UPDATE medicos SET MedFiA = " & vNueFicAct & " WHERE MedNom = " & "'" & vMedNom & "'"
            Cn.Execute actmed
            'Label20.Caption = Val(Label20.Caption) - 1
        End If
        Adodc1.Refresh
        seleccion = 0
    Else
        MsgBox "Persona sin ficha asignada", vbInformation, empresa
    End If
Else
    MsgBox "Debe seleccionar una persona de la lista", vbInformation, empresa
End If

End Sub

Private Sub Label36_Click()
If seleccion = 1 Then
    If Adodc1.Recordset.Fields("NueFic") > 0 Then
        Dim Cn As New ADODB.Connection
        Dim rsf As New ADODB.Recordset
        Cn.ConnectionString = Cadena
        Cn.Open
    
        vMedNom = SSFrame6.Caption
        vnueid = Adodc1.Recordset.Fields("NueId")
        vNueFic = Val(Label7.Caption)
        
        'Actuliza tabla nuevos
        otorga = "UPDATE nuevos SET MedNom = '', NueFic = " & 0 & ", NueEst = " & 5 & " WHERE NueId = " & vnueid
        Cn.Execute otorga
        
        'Actualiza Tabla Médicos
        If vNueFic = Adodc1.Recordset.Fields("NueFic") Then
            Label7.Caption = Val(Label7.Caption) - 1
            vNueFicAct = Label7.Caption
            actmed = "UPDATE medicos SET MedFiA = " & vNueFicAct & " WHERE MedNom = " & "'" & vMedNom & "'"
            Cn.Execute actmed
            'Label20.Caption = Val(Label20.Caption) - 1
        End If
        Adodc1.Refresh
        seleccion = 0
    Else
        MsgBox "Persona sin ficha asignada", vbInformation, empresa
    End If
Else
    MsgBox "Debe seleccionar una persona de la lista", vbInformation, empresa
End If

End Sub

Private Sub Label37_Click()
If seleccion = 1 Then
    If Adodc1.Recordset.Fields("NueFic") > 0 Then
        Dim Cn As New ADODB.Connection
        Dim rsf As New ADODB.Recordset
        Cn.ConnectionString = Cadena
        Cn.Open
    
        vMedNom = SSFrame4.Caption
        vnueid = Adodc1.Recordset.Fields("NueId")
        vNueFic = Val(Label13.Caption)
        
        'Actuliza tabla nuevos
        otorga = "UPDATE nuevos SET MedNom = '', NueFic = " & 0 & ", NueEst = " & 5 & " WHERE NueId = " & vnueid
        Cn.Execute otorga
        
        'Actualiza Tabla Médicos
        If vNueFic = Adodc1.Recordset.Fields("NueFic") Then
            Label13.Caption = Val(Label13.Caption) - 1
            vNueFicAct = Label13.Caption
            actmed = "UPDATE medicos SET MedFiA = " & vNueFicAct & " WHERE MedNom = " & "'" & vMedNom & "'"
            Cn.Execute actmed
            'Label20.Caption = Val(Label20.Caption) - 1
        End If
        Adodc1.Refresh
        seleccion = 0
    Else
        MsgBox "Persona sin ficha asignada", vbInformation, empresa
    End If
Else
    MsgBox "Debe seleccionar una persona de la lista", vbInformation, empresa
End If

End Sub

Private Sub Label38_Click()
If seleccion = 1 Then
    If Adodc1.Recordset.Fields("NueFic") > 0 Then
        Dim Cn As New ADODB.Connection
        Dim rsf As New ADODB.Recordset
        Cn.ConnectionString = Cadena
        Cn.Open
    
        vMedNom = SSFrame3.Caption
        vnueid = Adodc1.Recordset.Fields("NueId")
        vNueFic = Val(Label16.Caption)
        
        'Actuliza tabla nuevos
        otorga = "UPDATE nuevos SET MedNom = '', NueFic = " & 0 & ", NueEst = " & 5 & " WHERE NueId = " & vnueid
        Cn.Execute otorga
        
        'Actualiza Tabla Médicos
        If vNueFic = Adodc1.Recordset.Fields("NueFic") Then
            Label16.Caption = Val(Label16.Caption) - 1
            vNueFicAct = Label16.Caption
            actmed = "UPDATE medicos SET MedFiA = " & vNueFicAct & " WHERE MedNom = " & "'" & vMedNom & "'"
            Cn.Execute actmed
            'Label20.Caption = Val(Label20.Caption) - 1
        End If
        Adodc1.Refresh
        seleccion = 0
    Else
        MsgBox "Persona sin ficha asignada", vbInformation, empresa
    End If
Else
    MsgBox "Debe seleccionar una persona de la lista", vbInformation, empresa
End If

End Sub

Private Sub Label39_Click()
If seleccion = 1 Then
    If Adodc1.Recordset.Fields("NueFic") > 0 Then
        Dim Cn As New ADODB.Connection
        Dim rsf As New ADODB.Recordset
        Cn.ConnectionString = Cadena
        Cn.Open
    
        vMedNom = SSFrame5.Caption
        vnueid = Adodc1.Recordset.Fields("NueId")
        vNueFic = Val(Label20.Caption)
        
        'Actuliza tabla nuevos
        otorga = "UPDATE nuevos SET MedNom = '', NueFic = " & 0 & ", NueEst = " & 5 & " WHERE NueId = " & vnueid
        Cn.Execute otorga
        
        'Actualiza Tabla Médicos
        If vNueFic = Adodc1.Recordset.Fields("NueFic") Then
            Label20.Caption = Val(Label20.Caption) - 1
            vNueFicAct = Label20.Caption
            actmed = "UPDATE medicos SET MedFiA = " & vNueFicAct & " WHERE MedNom = " & "'" & vMedNom & "'"
            Cn.Execute actmed
            'Label20.Caption = Val(Label20.Caption) - 1
        End If
        Adodc1.Refresh
        seleccion = 0
    Else
        MsgBox "Persona sin ficha asignada", vbInformation, empresa
    End If
Else
    MsgBox "Debe seleccionar una persona de la lista", vbInformation, empresa
End If

End Sub

Private Sub Label40_Click()
Dim Cn As New ADODB.Connection
Cn.ConnectionString = Cadena
Cn.Open
cero = "UPDATE medicos SET MedFiA = " & 0
Cn.Execute cero
fichasmedicos
Label4.ForeColor = &H8000&
Label2.ForeColor = &H8000&
Label7.ForeColor = &H8000&
Label3.ForeColor = &H8000&
Label6.ForeColor = &H8000&
Label20.ForeColor = &H8000&
End Sub

Private Sub Label8_Click()
If Len(Trim(Text8.Text)) > 0 Then
    Dim db1 As String
    db1 = Text8.Text
    Adodc1.RecordSource = "SELECT * from nuevos WHERE  NueNom LIKE " & "'%" & db1 & "%'"
    Adodc1.Refresh
    Text8.Text = ""
    SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " registros encontrados"
Else
    Adodc1.RecordSource = "Select * from nuevos Where NueEst = " & 1 & " AND ProEsLa = " & -1 & " ORDER BY NueFem DESC"
    Adodc1.Refresh
End If

End Sub

Private Sub SSOleDBGrid1_Click()
seleccion = 1
End Sub

