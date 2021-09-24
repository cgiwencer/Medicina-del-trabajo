VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{B074BC93-5A5B-11CE-98BD-0000C0E6B88E}#2.0#0"; "sstabs32.ocx"
Begin VB.Form Programacion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8700
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   17550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   17550
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame SSFrame4 
      Height          =   1275
      Left            =   420
      TabIndex        =   23
      Top             =   7230
      Width           =   13440
      _ExtentX        =   23707
      _ExtentY        =   2249
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo3 
         Height          =   285
         Left            =   11145
         TabIndex        =   4
         Top             =   360
         Width           =   1905
         _Version        =   65537
         _ExtentX        =   3360
         _ExtentY        =   503
         _StockProps     =   93
         Enabled         =   0   'False
         DefaultDate     =   ""
         AllowNullDate   =   -1  'True
      End
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo1 
         Height          =   285
         Left            =   8835
         TabIndex        =   2
         Top             =   360
         Width           =   1905
         _Version        =   65537
         _ExtentX        =   3360
         _ExtentY        =   503
         _StockProps     =   93
         Enabled         =   0   'False
         DefaultDate     =   ""
         AllowNullDate   =   -1  'True
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   375
         Left            =   11685
         TabIndex        =   5
         Top             =   765
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   661
         _Version        =   196608
         ForeColor       =   16777215
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Registrar"
         ButtonStyle     =   4
         BevelWidth      =   0
      End
      Begin Threed.SSCheck SSCheck2 
         Height          =   240
         Left            =   8595
         TabIndex        =   1
         Top             =   360
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   423
         _Version        =   196608
         BackStyle       =   1
      End
      Begin Threed.SSCheck SSCheck1 
         Height          =   240
         Left            =   10890
         TabIndex        =   3
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   423
         _Version        =   196608
         BackStyle       =   1
      End
      Begin VB.Image Image11 
         Height          =   285
         Left            =   11295
         Picture         =   "Programacion.frx":0000
         Stretch         =   -1  'True
         Top             =   810
         Width           =   285
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Laboratorio"
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
         Left            =   8880
         TabIndex        =   29
         Top             =   135
         Width           =   1950
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
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
         Left            =   3900
         TabIndex        =   28
         Top             =   90
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   3870
         TabIndex        =   27
         Top             =   345
         Width           =   3780
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
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
         Left            =   120
         TabIndex        =   26
         Top             =   105
         Width           =   1095
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Rx"
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
         Left            =   11160
         TabIndex        =   25
         Top             =   135
         Width           =   1590
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   3555
      End
      Begin VB.Image Image12 
         Height          =   465
         Left            =   11130
         Picture         =   "Programacion.frx":1D7B
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   720
         Width           =   1905
      End
   End
   Begin SSDesignerWidgetsTabs.SSIndexTab SSIndexTab1 
      Height          =   5280
      Left            =   360
      TabIndex        =   13
      Top             =   945
      Width           =   14550
      _Version        =   131078
      _ExtentX        =   25665
      _ExtentY        =   9313
      _StockProps     =   13
      CoverAllowClose =   0   'False
      CoverMarginX    =   200
      CoverMarginY    =   200
      RingHoleMargin  =   500
      RingMarginTop   =   100
      RingMarginBottom=   100
      RingSeparator   =   200
      RingSize        =   1000
      RingWidth       =   300
      ActualTTO       =   500
      GutterWidth     =   100
      PageAnimationFrames=   20
      BevelColorFace  =   14737632
      ActiveTab3DBackColor=   16777215
      TabVisibleLast  =   4
      TabsPerRow      =   5
      RingCount       =   9
      RingGroups      =   3
      PageTabOrientation=   0
      PageTabsPerRow  =   5
      PageAlignmentCaption=   7
      PageAlignmentPicture=   1
      BeginProperty FontSub {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ActivePageFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabsNum         =   5
      Tabs(IC).PictureMetaWidth=   0
      Tabs(IC).PictureMetaHeight=   0
      Tabs(IC).Page   =   0
      Tabs(IC).ControlCount=   0
      Tabs(IC).ControlEnabled=   0   'False
      Tabs(IC).Pages(0).PictureMetaWidth=   0
      Tabs(IC).Pages(0).PictureMetaHeight=   0
      Tabs(IC).Pages(0).Tag=   ""
      Tabs(IC).Pages(0).Caption=   "Page 0"
      Tabs(IC).Pages(0).Name=   "page 0"
      Tabs(IC).Pages(0).CtlCount=   0
      Tabs(IC).Pages(0).CtlEnabled=   0   'False
      Tabs(IC).Tag    =   ""
      Tabs(IC).Caption=   ""
      Tabs(IC).Name   =   ""
      Tabs(0).PictureMetaWidth=   0
      Tabs(0).PictureMetaHeight=   0
      Tabs(0).Page    =   0
      Tabs(0).ControlCount=   0
      Tabs(0).ControlEnabled=   0   'False
      Tabs(0).Pages(0).PictureMetaWidth=   0
      Tabs(0).Pages(0).PictureMetaHeight=   0
      Tabs(0).Pages(0).Tag=   ""
      Tabs(0).Pages(0).Caption=   "Page 0"
      Tabs(0).Pages(0).Name=   "page 0"
      Tabs(0).Pages(0).CtlCount=   2
      Tabs(0).Pages(0).CtlEnabled=   -1  'True
      Tabs(0).Pages(0).Ctl(0)=   "Adodc1"
      Tabs(0).Pages(0).Ctl(1)=   "SSOleDBGrid1"
      Tabs(0).Tag     =   ""
      Tabs(0).Caption =   "Preocupacional"
      Tabs(0).Name    =   "tab 0"
      Tabs(1).PictureMetaWidth=   0
      Tabs(1).PictureMetaHeight=   0
      Tabs(1).Page    =   0
      Tabs(1).ControlCount=   0
      Tabs(1).ControlEnabled=   0   'False
      Tabs(1).Pages(0).PictureMetaWidth=   0
      Tabs(1).Pages(0).PictureMetaHeight=   0
      Tabs(1).Pages(0).Tag=   ""
      Tabs(1).Pages(0).Caption=   "Page 0"
      Tabs(1).Pages(0).Name=   "page 0"
      Tabs(1).Pages(0).CtlCount=   2
      Tabs(1).Pages(0).CtlEnabled=   0   'False
      Tabs(1).Pages(0).Ctl(0)=   "Adodc2"
      Tabs(1).Pages(0).Ctl(1)=   "SSOleDBGrid2"
      Tabs(1).Tag     =   ""
      Tabs(1).Caption =   "Postocupacional"
      Tabs(1).Name    =   "tab 1"
      Tabs(2).PictureMetaWidth=   0
      Tabs(2).PictureMetaHeight=   0
      Tabs(2).Page    =   0
      Tabs(2).ControlCount=   0
      Tabs(2).ControlEnabled=   0   'False
      Tabs(2).Pages(0).PictureMetaWidth=   0
      Tabs(2).Pages(0).PictureMetaHeight=   0
      Tabs(2).Pages(0).Tag=   ""
      Tabs(2).Pages(0).Caption=   "Page 0"
      Tabs(2).Pages(0).Name=   "page 0"
      Tabs(2).Pages(0).CtlCount=   2
      Tabs(2).Pages(0).CtlEnabled=   0   'False
      Tabs(2).Pages(0).Ctl(0)=   "Adodc3"
      Tabs(2).Pages(0).Ctl(1)=   "SSOleDBGrid3"
      Tabs(2).Tag     =   ""
      Tabs(2).Caption =   "Reprogramación"
      Tabs(2).Name    =   "tab 2"
      Tabs(3).PictureMetaWidth=   0
      Tabs(3).PictureMetaHeight=   0
      Tabs(3).Page    =   0
      Tabs(3).ControlCount=   0
      Tabs(3).ControlEnabled=   0   'False
      Tabs(3).Pages(0).PictureMetaWidth=   0
      Tabs(3).Pages(0).PictureMetaHeight=   0
      Tabs(3).Pages(0).Tag=   ""
      Tabs(3).Pages(0).Caption=   "Page 0"
      Tabs(3).Pages(0).Name=   "page 0"
      Tabs(3).Pages(0).CtlCount=   2
      Tabs(3).Pages(0).CtlEnabled=   0   'False
      Tabs(3).Pages(0).Ctl(0)=   "Adodc4"
      Tabs(3).Pages(0).Ctl(1)=   "SSOleDBGrid4"
      Tabs(3).Tag     =   ""
      Tabs(3).Caption =   "Depurados"
      Tabs(3).Name    =   "tab 3"
      Tabs(4).PictureMetaWidth=   0
      Tabs(4).PictureMetaHeight=   0
      Tabs(4).Page    =   0
      Tabs(4).ControlCount=   0
      Tabs(4).ControlEnabled=   0   'False
      Tabs(4).Pages(0).PictureMetaWidth=   0
      Tabs(4).Pages(0).PictureMetaHeight=   0
      Tabs(4).Pages(0).Tag=   ""
      Tabs(4).Pages(0).Caption=   "Page 0"
      Tabs(4).Pages(0).Name=   "page 0"
      Tabs(4).Pages(0).CtlCount=   2
      Tabs(4).Pages(0).CtlEnabled=   0   'False
      Tabs(4).Pages(0).Ctl(0)=   "SSOleDBGrid5"
      Tabs(4).Pages(0).Ctl(1)=   "Adodc5"
      Tabs(4).Tag     =   ""
      Tabs(4).Caption =   "Sin exámen"
      Tabs(4).Name    =   "tab 4"
      Templates(0).PictureMetaWidth=   0
      Templates(0).PictureMetaHeight=   0
      Templates(0).Tag=   ""
      Templates(0).Caption=   "Page 0"
      Templates(0).Name=   "page 0"
      Templates(0).CtlCount=   0
      Templates(0).CtlEnabled=   -1  'True
      Templates(1).PictureMetaWidth=   0
      Templates(1).PictureMetaHeight=   0
      Templates(1).Tag=   ""
      Templates(1).Caption=   "Page 1"
      Templates(1).Name=   "page 1"
      Templates(1).CtlCount=   0
      Templates(1).CtlEnabled=   0   'False
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   3195
         Top             =   3240
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
         RecordSource    =   "Select * from nuevos Where NueEst = 1 AND NueTip=""PREOCUPACIONAL""  ORDER BY NueFeP DESC"
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
         Bindings        =   "Programacion.frx":4775
         Height          =   4575
         Left            =   150
         TabIndex        =   0
         Top             =   540
         Width           =   14205
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
         Columns.Count   =   9
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
         Columns(2).Width=   2566
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
         Columns(4).Width=   926
         Columns(4).Caption=   "Emb"
         Columns(4).Name =   "NueEmb"
         Columns(4).Alignment=   1
         Columns(4).CaptionAlignment=   1
         Columns(4).DataField=   "NueEmb"
         Columns(4).DataType=   3
         Columns(4).FieldLen=   256
         Columns(4).Locked=   -1  'True
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
         Columns(6).Width=   2249
         Columns(6).Caption=   "Rx"
         Columns(6).Name =   "FecRx"
         Columns(6).Alignment=   1
         Columns(6).CaptionAlignment=   1
         Columns(6).DataField=   "FecRx"
         Columns(6).DataType=   7
         Columns(6).NumberFormat=   "dd/mm/yyyy"
         Columns(6).FieldLen=   256
         Columns(7).Width=   979
         Columns(7).Caption=   "Lab"
         Columns(7).Name =   "ProEsLa"
         Columns(7).Alignment=   1
         Columns(7).CaptionAlignment=   2
         Columns(7).DataField=   "ProEsLa"
         Columns(7).DataType=   3
         Columns(7).FieldLen=   256
         Columns(7).Style=   2
         Columns(8).Width=   2170
         Columns(8).Caption=   "Laboratorio"
         Columns(8).Name =   "FecLab"
         Columns(8).Alignment=   1
         Columns(8).CaptionAlignment=   1
         Columns(8).DataField=   "FecLab"
         Columns(8).DataType=   7
         Columns(8).NumberFormat=   "dd/mm/yyyy"
         Columns(8).FieldLen=   256
         Columns(8).Locked=   -1  'True
         _ExtentX        =   25056
         _ExtentY        =   8070
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   -71865
         Top             =   3240
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
         RecordSource    =   "Select * from nuevos Where NueEst = 1 AND NueTip=""POSTOCUPACIONAL"""
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid2 
         Bindings        =   "Programacion.frx":478A
         Height          =   4575
         Left            =   -74865
         TabIndex        =   14
         Top             =   540
         Width           =   14310
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
         Columns.Count   =   9
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
         Columns(2).Width=   2566
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
         Columns(4).Width=   926
         Columns(4).Caption=   "Emb"
         Columns(4).Name =   "NueEmb"
         Columns(4).Alignment=   1
         Columns(4).CaptionAlignment=   1
         Columns(4).DataField=   "NueEmb"
         Columns(4).DataType=   3
         Columns(4).FieldLen=   256
         Columns(4).Locked=   -1  'True
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
         Columns(6).Width=   2170
         Columns(6).Caption=   "Rx"
         Columns(6).Name =   "FecRx"
         Columns(6).Alignment=   1
         Columns(6).CaptionAlignment=   1
         Columns(6).DataField=   "FecRx"
         Columns(6).DataType=   7
         Columns(6).NumberFormat=   "dd/mm/yyyy"
         Columns(6).FieldLen=   256
         Columns(7).Width=   979
         Columns(7).Caption=   "Lab"
         Columns(7).Name =   "ProEsLa"
         Columns(7).Alignment=   1
         Columns(7).CaptionAlignment=   2
         Columns(7).DataField=   "ProEsLa"
         Columns(7).DataType=   3
         Columns(7).FieldLen=   256
         Columns(7).Style=   2
         Columns(8).Width=   2381
         Columns(8).Caption=   "Laboratorio"
         Columns(8).Name =   "FecLab"
         Columns(8).Alignment=   1
         Columns(8).CaptionAlignment=   1
         Columns(8).DataField=   "FecLab"
         Columns(8).DataType=   7
         Columns(8).NumberFormat=   "dd/mm/yyyy"
         Columns(8).FieldLen=   256
         _ExtentX        =   25241
         _ExtentY        =   8070
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
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   -71865
         Top             =   3150
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
         RecordSource    =   "Select * from nuevos Where NueEst = 1 OR NueTip=""REPROGRAMACION"""
         Caption         =   "Adodc3"
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid3 
         Bindings        =   "Programacion.frx":479F
         Height          =   4575
         Left            =   -74640
         TabIndex        =   15
         Top             =   540
         Width           =   14070
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
         Columns.Count   =   9
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
         Columns(2).Width=   2566
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
         Columns(4).Width=   926
         Columns(4).Caption=   "Emb"
         Columns(4).Name =   "NueEmb"
         Columns(4).Alignment=   1
         Columns(4).CaptionAlignment=   1
         Columns(4).DataField=   "NueEmb"
         Columns(4).DataType=   3
         Columns(4).FieldLen=   256
         Columns(4).Locked=   -1  'True
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
         Columns(6).Width=   2064
         Columns(6).Caption=   "Rx"
         Columns(6).Name =   "FecRx"
         Columns(6).Alignment=   1
         Columns(6).CaptionAlignment=   1
         Columns(6).DataField=   "FecRx"
         Columns(6).DataType=   7
         Columns(6).NumberFormat=   "dd/mm/yyyy"
         Columns(6).FieldLen=   256
         Columns(7).Width=   979
         Columns(7).Caption=   "Lab"
         Columns(7).Name =   "ProEsLa"
         Columns(7).Alignment=   1
         Columns(7).CaptionAlignment=   2
         Columns(7).DataField=   "ProEsLa"
         Columns(7).DataType=   3
         Columns(7).FieldLen=   256
         Columns(7).Style=   2
         Columns(8).Width=   2143
         Columns(8).Caption=   "Laboratorio"
         Columns(8).Name =   "FecLab"
         Columns(8).Alignment=   1
         Columns(8).CaptionAlignment=   1
         Columns(8).DataField=   "FecLab"
         Columns(8).DataType=   7
         Columns(8).NumberFormat=   "dd/mm/yyyy"
         Columns(8).FieldLen=   256
         _ExtentX        =   24818
         _ExtentY        =   8070
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
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   330
         Left            =   -71910
         Top             =   3090
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
         RecordSource    =   "Select * from nuevos Where NueEst = 3"
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid4 
         Bindings        =   "Programacion.frx":47B4
         Height          =   4575
         Left            =   -74910
         TabIndex        =   16
         Top             =   540
         Width           =   14295
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
         Columns.Count   =   9
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
         Columns(2).Width=   2566
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
         Columns(4).Width=   926
         Columns(4).Caption=   "Emb"
         Columns(4).Name =   "NueEmb"
         Columns(4).Alignment=   1
         Columns(4).CaptionAlignment=   1
         Columns(4).DataField=   "NueEmb"
         Columns(4).DataType=   3
         Columns(4).FieldLen=   256
         Columns(4).Locked=   -1  'True
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
         Columns(6).Width=   2540
         Columns(6).Caption=   "Rx"
         Columns(6).Name =   "FecRx"
         Columns(6).Alignment=   1
         Columns(6).CaptionAlignment=   1
         Columns(6).DataField=   "FecRx"
         Columns(6).DataType=   7
         Columns(6).NumberFormat=   "dd/mm/yyyy"
         Columns(6).FieldLen=   256
         Columns(7).Width=   979
         Columns(7).Caption=   "Lab"
         Columns(7).Name =   "ProEsLa"
         Columns(7).Alignment=   1
         Columns(7).CaptionAlignment=   2
         Columns(7).DataField=   "ProEsLa"
         Columns(7).DataType=   3
         Columns(7).FieldLen=   256
         Columns(7).Style=   2
         Columns(8).Width=   2037
         Columns(8).Caption=   "Laboratorio"
         Columns(8).Name =   "FecLab"
         Columns(8).Alignment=   1
         Columns(8).CaptionAlignment=   1
         Columns(8).DataField=   "FecLab"
         Columns(8).DataType=   7
         Columns(8).NumberFormat=   "dd/mm/yyyy"
         Columns(8).FieldLen=   256
         _ExtentX        =   25215
         _ExtentY        =   8070
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
      Begin MSAdodcLib.Adodc Adodc5 
         Height          =   330
         Left            =   -71955
         Top             =   3105
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
         RecordSource    =   "Select * from nuevos Where NueEst = 4"
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid5 
         Bindings        =   "Programacion.frx":47C9
         Height          =   4575
         Left            =   -73800
         TabIndex        =   17
         Top             =   540
         Width           =   9975
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
         Columns.Count   =   4
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
         Columns(2).Width=   2566
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
         _ExtentX        =   17595
         _ExtentY        =   8070
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
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   1545
      Left            =   15120
      TabIndex        =   10
      Top             =   1575
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   2725
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.Image Image16 
         Height          =   285
         Left            =   315
         Picture         =   "Programacion.frx":47DE
         Stretch         =   -1  'True
         Top             =   945
         Width           =   285
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " Sin   Exámenes"
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
         Height          =   555
         Left            =   495
         TabIndex        =   12
         Top             =   855
         Width           =   1680
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image4 
         Height          =   285
         Left            =   315
         Picture         =   "Programacion.frx":A6C1
         Stretch         =   -1  'True
         Top             =   225
         Width           =   285
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "    Depurar"
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
         TabIndex        =   11
         Top             =   270
         Width           =   1725
      End
      Begin VB.Image Image2 
         Height          =   600
         Left            =   180
         Picture         =   "Programacion.frx":E290
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   90
         Width           =   1905
      End
      Begin VB.Image Image17 
         Height          =   600
         Left            =   180
         Picture         =   "Programacion.frx":10845
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   810
         Width           =   1905
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   780
      Left            =   15120
      TabIndex        =   8
      Top             =   3870
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   1376
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.Image Image13 
         Height          =   285
         Left            =   225
         Picture         =   "Programacion.frx":12DFA
         Stretch         =   -1  'True
         Top             =   225
         Width           =   285
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "     Retornar"
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
         TabIndex        =   9
         Top             =   270
         Width           =   1770
      End
      Begin VB.Image Image15 
         Height          =   510
         Left            =   90
         Picture         =   "Programacion.frx":175C1
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   135
         Width           =   2085
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   780
      Left            =   15120
      TabIndex        =   18
      Top             =   3150
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   1376
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "     Reprogramar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   360
         TabIndex        =   19
         Top             =   225
         Width           =   1770
      End
      Begin VB.Image Image1 
         Height          =   285
         Left            =   270
         Picture         =   "Programacion.frx":19B76
         Stretch         =   -1  'True
         Top             =   225
         Width           =   285
      End
      Begin VB.Image Image3 
         Height          =   555
         Left            =   90
         Picture         =   "Programacion.frx":1BA5D
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   90
         Width           =   2085
      End
   End
   Begin Threed.SSFrame SSFrame7 
      Height          =   735
      Left            =   390
      TabIndex        =   20
      Top             =   6270
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   1296
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   90
         TabIndex        =   21
         Top             =   180
         Width           =   7080
      End
      Begin VB.Label Label24 
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
         TabIndex        =   22
         Top             =   225
         Width           =   1785
      End
      Begin VB.Image Image6 
         Height          =   285
         Left            =   7470
         Picture         =   "Programacion.frx":1E012
         Stretch         =   -1  'True
         Top             =   225
         Width           =   285
      End
      Begin VB.Image Image7 
         Height          =   465
         Left            =   7335
         Picture         =   "Programacion.frx":2D87B
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   135
         Width           =   1905
      End
   End
   Begin VB.Image Image22 
      Height          =   600
      Left            =   45
      Picture         =   "Programacion.frx":302B3
      Stretch         =   -1  'True
      Top             =   45
      Width           =   600
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTRO DE EXAMENES"
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
      Left            =   14445
      TabIndex        =   7
      Top             =   225
      Width           =   2805
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
      Left            =   15480
      TabIndex        =   6
      Top             =   5640
      Width           =   1740
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   15405
      Picture         =   "Programacion.frx":3AD10
      Stretch         =   -1  'True
      Top             =   5610
      Width           =   345
   End
   Begin VB.Image Image14 
      Height          =   690
      Left            =   0
      Picture         =   "Programacion.frx":3D487
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17835
   End
   Begin VB.Image Image8 
      Height          =   510
      Left            =   15270
      Picture         =   "Programacion.frx":5A5F1
      Stretch         =   -1  'True
      ToolTipText     =   "Nuevo"
      Top             =   5520
      Width           =   1965
   End
End
Attribute VB_Name = "Programacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vnueid1, vnueid2, vnueid3, vnueid4 As Integer
Private Sub Form_Load()
KeyPreview = True
SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " personas registradas"
Adodc1.Refresh
SSIndexTab1.Tab = 0
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
Label3_Click
End Sub

Private Sub Image13_Click()
Label5_Click
End Sub

Private Sub Image4_Click()
Label8_Click
End Sub

Private Sub Image5_Click()
Label10_Click
End Sub

Private Sub Image6_Click()
Label1_Click
End Sub

Private Sub Image9_Click()
Label2_Click
End Sub

Private Sub Label1_Click()
If seleccion = 1 Then
    modoR = "R"
    Load NuevosIng
    NuevosIng.Text1.Text = ""
    NuevosIng.SSFrame5.Visible = True
    NuevosIng.SSFrame6.Visible = True
    'NuevosIng.SSOption5.Value = -1
    
    NuevosIng.SSFrame5.Enabled = False
    NuevosIng.SSFrame6.Enabled = False
    If modop = "S" Then
        vnueids = Adodc5.Recordset.Fields("NueId")
        If Programacion.Adodc5.Recordset.Fields("NueTip") = "PREOCUPACIONAL" Then
            NuevosIng.SSOption3.Value = -1
        ElseIf Programacion.Adodc5.Recordset.Fields("NueTip") = "POSTOCUPACIONAL" Then
            NuevosIng.SSOption4.Value = -1
        ElseIf Programacion.Adodc5.Recordset.Fields("NueTip") = "REINGRESO" Then
            NuevosIng.SSOption5.Value = -1
        End If
        
        
        
        NuevosIng.Text1.Text = Programacion.Adodc5.Recordset.Fields("NueApa")
        NuevosIng.Text9.Text = Programacion.Adodc5.Recordset.Fields("NueAma")
        NuevosIng.Text10.Text = Programacion.Adodc5.Recordset.Fields("NueNom")
        NuevosIng.Combo2.Text = Programacion.Adodc5.Recordset.Fields("NueSex")
        NuevosIng.SSDateCombo2.Text = Programacion.Adodc5.Recordset.Fields("NueFeN")
        NuevosIng.Text5.Text = Programacion.Adodc5.Recordset.Fields("NueEda")
        NuevosIng.Text3.Text = Programacion.Adodc5.Recordset.Fields("NueTeI")
        NuevosIng.Text6.Text = Programacion.Adodc5.Recordset.Fields("NueTeR")
        NuevosIng.Combo1.Text = Programacion.Adodc5.Recordset.Fields("EmpDes")
        NuevosIng.SSDateCombo3.Text = Programacion.Adodc5.Recordset.Fields("NueFeI")
        'Label20.Caption = Adodc1.Recordset.Fields("NueFeS")
        'SSDateCombo1.Text = Adodc1.Recordset.Fields("NueFeP")
        'SSDateCombo4.Text = Adodc1.Recordset.Fields("NueFeM")
        If NuevosIng.Combo2.Text = "FEMENINO" Then
            NuevosIng.SSCheck1.Visible = True
            If NuevosIng.Adodc1.Recordset.Fields("NueEmb") = -1 Then
                NuevosIng.SSCheck1.Value = -1
            Else
                NuevosIng.SSCheck1.Value = 0
            End If
        Else
            NuevosIng.SSCheck1.Visible = False
        End If
        vnueid = Programacion.Adodc5.Recordset.Fields("NueId")
    ElseIf modop = "D" Then
        vnueidr = Adodc4.Recordset.Fields("NueId")
        If Programacion.Adodc4.Recordset.Fields("NueTip") = "PREOCUPACIONAL" Then
            NuevosIng.SSOption3.Value = -1
        ElseIf Programacion.Adodc4.Recordset.Fields("NueTip") = "POSTOCUPACIONAL" Then
            NuevosIng.SSOption4.Value = -1
        ElseIf Programacion.Adodc4.Recordset.Fields("NueTip") = "REINGRESO" Then
            NuevosIng.SSOption5.Value = -1
        End If
        
        NuevosIng.Text1.Text = Programacion.Adodc4.Recordset.Fields("NueApa")
        NuevosIng.Text9.Text = Programacion.Adodc4.Recordset.Fields("NueAma")
        NuevosIng.Text10.Text = Programacion.Adodc4.Recordset.Fields("NueNombres")
        NuevosIng.Combo2.Text = Programacion.Adodc4.Recordset.Fields("NueSex")
        NuevosIng.SSDateCombo2.Text = Programacion.Adodc4.Recordset.Fields("NueFeN")
        NuevosIng.Text5.Text = Programacion.Adodc4.Recordset.Fields("NueEda")
        NuevosIng.Text3.Text = Programacion.Adodc4.Recordset.Fields("NueTeI")
        NuevosIng.Text6.Text = Programacion.Adodc4.Recordset.Fields("NueTeR")
        NuevosIng.Combo1.Text = Programacion.Adodc4.Recordset.Fields("EmpDes")
        NuevosIng.SSDateCombo3.Text = Programacion.Adodc4.Recordset.Fields("NueFeI")
        'Label20.Caption = Adodc1.Recordset.Fields("NueFeS")
        'SSDateCombo1.Text = Adodc1.Recordset.Fields("NueFeP")
        'SSDateCombo4.Text = Adodc1.Recordset.Fields("NueFeM")
        If NuevosIng.Combo2.Text = "FEMENINO" Then
            NuevosIng.SSCheck1.Visible = True
            If NuevosIng.Adodc1.Recordset.Fields("NueEmb") = -1 Then
                NuevosIng.SSCheck1.Value = -1
            Else
                NuevosIng.SSCheck1.Value = 0
            End If
        Else
            NuevosIng.SSCheck1.Visible = False
        End If
        vnueid = Programacion.Adodc4.Recordset.Fields("NueId")
    End If
    
    NuevosIng.Label25.Visible = True
    NuevosIng.Show
    modo = "N"
    NuevosIng.SSFrame1.Enabled = True
    NuevosIng.SSFrame5.Visible = True
    NuevosIng.SSFrame6.Visible = True
    NuevosIng.SSFrame5.Enabled = False
    NuevosIng.SSFrame6.Enabled = True
    NuevosIng.SSFrame3.Enabled = False
    NuevosIng.Label20.Caption = Date
Else
    MsgBox "Debe seleccionar una persona de la lista", vbInformation, empresa
End If
End Sub

Private Sub Label10_Click()
'If Adodc1.Recordset.RecordCount = 1 Then
'    Adodc1.Recordset.AddNew
'    Adodc1.Recordset.Fields("NueTIP") = "PREOCUPACIONAL"
'    Adodc1.Recordset.Fields("NueEst") = 1
'    Adodc1.Recordset.Update
'    Adodc1.Refresh
'    SSOleDBGrid1.MoveFirst
'    SSOleDBGrid1.MoveNext
'Else
SSOleDBGrid1.MoveFirst
SSOleDBGrid1.MoveNext
'End If
Unload Programacion
Set Programacion = Nothing
End Sub

Private Sub Label2_Click()
Adodc1.RecordSource = "SELECT * from nuevos WHERE NueEst = " & 3
Adodc1.Refresh
SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " personas depuradas"
SSFrame1.Visible = True
End Sub

Private Sub Label3_Click()
Adodc1.RecordSource = "SELECT * from nuevos WHERE NueEst = " & 2
SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " personas cerradas"
Adodc1.Refresh
SSFrame1.Visible = False
End Sub

Private Sub Label24_Click()
Dim db1 As String
If SSIndexTab1.Tab = 0 Then
    If Len(Trim(Text8.Text)) > 0 Then
        
        db1 = Text8.Text
        Adodc1.RecordSource = "SELECT * from nuevos WHERE  NueNom LIKE " & "'%" & db1 & "%' AND NueEst = 1 AND NueTip= 'PREOCUPACIONAL'  ORDER BY NueFem DESC"
        Adodc1.Refresh
        Text8.Text = ""
        SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " registros encontrados"
    Else
        Adodc1.RecordSource = "Select * from nuevos Where NueEst = 1 AND NueTip= 'PREOCUPACIONAL'  ORDER BY NueFeP DESC"
        Adodc1.Refresh
    End If

ElseIf SSIndexTab1.Tab = 1 Then
    If Len(Trim(Text8.Text)) > 0 Then
        
        db1 = Text8.Text
        Adodc2.RecordSource = "SELECT * from nuevos WHERE  NueNom LIKE " & "'%" & db1 & "%' AND NueEst = 1 AND NueTip= 'POSTOCUPACIONAL'  ORDER BY NueFem DESC"
        Adodc2.Refresh
        Text8.Text = ""
        SSOleDBGrid1.Caption = Adodc2.Recordset.RecordCount & " registros encontrados"
    Else
        Adodc2.RecordSource = "Select * from nuevos Where NueEst = 1 AND NueTip= 'POSTOCUPACIONAL'  ORDER BY NueFeP DESC"
        Adodc2.Refresh
    End If
ElseIf SSIndexTab1.Tab = 2 Then
    If Len(Trim(Text8.Text)) > 0 Then
        
        db1 = Text8.Text
        Adodc3.RecordSource = "SELECT * from nuevos WHERE  NueNom LIKE " & "'%" & db1 & "%' AND NueEst = 1 AND NueTip= 'REINGRESO'  ORDER BY NueFem DESC"
        Adodc3.Refresh
        Text8.Text = ""
        SSOleDBGrid1.Caption = Adodc3.Recordset.RecordCount & " registros encontrados"
    Else
        Adodc3.RecordSource = "Select * from nuevos Where NueEst = 1 AND NueTip= 'REINGRESO'  ORDER BY NueFeP DESC"
        Adodc3.Refresh
    End If
ElseIf SSIndexTab1.Tab = 3 Then
    If Len(Trim(Text8.Text)) > 0 Then
        
        db1 = Text8.Text
        Adodc4.RecordSource = "SELECT * from nuevos WHERE  NueNom LIKE " & "'%" & db1 & "%' AND NueEst = 3 ORDER BY NueFeP DESC"
        Adodc4.Refresh
        Text8.Text = ""
        SSOleDBGrid1.Caption = Adodc4.Recordset.RecordCount & " registros encontrados"
    Else
        Adodc4.RecordSource = "Select * from nuevos Where NueEst = 3 ORDER BY NueFem DESC"
        Adodc4.Refresh
    End If
ElseIf SSIndexTab1.Tab = 4 Then
    If Len(Trim(Text8.Text)) > 0 Then
        
        db1 = Text8.Text
        Adodc5.RecordSource = "SELECT * from nuevos WHERE  NueNom LIKE " & "'%" & db1 & "%' AND NueEst = 4 ORDER BY NueFeP DESC"
        Adodc5.Refresh
        Text8.Text = ""
        SSOleDBGrid1.Caption = Adodc5.Recordset.RecordCount & " registros encontrados"
    Else
        Adodc5.RecordSource = "Select * from nuevos Where NueEst = 4 ORDER BY NueFem DESC"
        Adodc5.Refresh
    End If
End If
End Sub

Private Sub Label5_Click()
If seleccion = 1 Then
    If MsgBox("Esta seguro de retornar a esta persona ?", vbYesNo, empresa) = vbYes Then
        Dim Cn As New ADODB.Connection
        Cn.ConnectionString = Cadena
        Cn.Open
        vnueid = Adodc4.Recordset.Fields("NueId")
        depura = "UPDATE nuevos SET NueEst= " & 1 & " WHERE NueId = " & vnueid
        Cn.Execute depura
        SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " personas registradas"
        Adodc1.Refresh
    End If
Else
    MsgBox "Debe seleccionar una persona de la lista", vbInformation, empresa
End If
End Sub

Private Sub Label6_Click()
If seleccion = 1 Then
    If MsgBox("Esta seguro de llevar a la lista de 'Sin exámen' a esta persona ?", vbYesNo, empresa) = vbYes Then
        Dim Cn As New ADODB.Connection
        Cn.ConnectionString = Cadena
        Cn.Open
        
        depura = "UPDATE nuevos SET NueEst= " & 4 & " WHERE NueId = " & vnueid
        Cn.Execute depura
        SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " personas registradas"
        Adodc1.Refresh
    End If
Else
    MsgBox "Debe seleccionar una persona de la lista", vbInformation, empresa
End If
End Sub

Private Sub Label8_Click()
If seleccion = 1 Then
    If MsgBox("Esta seguro de depurar a esta persona ?", vbYesNo, empresa) = vbYes Then
        If SSIndexTab1.Tab = 0 Then
            vnueid = Adodc1.Recordset.Fields("NueId")
        End If
        If SSIndexTab1.Tab = 1 Then
            vnueid = Adodc2.Recordset.Fields("NueId")
        End If
        If SSIndexTab1.Tab = 2 Then
            vnueid = Adodc3.Recordset.Fields("NueId")
        End If
        If SSIndexTab1.Tab = 3 Then
            vnueid = Adodc4.Recordset.Fields("NueId")
        End If
        If SSIndexTab1.Tab = 3 Then
            vnueid = Adodc5.Recordset.Fields("NueId")
        End If
        Dim Cn As New ADODB.Connection
        Cn.ConnectionString = Cadena
        Cn.Open
        
        depura = "UPDATE nuevos SET NueEst= " & 3 & " WHERE NueId = " & vnueid
        Cn.Execute depura
        
        Adodc1.Refresh
        Adodc2.Refresh
        Adodc3.Refresh
        Adodc4.Refresh
        Adodc5.Refresh
        
        SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " personas registradas"
        Adodc1.Refresh
    End If
Else
    MsgBox "Debe seleccionar una persona de la lista", vbInformation, empresa
End If
End Sub

Private Sub SSCheck1_Click(Value As Integer)
If SSCheck1.Value = -1 Then
    SSDateCombo3.Enabled = True
    SSDateCombo3.SetFocus
Else
    SSDateCombo3.Text = ""
    SSDateCombo3.Enabled = False
End If
End Sub

Private Sub SSCheck2_Click(Value As Integer)
If SSCheck2.Value = -1 Then
    SSDateCombo1.Enabled = True
    SSDateCombo1.SetFocus
Else
    SSDateCombo3.Text = ""
    SSDateCombo1.Enabled = False
End If
End Sub
Private Sub SSCommand1_Click()
On Error GoTo errorex

If seleccion = 1 Then
    If SSCheck1.Value = -1 Or SSCheck2.Value = -1 Then
        Dim Cn As New ADODB.Connection
        Cn.ConnectionString = Cadena
        Cn.Open
        
        If SSIndexTab1.Tab = 0 Then
            vnueid = Adodc1.Recordset.Fields("NueId")
            If SSCheck1.Value = -1 Then
                vfecrx = Format(SSDateCombo3.Text, "yyyy-mm-dd")
                modifica = "UPDATE nuevos SET FecRx = " & "'" & vfecrx & "', ProEsRx = " & -1 & ", NueEst = " & 5 & ", UsuRes2 = " & "'" & vUsuario & "' WHERE NueId = " & vnueid
                Cn.Execute modifica
            End If
            If SSCheck2.Value = -1 Then
                vfecla = Format(SSDateCombo1.Text, "yyyy-mm-dd")
                modifica = "UPDATE nuevos SET FecLab = " & "'" & vfecla & "', ProEsla = " & -1 & ", NueEst = " & 5 & ", UsuRes2 = " & "'" & vUsuario & "' WHERE NueId = " & vnueid
                Cn.Execute modifica
            End If
            Adodc1.Refresh
            seleccion = 0
            limpiadatos
        End If
        
        If SSIndexTab1.Tab = 1 Then
            vnueid = Adodc2.Recordset.Fields("NueId")
            If SSCheck1.Value = -1 Then
                vfecrx = Format(SSDateCombo3.Text, "yyyy-mm-dd")
                modifica = "UPDATE nuevos SET FecRx = " & "'" & vfecrx & "', ProEsRx = " & -1 & ", UsuRes2 = " & "'" & vUsuario & "' WHERE NueId = " & vnueid
                Cn.Execute modifica
            End If
            If SSCheck2.Value = -1 Then
                vfecla = Format(SSDateCombo1.Text, "yyyy-mm-dd")
                modifica = "UPDATE nuevos SET FecLab = " & "'" & vfecla & "', ProEsla = " & -1 & ", UsuRes2 = " & "'" & vUsuario & "' WHERE NueId = " & vnueid
                Cn.Execute modifica
            End If
            Adodc2.Refresh
            limpiadatos
        End If
        
        If SSIndexTab1.Tab = 2 Then
            vnueid = Adodc3.Recordset.Fields("NueId")
            If SSCheck1.Value = -1 Then
                vfecrx = Format(SSDateCombo3.Text, "yyyy-mm-dd")
                modifica = "UPDATE nuevos SET FecRx = " & "'" & vfecrx & "', ProEsRx = " & -1 & ", UsuRes2 = " & "'" & vUsuario & "' WHERE NueId = " & vnueid
                Cn.Execute modifica
            End If
            If SSCheck2.Value = -1 Then
                vfecla = Format(SSDateCombo1.Text, "yyyy-mm-dd")
                modifica = "UPDATE nuevos SET FecLab = " & "'" & vfecla & "', ProEsla = " & -1 & ", UsuRes2 = " & "'" & vUsuario & "' WHERE NueId = " & vnueid
                Cn.Execute modifica
            End If
            Adodc3.Refresh
            limpiadatos
        End If
        
        If SSIndexTab1.Tab = 3 Then
            vnueid = Adodc4.Recordset.Fields("NueId")
            If SSCheck1.Value = -1 Then
                vfecrx = Format(SSDateCombo3.Text, "yyyy-mm-dd")
                modifica = "UPDATE nuevos SET FecRx = " & "'" & vfecrx & "', ProEsRx = " & -1 & ", UsuRes2 = " & "'" & vUsuario & "' WHERE NueId = " & vnueid
                Cn.Execute modifica
            End If
            If SSCheck2.Value = -1 Then
                vfecla = Format(SSDateCombo1.Text, "yyyy-mm-dd")
                modifica = "UPDATE nuevos SET FecLab = " & "'" & vfecla & "', ProEsla = " & -1 & ", UsuRes2 = " & "'" & vUsuario & "' WHERE NueId = " & vnueid
                Cn.Execute modifica
            End If
            Adodc4.Refresh
            seleccion = 0
            limpiadatos
        End If
        
        Cn.Close
    Else
        MsgBox "No existen fechas registradas", vbInformation, empresa
    End If
Else
    MsgBox "Seleccione una persona de la lista", vbInformation, empresa
End If

errorex:
If Err.Number = -2147467259 Then
    MsgBox "Existe error en las fechas", vbInformation, empresa
End If
End Sub

Private Sub SSIndexTab1_Click(PreviousTab As Integer)
limpiadatos
If SSIndexTab1.Tab = 0 Then
    Adodc1.RecordSource = "Select * from nuevos Where NueEst = 1 AND NueTip='PREOCUPACIONAL' ORDER BY NueFeP DESC"
    SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " Personas registradas"
    Adodc1.Refresh
End If
If SSIndexTab1.Tab = 1 Then
    Adodc2.RecordSource = "Select * from nuevos Where NueEst = 1 AND NueTip = 'POSTOCUPACIONAL' ORDER BY NueFeP DESC"
    SSOleDBGrid2.Caption = Adodc2.Recordset.RecordCount & " Personas registradas"
    Adodc2.Refresh
ElseIf SSIndexTab1.Tab = 2 Then
    Adodc3.RecordSource = "Select * from nuevos Where NueEst = 1 AND NueTip = 'REPROGRAMACION' ORDER BY NueFeP DESC"
    SSOleDBGrid3.Caption = Adodc3.Recordset.RecordCount & " Personas registradas"
    Adodc3.Refresh
ElseIf SSIndexTab1.Tab = 3 Then
    SSFrame1.Visible = True
    SSFrame2.Visible = True
    Adodc4.RecordSource = "Select * from nuevos Where NueEst = " & 3
    SSOleDBGrid4.Caption = Adodc4.Recordset.RecordCount & " Personas registradas"
    Adodc4.Refresh
    modop = "D"
    SSFrame1.Visible = True
ElseIf SSIndexTab1.Tab = 4 Then
    Adodc5.RecordSource = "Select * from nuevos Where NueEst = " & 4
    SSOleDBGrid5.Caption = Adodc5.Recordset.RecordCount & " Personas registradas"
    Adodc5.Refresh
    SSFrame2.Visible = True
    SSFrame1.Visible = False
    modop = "S"
End If

End Sub

Private Sub SSOleDBGrid1_Click()
If Adodc1.Recordset.RecordCount > 0 Then
    seleccion = 1
    limpiadatos
    vnueid1 = Adodc1.Recordset.Fields("NueId")
    Label20.Caption = Adodc1.Recordset.Fields("NueNom")
    Label2.Caption = Adodc1.Recordset.Fields("EmpDes")
    'SSDateCombo3.Text = Adodc1.Recordset.Fields("FecRx") & ""
    'SSDateCombo1.Text = Adodc1.Recordset.Fields("FecLab") & ""
    'If Adodc1.Recordset.Fields("ProEsRx") = -1 Then
    '    SSCheck1.Value = -1
    'Else
    '    SSCheck1.Value = 0
    'End If
    'If Adodc1.Recordset.Fields("ProEsLa") = -1 Then
    '    SSCheck2.Value = -1
    'Else
    '    SSCheck2.Value = 0
    'End If
    SSCheck1.SetFocus
Else
    seleccion = 0
    MsgBox "No existen datos en la lista", vbInformation, empresa
End If
End Sub


Private Sub SSOleDBGrid2_Click()
If Adodc2.Recordset.RecordCount > 0 Then
    seleccion = 1
    limpiadatos
    vnueid2 = Adodc2.Recordset.Fields("NueId")
    Label20.Caption = Adodc2.Recordset.Fields("NueNom")
    Label2.Caption = Adodc2.Recordset.Fields("EmpDes")
    SSCheck1.SetFocus
Else
    seleccion = 0
    MsgBox "No existen datos en la lista", vbInformation, empresa
End If
End Sub

Private Sub SSOleDBGrid3_Click()
If Adodc3.Recordset.RecordCount > 0 Then
    seleccion = 1
    limpiadatos
    vnueid3 = Adodc3.Recordset.Fields("NueId")
    Label20.Caption = Adodc3.Recordset.Fields("NueNom")
    Label2.Caption = Adodc3.Recordset.Fields("EmpDes")
    'SSDateCombo3.SetFocus
Else
    seleccion = 0
    MsgBox "No existen datos en la lista", vbInformation, empresa
End If
End Sub
Private Sub SSOleDBGrid4_Click()
If Adodc4.Recordset.RecordCount > 0 Then
    seleccion = 1
    limpiadatos
    vnueid4 = Adodc4.Recordset.Fields("NueId")
    Label20.Caption = Adodc4.Recordset.Fields("NueNom")
    Label2.Caption = Adodc4.Recordset.Fields("EmpDes")
    SSCheck1.SetFocus
Else
    seleccion = 0
    MsgBox "No existen datos en la lista", vbInformation, empresa
End If
End Sub

Private Sub SSOleDBGrid5_Click()
If Adodc5.Recordset.RecordCount > 0 Then
    seleccion = 1
Else
    seleccion = 0
    MsgBox "No existen datos en la lista", vbInformation, empresa
End If

End Sub
Private Sub Text8_GotFocus()
Text8.BackColor = &HC0FFFF
End Sub
Private Sub Text8_LostFocus()
Text8.BackColor = &HFFFFFF
Text8.Text = UCase(Text8.Text)
End Sub

Private Function limpiadatos()
Label20.Caption = ""
Label2.Caption = ""
'SSDateCombo1.Text = Date
'SSDateCombo3.Text = Date
'SSCheck1.Value = 0
'SSCheck2.Value = 0
End Function
