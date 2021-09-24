VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form NuevosIng 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10980
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   19755
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   19755
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   17685
      Top             =   6435
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowProgressCtls=   0   'False
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Threed.SSFrame SSFrame7 
      Height          =   1545
      Left            =   90
      TabIndex        =   54
      Top             =   4725
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2725
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.TextBox Text15 
         Height          =   375
         Left            =   90
         TabIndex        =   78
         Top             =   990
         Width           =   4515
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   90
         TabIndex        =   0
         Top             =   270
         Width           =   4515
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar"
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
         Left            =   4950
         TabIndex        =   80
         Top             =   1035
         Width           =   1455
      End
      Begin VB.Image Image20 
         Height          =   285
         Left            =   4815
         Picture         =   "NuevosIng.frx":0000
         Stretch         =   -1  'True
         Top             =   1035
         Width           =   285
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Por Empresa"
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
         Height          =   285
         Left            =   90
         TabIndex        =   79
         Top             =   675
         Width           =   1500
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Por Nombre"
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
         Left            =   90
         TabIndex        =   65
         Top             =   0
         Width           =   1500
      End
      Begin VB.Image Image16 
         Height          =   285
         Left            =   4815
         Picture         =   "NuevosIng.frx":F869
         Stretch         =   -1  'True
         Top             =   315
         Width           =   285
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar"
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
         Left            =   4860
         TabIndex        =   55
         Top             =   315
         Width           =   1545
      End
      Begin VB.Image Image17 
         Height          =   465
         Left            =   4725
         Picture         =   "NuevosIng.frx":1F0D2
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   225
         Width           =   1545
      End
      Begin VB.Image Image21 
         Height          =   465
         Left            =   4725
         Picture         =   "NuevosIng.frx":21B0A
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   945
         Width           =   1545
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   1005
      Left            =   17370
      TabIndex        =   39
      Top             =   3825
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   1773
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   780
         TabIndex        =   40
         Top             =   315
         Width           =   1185
      End
      Begin VB.Image Image5 
         Height          =   285
         Left            =   360
         Picture         =   "NuevosIng.frx":24542
         Stretch         =   -1  'True
         Top             =   315
         Width           =   285
      End
      Begin VB.Image Image8 
         Height          =   600
         Left            =   180
         Picture         =   "NuevosIng.frx":26CB9
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   180
         Width           =   1995
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   2940
      Left            =   17370
      TabIndex        =   36
      Top             =   765
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   5186
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   765
         TabIndex        =   38
         Top             =   690
         Width           =   1275
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Modificar Datos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   780
         TabIndex        =   37
         Top             =   1710
         Width           =   1230
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image3 
         Height          =   285
         Left            =   270
         Picture         =   "NuevosIng.frx":293FA
         Stretch         =   -1  'True
         Top             =   720
         Width           =   285
      End
      Begin VB.Image Image4 
         Height          =   285
         Left            =   225
         Picture         =   "NuevosIng.frx":2C246
         Stretch         =   -1  'True
         Top             =   1845
         Width           =   285
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   90
         Picture         =   "NuevosIng.frx":2E273
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   540
         Width           =   2175
      End
      Begin VB.Image Image2 
         Height          =   660
         Left            =   60
         Picture         =   "NuevosIng.frx":30828
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   1665
         Width           =   2175
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
      RecordSource    =   "Select * from nuevos  ORDER BY NueFeS DESC"
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
   Begin Threed.SSFrame SSFrame1 
      Height          =   4605
      Left            =   90
      TabIndex        =   23
      Top             =   6300
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   8123
      _Version        =   196608
      BackStyle       =   1
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.TextBox Text11 
         Height          =   735
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   64
         Top             =   3735
         Width           =   8970
      End
      Begin Threed.SSCheck SSCheck2 
         Height          =   465
         Left            =   13320
         TabIndex        =   60
         Top             =   2880
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   820
         _Version        =   196608
         ForeColor       =   192
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
         Caption         =   "U R G E N T E "
         Alignment       =   1
      End
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   345
         Left            =   9405
         Top             =   4050
         Visible         =   0   'False
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   609
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
         RecordSource    =   "Select * from nuevos order by NueId DESC"
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
      Begin Threed.SSCommand SSCommand1 
         Height          =   375
         Left            =   12420
         TabIndex        =   20
         Top             =   3870
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   661
         _Version        =   196608
         ForeColor       =   16777215
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
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
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   5655
         TabIndex        =   8
         Top             =   1515
         Width           =   2310
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   1515
         Width           =   2265
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   9405
         Top             =   3690
         Visible         =   0   'False
         Width           =   1995
         _ExtentX        =   3519
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
         RecordSource    =   "Select * from nuevos"
         Caption         =   "Adodc2"
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
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   10305
         TabIndex        =   16
         Top             =   2235
         Width           =   2040
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   12465
         TabIndex        =   11
         Top             =   1440
         Width           =   870
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "NuevosIng.frx":32DDD
         Left            =   8325
         List            =   "NuevosIng.frx":32DE7
         TabIndex        =   9
         Top             =   1515
         Width           =   1905
      End
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo1 
         Height          =   420
         Left            =   2925
         TabIndex        =   17
         Top             =   3015
         Width           =   1815
         _Version        =   65537
         _ExtentX        =   3201
         _ExtentY        =   741
         _StockProps     =   93
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   2295
         Width           =   4335
      End
      Begin VB.TextBox Text3 
         Height          =   330
         Left            =   7560
         TabIndex        =   15
         Top             =   2235
         Width           =   2040
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   90
         TabIndex        =   6
         Top             =   1515
         Width           =   2355
      End
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo2 
         Height          =   375
         Left            =   10350
         TabIndex        =   10
         Top             =   1440
         Width           =   1770
         _Version        =   65537
         _ExtentX        =   3122
         _ExtentY        =   661
         _StockProps     =   93
      End
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo3 
         Height          =   330
         Left            =   5535
         TabIndex        =   14
         Top             =   2250
         Width           =   1860
         _Version        =   65537
         _ExtentX        =   3281
         _ExtentY        =   582
         _StockProps     =   93
      End
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo4 
         Height          =   420
         Left            =   5535
         TabIndex        =   18
         Top             =   3015
         Width           =   1860
         _Version        =   65537
         _ExtentX        =   3281
         _ExtentY        =   741
         _StockProps     =   93
      End
      Begin Threed.SSCheck SSCheck1 
         Height          =   330
         Left            =   13725
         TabIndex        =   12
         Top             =   1440
         Visible         =   0   'False
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   582
         _Version        =   196608
         ForeColor       =   64
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
         Caption         =   "Posible Embarazo"
      End
      Begin Threed.SSFrame SSFrame5 
         Height          =   1050
         Left            =   45
         TabIndex        =   42
         Top             =   90
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   1852
         _Version        =   196608
         BackStyle       =   1
         ClipControls    =   0   'False
         Begin VB.ComboBox Combo3 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "NuevosIng.frx":32E00
            Left            =   4680
            List            =   "NuevosIng.frx":32E0A
            TabIndex        =   77
            Top             =   540
            Width           =   2085
         End
         Begin Threed.SSOption SSOption3 
            Height          =   420
            Left            =   45
            TabIndex        =   1
            Top             =   0
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   741
            _Version        =   196608
            ForeColor       =   64
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
            Caption         =   "Preocupacional"
            Alignment       =   1
            Value           =   -1
         End
         Begin Threed.SSOption SSOption4 
            Height          =   420
            Left            =   2250
            TabIndex        =   43
            Top             =   0
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   741
            _Version        =   196608
            ForeColor       =   64
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
            Caption         =   "Postocupacional"
            Alignment       =   1
         End
         Begin Threed.SSOption SSOption5 
            Height          =   420
            Left            =   4680
            TabIndex        =   44
            Top             =   0
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   741
            _Version        =   196608
            ForeColor       =   64
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
            Caption         =   "Reprogramación"
            Alignment       =   1
         End
      End
      Begin Threed.SSFrame SSFrame6 
         Height          =   825
         Left            =   6975
         TabIndex        =   45
         Top             =   315
         Width           =   9915
         _ExtentX        =   17489
         _ExtentY        =   1455
         _Version        =   196608
         BackStyle       =   1
         ClipControls    =   0   'False
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1305
            TabIndex        =   3
            Top             =   315
            Width           =   1905
         End
         Begin VB.TextBox Text7 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4500
            TabIndex        =   4
            Top             =   315
            Width           =   2085
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   375
            Left            =   7740
            TabIndex        =   5
            Top             =   315
            Width           =   2085
         End
         Begin Threed.SSOption SSOption1 
            Height          =   285
            Left            =   3240
            TabIndex        =   46
            Top             =   405
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            _Version        =   196608
            ForeColor       =   64
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
            Caption         =   "SIGEP"
            Alignment       =   1
         End
         Begin Threed.SSOption SSOption2 
            Height          =   330
            Left            =   6570
            TabIndex        =   47
            Top             =   405
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
            _Version        =   196608
            ForeColor       =   64
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
            Caption         =   "Cobro"
            Alignment       =   1
         End
         Begin Threed.SSOption SSOption6 
            Height          =   285
            Left            =   -45
            TabIndex        =   2
            Top             =   405
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   503
            _Version        =   196608
            ForeColor       =   64
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
            Caption         =   "Recibo"
            Alignment       =   1
            Value           =   -1
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "No. de Recibo "
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
            Left            =   1395
            TabIndex        =   50
            Top             =   45
            Width           =   1815
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "No. de Doc."
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
            Left            =   4500
            TabIndex        =   49
            Top             =   45
            Width           =   1635
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
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
            Height          =   285
            Left            =   7785
            TabIndex        =   48
            Top             =   45
            Width           =   1455
         End
      End
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo5 
         Height          =   420
         Left            =   8190
         TabIndex        =   19
         Top             =   3015
         Width           =   1770
         _Version        =   65537
         _ExtentX        =   3122
         _ExtentY        =   741
         _StockProps     =   93
         NullDateLabel   =   "0000-00-00"
      End
      Begin Threed.SSCheck SSCheck3 
         Height          =   330
         Left            =   13275
         TabIndex        =   61
         Top             =   2205
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   582
         _Version        =   196608
         ForeColor       =   64
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
         Caption         =   "Laboratorio"
         Value           =   1
      End
      Begin Threed.SSCheck SSCheck4 
         Height          =   330
         Left            =   15390
         TabIndex        =   62
         Top             =   2205
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   582
         _Version        =   196608
         ForeColor       =   64
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
         Caption         =   "Rx"
         Value           =   1
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
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
         Left            =   225
         TabIndex        =   63
         Top             =   3465
         Width           =   1590
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Depuración"
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
         Height          =   285
         Left            =   8190
         TabIndex        =   59
         Top             =   2700
         Width           =   2355
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombres"
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
         Left            =   5715
         TabIndex        =   58
         Top             =   1170
         Width           =   1125
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "AP. Materno "
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
         Left            =   2925
         TabIndex        =   57
         Top             =   1185
         Width           =   1440
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "REPROGRAMACIÓN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   7020
         TabIndex        =   56
         Top             =   0
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Limpiar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   14640
         TabIndex        =   41
         Top             =   3960
         Width           =   1275
      End
      Begin VB.Image Image13 
         Height          =   285
         Left            =   14235
         Picture         =   "NuevosIng.frx":32E2F
         Stretch         =   -1  'True
         Top             =   3945
         Width           =   285
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Left            =   270
         TabIndex        =   35
         Top             =   3015
         Width           =   1815
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Revisión Médica"
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
         Left            =   5370
         TabIndex        =   34
         Top             =   2685
         Width           =   2760
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha programada"
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
         Left            =   2835
         TabIndex        =   33
         Top             =   2700
         Width           =   2355
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono de encargado"
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
         Height          =   285
         Left            =   10260
         TabIndex        =   32
         Top             =   1965
         Width           =   2850
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Ingreso"
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
         Left            =   5535
         TabIndex        =   31
         Top             =   1965
         Width           =   1905
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Edad"
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
         Left            =   12600
         TabIndex        =   30
         Top             =   1170
         Width           =   600
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Nacim."
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
         Left            =   10395
         TabIndex        =   29
         Top             =   1170
         Width           =   2040
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Sexo"
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
         Left            =   8370
         TabIndex        =   28
         Top             =   1170
         Width           =   600
      End
      Begin VB.Image Image11 
         Height          =   285
         Left            =   11970
         Picture         =   "NuevosIng.frx":3A6CB
         Stretch         =   -1  'True
         Top             =   3915
         Width           =   285
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Solicitada"
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
         Left            =   270
         TabIndex        =   27
         Top             =   2700
         Width           =   1860
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
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
         Height          =   375
         Left            =   270
         TabIndex        =   26
         Top             =   2010
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono Interesado(a)"
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
         Height          =   225
         Left            =   7515
         TabIndex        =   25
         Top             =   1980
         Width           =   2400
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "AP. Paterno "
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
         Height          =   240
         Left            =   135
         TabIndex        =   24
         Top             =   1170
         Width           =   1305
      End
      Begin VB.Image Image12 
         Height          =   600
         Left            =   11775
         Picture         =   "NuevosIng.frx":3C446
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   3765
         Width           =   2130
      End
      Begin VB.Image Image15 
         Height          =   555
         Left            =   14100
         Picture         =   "NuevosIng.frx":3EE40
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   3810
         Width           =   1905
      End
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid1 
      Bindings        =   "NuevosIng.frx":41360
      Height          =   3945
      Left            =   90
      TabIndex        =   22
      Top             =   720
      Width           =   17190
      _Version        =   196616
      BevelColorFrame =   14737632
      BevelColorHighlight=   14737632
      BevelColorFace  =   8421376
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
      BackColorEven   =   14737632
      BackColorOdd    =   14737632
      RowHeight       =   423
      ExtraHeight     =   370
      Columns.Count   =   10
      Columns(0).Width=   1905
      Columns(0).Caption=   "Fec. Solic."
      Columns(0).Name =   "NueFeS"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   1
      Columns(0).DataField=   "NueFeS"
      Columns(0).DataType=   7
      Columns(0).NumberFormat=   "dd/mm/yyyy"
      Columns(0).FieldLen=   256
      Columns(1).Width=   7064
      Columns(1).Caption=   "Nombre"
      Columns(1).Name =   "NueNom"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "NueNom"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1905
      Columns(2).Caption=   "Teléfonos"
      Columns(2).Name =   "NueTel"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "NueTeI"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   6271
      Columns(3).Caption=   "Empresas"
      Columns(3).Name =   "empdes"
      Columns(3).CaptionAlignment=   0
      Columns(3).DataField=   "EmpDes"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1852
      Columns(4).Caption=   "Fecha Exa."
      Columns(4).Name =   "NueFeE"
      Columns(4).Alignment=   1
      Columns(4).CaptionAlignment=   1
      Columns(4).DataField=   "NueFeP"
      Columns(4).DataType=   7
      Columns(4).NumberFormat=   "dd/mm/yyyy"
      Columns(4).FieldLen=   256
      Columns(5).Width=   1826
      Columns(5).Caption=   "Rev. Med."
      Columns(5).Name =   "NueFeM"
      Columns(5).Alignment=   1
      Columns(5).CaptionAlignment=   1
      Columns(5).DataField=   "NueFeM"
      Columns(5).DataType=   7
      Columns(5).NumberFormat=   "dd/mm/yyyy"
      Columns(5).FieldLen=   256
      Columns(6).Width=   1799
      Columns(6).Caption=   "Tipo"
      Columns(6).Name =   "NueTip"
      Columns(6).CaptionAlignment=   0
      Columns(6).DataField=   "NueTip"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   2011
      Columns(7).Caption=   "Recibo"
      Columns(7).Name =   "NueRec"
      Columns(7).CaptionAlignment=   0
      Columns(7).DataField=   "NueRec"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   2064
      Columns(8).Caption=   "Sigep"
      Columns(8).Name =   "NueSigep"
      Columns(8).CaptionAlignment=   0
      Columns(8).DataField=   "NueSigep"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   2540
      Columns(9).Caption=   "Cobro"
      Columns(9).Name =   "NueCobro"
      Columns(9).CaptionAlignment=   0
      Columns(9).DataField=   "NueCobro"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      _ExtentX        =   30321
      _ExtentY        =   6959
      _StockProps     =   79
      Caption         =   "SSOleDBGrid1"
      ForeColor       =   16777215
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
      Height          =   3735
      Left            =   17280
      TabIndex        =   51
      Top             =   6750
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   6588
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.Image Image6 
         Height          =   285
         Left            =   375
         Picture         =   "NuevosIng.frx":41375
         Stretch         =   -1  'True
         Top             =   675
         Width           =   285
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Imprimir RX"
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
         Height          =   285
         Left            =   360
         TabIndex        =   53
         Top             =   720
         Width           =   1725
      End
      Begin VB.Image Image9 
         Height          =   285
         Left            =   360
         Picture         =   "NuevosIng.frx":443D7
         Stretch         =   -1  'True
         Top             =   2430
         Width           =   285
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Imprimir Lab."
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
         Left            =   315
         TabIndex        =   52
         Top             =   2475
         Width           =   1905
      End
      Begin VB.Image Image7 
         Height          =   555
         Left            =   210
         Picture         =   "NuevosIng.frx":47439
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   540
         Width           =   2130
      End
      Begin VB.Image Image10 
         Height          =   555
         Left            =   180
         Picture         =   "NuevosIng.frx":499EE
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   2295
         Width           =   2130
      End
   End
   Begin Threed.SSFrame SSFrame8 
      Height          =   1545
      Left            =   6480
      TabIndex        =   66
      Top             =   4725
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   2725
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.TextBox Text14 
         Enabled         =   0   'False
         Height          =   375
         Left            =   6930
         TabIndex        =   70
         Top             =   810
         Width           =   1770
      End
      Begin VB.TextBox Text13 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3870
         TabIndex        =   69
         Top             =   810
         Width           =   1860
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1260
         TabIndex        =   68
         Top             =   810
         Width           =   1320
      End
      Begin Threed.SSOption SSOption7 
         Height          =   285
         Left            =   2610
         TabIndex        =   71
         Top             =   900
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         _Version        =   196608
         ForeColor       =   64
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
         Caption         =   "SIGEP"
         Alignment       =   1
      End
      Begin Threed.SSOption SSOption8 
         Height          =   330
         Left            =   5670
         TabIndex        =   72
         Top             =   900
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         _Version        =   196608
         ForeColor       =   64
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
         Caption         =   "Cobro"
         Alignment       =   1
      End
      Begin Threed.SSOption SSOption9 
         Height          =   285
         Left            =   -90
         TabIndex        =   73
         Top             =   900
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   196608
         ForeColor       =   64
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
         Caption         =   "Recibo"
         Alignment       =   1
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
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
         Height          =   285
         Left            =   7065
         TabIndex        =   76
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "No. de Doc."
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
         Left            =   3915
         TabIndex        =   75
         Top             =   540
         Width           =   1635
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "No. de Recibo "
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
         Left            =   1215
         TabIndex        =   74
         Top             =   540
         Width           =   1815
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar"
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
         Left            =   9090
         TabIndex        =   67
         Top             =   855
         Width           =   1545
      End
      Begin VB.Image Image18 
         Height          =   285
         Left            =   8955
         Picture         =   "NuevosIng.frx":4BFA3
         Stretch         =   -1  'True
         Top             =   855
         Width           =   285
      End
      Begin VB.Image Image19 
         Height          =   465
         Left            =   8865
         Picture         =   "NuevosIng.frx":5B80C
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   765
         Width           =   1545
      End
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "PROGRAMACIÓN  DE EXÁMENES"
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
      Left            =   15390
      TabIndex        =   21
      Top             =   180
      Width           =   3615
   End
   Begin VB.Image Image22 
      Height          =   600
      Left            =   45
      Picture         =   "NuevosIng.frx":5E244
      Stretch         =   -1  'True
      Top             =   45
      Width           =   600
   End
   Begin VB.Image Image14 
      Height          =   690
      Left            =   0
      Picture         =   "NuevosIng.frx":68CA1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19770
   End
End
Attribute VB_Name = "NuevosIng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vEmpId As Integer
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Sub ChgEnterToTab(KeyCode As Integer)
If KeyCode = 13 Then
   KeyCode = 0
   SendKeys "{TAB}"
End If
End Sub
Private Sub Combo2_GotFocus()
If Len(Trim(Text10.Text)) = 0 Then
    MsgBox "Debe ingresar el nombre de la persona", vbInformation, EmpresaBox
    Text10.SetFocus
Else
    vnombre = Text1.Text & " " & Text9.Text & " " & Text10.Text
    Dim Cn As New ADODB.Connection
    Dim rsn As New ADODB.Recordset
    Cn.ConnectionString = Cadena
    Cn.Open
    
    rsn.CursorType = adOpenKeyset
    rsn.LockType = adLockOptimistic
    rsn.ActiveConnection = Cn
    rsn.Source = "Select * from nuevos Where NueNom = " & "'" & vnombre & "'"
    rsn.Open
    
    If Not rsn.EOF Then
        MsgBox "La persona ya tiene registro en la Base de Datos", vbInformation, EmpresaBox
        Combo2.Text = rsn!NueSex
        SSDateCombo2.Text = rsn!NueFeN
        Text5.Text = rsn!NueEda
        Combo1.Text = rsn!EmpDes
        vexiste = 1
    End If
    Cn.Close
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   SendKeys "{TAB}"
   KeyAscii = 0
End If
End Sub
Private Sub Combo1_GotFocus()
If vexiste = 0 Then
    
    Dim lResp As Long
    lResp = SendMessageLong(Combo1.hWnd, &H14F, True, 0)
    
    If Len(Trim(Combo1.Text)) > 0 Then
        vempresa = Combo1.Text
    End If
    
    Dim Cn As New ADODB.Connection
    Dim rsp As New ADODB.Recordset
    Cn.ConnectionString = Cadena
    Cn.Open
    
    rsp.CursorType = adOpenKeyset
    rsp.LockType = adLockOptimistic
    rsp.ActiveConnection = Cn
    rsp.Source = "Select * from empresas"
    rsp.Open
    
    If Not rsp.EOF Then
        rsp.MoveFirst
        Combo1.Clear
        Do While Not rsp.EOF
            Combo1.AddItem UCase(rsp!EmpDes)
            rsp.MoveNext
        Loop
    End If
    If Len(Trim(vempresa)) > 0 Then
        Combo1.Text = vempresa
    End If
    Combo1.SelStart = 0
    Combo1.SelLength = Len(Combo1.Text)
    Cn.Close
End If
End Sub
Private Sub Combo1_LostFocus()
Combo1.Text = UCase(Combo1.Text)

End Sub
Private Sub Form_Load()
KeyPreview = True
SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " personas registradas"
Adodc1.Refresh
End Sub
Private Sub Image13_Click()
Label21_Click
End Sub

Private Sub Image3_Click()
Label7_Click
End Sub
Private Sub Image4_Click()
Label8_Click
End Sub
Private Sub Image5_Click()
Label10_Click
End Sub

Private Sub Label1_Click()
Dim Cn As New ADODB.Connection
Dim rsdi As New ADODB.Recordset
Dim rsim As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open

borra = "DELETE FROM impbol"
Cn.Execute borra

reprx = "INSERT INTO impbol SELECT * FROM nuevos WHERE NueId = " & vnueid
Cn.Execute reprx

rsim.CursorType = adOpenKeyset
rsim.LockType = adLockOptimistic
rsim.ActiveConnection = Cn
rsim.Source = "Select * from impbol"
rsim.Open

rsdi.CursorType = adOpenKeyset
rsdi.LockType = adLockOptimistic
rsdi.ActiveConnection = Cn
rsdi.Source = "Select * from direcciones"
rsdi.Open

If Not rsim.EOF Then
    If rsim!NueEme = -1 Then
        direccionRx = rsdi!DirRxU
        direccionlab = rsdi!DirLabU
    ElseIf rsim!NueEme = 0 Then
        direccionRx = rsdi!DirRxN
        direccionlab = rsdi!DirLabN
        If Len(Trim(direccionRx)) = 0 Then
            reportenormal = 1
        End If
    End If
Else
    MsgBox "No existen direcciones de exámenes registradas", vbInformation, empresa
End If
Cn.Close

'Codigo QR
'codigoTXT = "1018223020|" & vNumFac & "|" & vAut & "|" & Text20.Text & "|" & vTotal & "|" & vTotal & "|" & Codcon & "|" & Text22.Text & "|0|0|0|0"
    
If reportenormal = 1 Then
    reportenormal = 0
    CrystalReport1.ReportFileName = App.Path & "\bolrxnuevo.rpt"
Else
    CrystalReport1.ReportFileName = App.Path & "\bolrxnuevo.rpt"
End If
'CrystalReport1.Formulas(0) = "direccion = " & "'" & direccionRx & "'"
CrystalReport1.Action = 1

End Sub
Private Sub Label10_Click()
Unload NuevosIng
Set NuevosIng = Nothing
End Sub
Private Sub Label2_Click()
Dim Cn As New ADODB.Connection
Dim rsdi As New ADODB.Recordset
Dim rsim As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open

borra = "DELETE FROM impbol"
Cn.Execute borra

reprx = "INSERT INTO impbol SELECT * FROM nuevos WHERE NueId = " & vnueid
Cn.Execute reprx

rsim.CursorType = adOpenKeyset
rsim.LockType = adLockOptimistic
rsim.ActiveConnection = Cn
rsim.Source = "Select * from impbol"
rsim.Open

rsdi.CursorType = adOpenKeyset
rsdi.LockType = adLockOptimistic
rsdi.ActiveConnection = Cn
rsdi.Source = "Select * from direcciones"
rsdi.Open

If Not rsim.EOF Then
    If rsim!NueEme = -1 Then
        direccionRx = rsdi!DirRxU
        direccionlab = rsdi!DirLabU
    ElseIf rsim!NueEme = 0 Then
        direccionRx = rsdi!DirRxN
        direccionlab = rsdi!DirLabN
        If Len(Trim(direccionRx)) = 0 Then
            reportenormal = 1
        End If
    End If
Else
    MsgBox "No existen direcciones de exámenes registradas", vbInformation, empresa
End If
Cn.Close

If reportenormal = 1 Then
    reportenormal = 0
    CrystalReport1.ReportFileName = App.Path & "\bollabnuevo.rpt"
Else
    CrystalReport1.ReportFileName = App.Path & "\bollabnuevo.rpt"
End If
'CrystalReport1.Formulas(0) = "direccion = " & "'" & direccionlab & "'"
CrystalReport1.Action = 1
End Sub
Private Sub Label21_Click()
limpiadatos
Text1.BackColor = &HFFFFFF
SSFrame1.Enabled = False
SSFrame3.Enabled = True
SSOleDBGrid1.SetFocus
End Sub
Private Sub Label24_Click()
If Len(Trim(Text8.Text)) > 0 Then
    Dim db1 As String
    db1 = Text8.Text
    Adodc1.RecordSource = "SELECT * from nuevos WHERE  NueNom LIKE " & "'%" & db1 & "%'"
    Adodc1.Refresh
    Text8.Text = ""
    SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " registros encontrados"
Else
    Adodc1.RecordSource = "Select * from nuevos  ORDER BY NueFeS DESC"
    Adodc1.Refresh
End If
End Sub

Private Sub Label31_Click()
Dim db1 As String
'Por Recibo
If SSOption9.Value = -1 Then
    If Len(Trim(Text12.Text)) > 0 Then
        
        db1 = Text12.Text
        Adodc1.RecordSource = "SELECT * from nuevos WHERE  NueRec LIKE " & "'%" & db1 & "%'"
        Adodc1.Refresh
        Text12.Text = ""
        SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " registros encontrados"
    Else
        Adodc1.RecordSource = "SELECT * from nuevos"
        Adodc1.Refresh
    End If
End If

'Por Sigep
If SSOption7.Value = -1 Then
    If Len(Trim(Text13.Text)) > 0 Then
        db1 = Text13.Text
        Adodc1.RecordSource = "SELECT * from nuevos WHERE  NueSigep LIKE " & "'%" & db1 & "%'"
        Adodc1.Refresh
        Text12.Text = ""
        SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " registros encontrados"
    Else
        Adodc1.RecordSource = "SELECT * from nuevos"
        Adodc1.Refresh
    End If
End If

'Por Cobro
If SSOption8.Value = -1 Then
    If Len(Trim(Text14.Text)) > 0 Then
        db1 = Text14.Text
        Adodc1.RecordSource = "SELECT * from nuevos WHERE  NueCOBRO LIKE " & "'%" & db1 & "%'"
        Adodc1.Refresh
        Text14.Text = ""
        SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " registros encontrados"
    Else
        Adodc1.RecordSource = "SELECT * from nuevos"
        Adodc1.Refresh
    End If
End If
End Sub

Private Sub Label35_Click()
If Len(Trim(Text15.Text)) > 0 Then
    Dim db1 As String
    db1 = Text15.Text
    Adodc1.RecordSource = "SELECT * from nuevos WHERE  EmpDes LIKE " & "'%" & db1 & "%' ORDER BY NueNom"
    Adodc1.Refresh
    Text15.Text = ""
    SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " registros encontrados"
Else
    Adodc1.RecordSource = "Select * from nuevos  ORDER BY NueFeS DESC"
    Adodc1.Refresh
End If

End Sub

Private Sub Label7_Click()
limpiadatos
modo = "N"
SSFrame1.Enabled = True
SSFrame5.Visible = True
SSFrame6.Visible = True
SSFrame5.Enabled = True
SSFrame6.Enabled = True
SSFrame3.Enabled = False
Label20.Caption = Date
SSOption3.SetFocus
End Sub

Private Sub Label8_Click()
If seleccion = 1 Then
    modo = "M"
    SSFrame1.Enabled = True
    SSFrame5.Visible = True
    SSFrame6.Visible = True
    SSFrame5.Enabled = True
    SSFrame6.Enabled = True
    SSFrame3.Enabled = False
    'Label20.Caption = Date
    Text1.SetFocus
Else
    MsgBox "Debe seleccionar el registro a modificar", vbInformation, EmpresaBox
End If
End Sub

Private Sub SSCheck1_Click(Value As Integer)
If SSCheck1.Value = -1 Then
    SSCheck4.Visible = False
Else
    SSCheck4.Visible = True
End If

End Sub

Private Sub SSCommand1_Click()
On Error GoTo errorreg

Dim Cn As New ADODB.Connection
Dim rsp As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open

If SSOption3.Value = -1 Then
    vNueTip = "PREOCUPACIONAL"
End If
If SSOption4.Value = -1 Then
    vNueTip = "POSTOCUPACIONAL"
End If
If SSOption5.Value = -1 Then
    vNueTip = "REPROGRAMACION"
End If


vnueapa = Trim(Text1.Text)
vnueama = Trim(Text9.Text)
vnuenombres = Trim(Text10.Text)
vnuenom = vnueapa & " " & vnueama & " " & vnuenombres
vNueSex = Combo2.Text
vNueFeN = Format(SSDateCombo2.Text, "yyyy-mm-dd")
vNueEda = Text5.Text
vNueTeI = Text3.Text
vNueTeR = Text6.Text
vEmpDes = Combo1.Text
vNueFeI = Format(SSDateCombo3.Text, "yyyy-mm-dd")
vNueFeS = Format(Label20.Caption, "yyyy-mm-dd")
vNueFeP = Format(SSDateCombo1.Text, "yyyy-mm-dd")
vNueFeM = Format(SSDateCombo4.Text, "yyyy-mm-dd")
vNueFeD = Format(SSDateCombo5.Text, "yyyy-mm-dd")
vNueRec = Text4.Text
vNueSigep = Text7.Text
vNueCobro = Text2.Text
vNueObs = Text11.Text
If vNueSex = "FEMENINO" Then
    If SSCheck1.Value = -1 Then
        vNueEmb = -1
    Else
        vNueEmb = 0
    End If
Else
    vNueEmb = 0
End If

If SSCheck2.Value = -1 Then
    vNueEme = -1
Else
    vNueEme = 0
End If
If Len(Trim(Combo3.Text)) > 0 Then
    vRepTip = Combo3.Text
Else
    vRepTip = ""
End If

If modo = "N" Then
    grabanuevo = "INSERT INTO Nuevos SET NueTip = " & "'" & vNueTip & "', NueApa = " & "'" & vnueapa & "', NueAma = " & "'" & vnueama & "', NueNombres = " & "'" & vnuenombres & _
    "', NueNom = " & "'" & vnuenom & "', NueSex = " & "'" & vNueSex & "', NueFeN = " & "'" & vNueFeN & "', NueEda = " & "'" & vNueEda & _
    "', NueTeI = " & "'" & vNueTeI & "', NueTeR = " & "'" & vNueTeR & "', EmpDes = " & "'" & vEmpDes & "', NueFeI = " & "'" & vNueFeI & "', NueFeS = " & "'" & vNueFeS & "', NueFeD = " & "'" & vNueFeD & _
    "', NueFeP = " & "'" & vNueFeP & "', NueFeM = " & "'" & vNueFeM & "', NueRec = " & "'" & vNueRec & "', Nuesigep = " & "'" & vNueSigep & "', Nuecobro = " & "'" & vNueCobro & _
    "', NueEmb = " & vNueEmb & ", NueEst = " & 1 & ", UsuRes1 = " & "'" & vUsuario & "', NueEme = " & vNueEme & ", NueObs = " & "'" & vNueObs & "', RepTip = " & "'" & vRepTip & "'"

    Cn.Execute grabanuevo
    SSOleDBGrid1.Caption = Adodc1.Recordset.RecordCount & " personas registradas"
    SSFrame2.Visible = True
    If SSCheck3.Value = -1 Then
        Label2.Enabled = True
    Else
        Label2.Enabled = False
    End If
    
    If SSCheck4.Value = -1 Then
        Label1.Enabled = True
    Else
        Label1.Enabled = False
    End If
    
    If SSCheck1.Value = -1 Then
        Label1.Enabled = False
    Else
        Label1.Enabled = True
    End If
    Adodc1.RecordSource = "Select * from nuevos ORDER BY NueFem DESC"
    Adodc1.Refresh
    Adodc3.Refresh
    vnueid = Adodc3.Recordset.Fields("Nueid")
    modo = ""
ElseIf modo = "M" Then
    modif = "UPDATE Nuevos SET NueTip = " & "'" & vNueTip & "', NueApa = " & "'" & vnueapa & "', NueAma = " & "'" & vnueama & "', NueNombres = " & "'" & vnuenombres & _
    "', NueNom = " & "'" & vnuenom & "', NueSex = " & "'" & vNueSex & "', NueFeN = " & "'" & vNueFeN & "', NueEda = " & "'" & vNueEda & _
    "', NueTeI = " & "'" & vNueTeI & "', NueTeR = " & "'" & vNueTeR & "', EmpDes = " & "'" & vEmpDes & "', NueFeI = " & "'" & vNueFeI & "', NueFeS = " & "'" & vNueFeS & "', NueFeS = " & "'" & vNueFeS & _
    "', NueFeD = " & "'" & vNueFeD & "', NueFeP = " & "'" & vNueFeP & "', NueFeM = " & "'" & vNueFeM & "', NueRec = " & "'" & vNueRec & "', Nuesigep = " & "'" & vNueSigep & "', Nuecobro = " & "'" & vNueCobro & _
    "', NueEmb = " & vNueEmb & ", UsuRes1 = " & "'" & vUsuario & "', NueEme = " & vNueEme & ", NueObs = " & "'" & vNueObs & "', RepTip = " & "'" & vRepTip & "' WHERE NueId = " & vnueid
    Cn.Execute modif
    
    'Adodc1.RecordSource = "Select * from nuevos ORDER BY NueFem DESC"
    Adodc1.Refresh
    Adodc3.Refresh
    MsgBox "Datos Modificados", vbInformation, empresa
    vnueid = Adodc1.Recordset.Fields("Nueid")
    modo = ""
End If
If modop = "S" Then
    actest = "UPDATE nuevos SET NueEst = " & 2 & " WHERE NueId = " & vnueids
    Cn.Execute actest
    modop = ""
    Programacion.Adodc5.Refresh
End If

SSFrame3.Enabled = True
SSFrame1.Enabled = False
seleccion = 0
modo = ""
If vexiste = 1 Then
    vexiste = 0
End If

errorreg:
If Err.Number = -2147467259 Then
    MsgBox "Existen datos faltantes.  Por favor revise"
End If
End Sub

Private Sub SSDateCombo1_LostFocus()
verificafecha
End Sub

Private Sub SSDateCombo2_GotFocus()
If Len(Trim(Combo2.Text)) = 0 Then
    MsgBox "Debe selecciona sexo", vbInformation, EmpresaBox
    Combo2.SetFocus
Else
    If Combo2.Text = "FEMENINO" Then
        SSCheck1.Visible = True
    Else
        SSCheck1.Visible = False
    End If
End If
End Sub

Private Sub SSDateCombo3_GotFocus()
If Len(Trim(Combo1.Text)) = 0 Then
    MsgBox "Debe ingresar el nombre de la empresa", vbInformation, empresa
    Combo1.SetFocus
Else
    vempresa = UCase(Combo1.Text)
    Dim Cn As New ADODB.Connection
    Dim rsp As New ADODB.Recordset
    Cn.ConnectionString = Cadena
    Cn.Open
    
    rsp.CursorType = adOpenKeyset
    rsp.LockType = adLockOptimistic
    rsp.ActiveConnection = Cn
    rsp.Source = "Select * from empresas WHERE EmpDes = " & "'" & vempresa & "'"
    rsp.Open
    
    If rsp.EOF Then
       If MsgBox("Empresa sin registro. Desea crearla?", vbYesNo, EmpresaBox) = vbYes Then
           If Combo1.Text <> "" Then
                creae = "INSERT INTO empresas SET EmpDes = " & "'" & vempresa & "'"
                Cn.Execute creae
           Else
                MsgBox "Denbe ingresar el nombre de la empresa", vbInformation, empresa
                Combo1.SetFocus
           End If
       End If
    Else
        vEmpId = rsp!EmpId
    End If
End If
End Sub

Private Sub SSOleDBGrid1_Click()
vid = Adodc1.Recordset.Fields("NueId")
Dim Cn As New ADODB.Connection
Dim rsp As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open

rsp.CursorType = adOpenKeyset
rsp.LockType = adLockOptimistic
rsp.ActiveConnection = Cn
rsp.Source = "Select * from nuevos WHERE NueId = " & vid
rsp.Open

seleccion = 1
SSFrame5.Visible = True
SSFrame6.Visible = True

If rsp!NueTip = "PREOCUPACIONAL" Then
    SSOption3.Value = -1
ElseIf rsp!NueTip = "POSTOCUPACIONAL" Then
    SSOption4.Value = -1
ElseIf rsp!NueTip = "REPROGRAMACION" Then
    SSOption5.Value = -1
End If

If Len(Trim(rsp!NueRec)) > 0 Then
    SSOption6.Value = -1
    Text4.Text = rsp!NueRec
End If
If Len(Trim(rsp!NueSigep)) > 0 Then
    SSOption1.Value = -1
    Text7.Text = rsp!NueSigep
End If
If Len(Trim(rsp!NueCobro)) > 0 Then
    SSOption2.Value = -1
    Text2.Text = rsp!NueCobro
End If
SSFrame5.Enabled = False
SSFrame6.Enabled = False



Text1.Text = rsp!NueApa
Text9.Text = rsp!NueAma
Text10.Text = rsp!NueNombres
Combo2.Text = rsp!NueSex
SSDateCombo2.Text = rsp!NueFeN & ""
Text5.Text = rsp!NueEda
Text3.Text = rsp!NueTeI
Text6.Text = rsp!NueTeR
Combo1.Text = rsp!EmpDes
SSDateCombo3.Text = rsp!NueFeI & ""
Label20.Caption = rsp!NueFeS
SSDateCombo1.Text = rsp!NueFeP & ""
SSDateCombo4.Text = rsp!NueFeM & ""
SSDateCombo5.Text = rsp!NueFeD & ""
If Combo2.Text = "FEMENINO" Then
    SSCheck1.Visible = True
    If rsp!NueEmb = -1 Then
        SSCheck1.Value = -1
    Else
        SSCheck1.Value = 0
    End If
Else
    SSCheck1.Visible = False
End If
If rsp!NueEme = -1 Then
    SSCheck2.Value = -1
Else
    SSCheck2.Value = 0
End If

If rsp!NueTip = "REPROGRAMACION" Then
    Combo3.Enabled = True
    Combo3.Text = rsp!RepTip & ""
End If

vnueid = rsp!NueId
SSFrame3.Enabled = True
SSFrame2.Visible = True
Label1.Enabled = True
Label2.Enabled = True
End Sub

Private Sub SSOption1_Click(Value As Integer)
If SSOption1.Value = -1 Then
    Text4.Text = ""
    Text2.Text = ""
    Text7.Enabled = True
    Text2.Enabled = False
    Text4.Enabled = False
    If SSFrame1.Enabled Then
        Text7.SetFocus
    End If
End If
End Sub

Private Sub SSOption2_Click(Value As Integer)
If SSOption2.Value = -1 Then
    Text4.Text = ""
    Text7.Text = ""
    Text7.Enabled = False
    Text2.Enabled = True
    Text4.Enabled = False
    If SSFrame1.Enabled Then
        Text2.SetFocus
    End If
End If
End Sub

Private Sub SSOption3_Click(Value As Integer)
If SSOption3.Value = -1 Then
    Combo3.Text = ""
    Combo3.Enabled = False
End If
End Sub

Private Sub SSOption4_Click(Value As Integer)
If SSOption4.Value = -1 Then
    Combo3.Text = ""
    Combo3.Enabled = False
End If
End Sub

Private Sub SSOption5_Click(Value As Integer)
If SSOption5.Value = -1 Then
    Combo3.Enabled = True
    Combo3.ListIndex = 0
'    Combo3.SetFocus
End If
End Sub

Private Sub SSOption6_Click(Value As Integer)
If SSOption6.Value = -1 Then
    Text7.Text = ""
    Text2.Text = ""
    Text7.Enabled = False
    Text2.Enabled = False
    Text4.Enabled = True
    If SSFrame1.Enabled Then
       Text4.SetFocus
    End If
End If
End Sub

Private Sub SSOption7_Click(Value As Integer)
If SSOption7.Value = -1 Then
    Text12.Text = ""
    Text14.Text = ""
    Text12.Enabled = False
    Text14.Enabled = False
    Text13.Enabled = True
    Text12.BackColor = &H80000005
    Text14.BackColor = &H80000005
    Text13.BackColor = &HC0FFFF
    Text13.SelStart = 0
    Text13.SelLength = Len(Text13.Text)
    Text13.SetFocus
End If
End Sub

Private Sub SSOption8_Click(Value As Integer)
If SSOption8.Value = -1 Then
    Text12.Text = ""
    Text13.Text = ""
    Text12.Enabled = False
    Text13.Enabled = False
    Text14.Enabled = True
    Text14.BackColor = &HC0FFFF
    Text14.SelStart = 0
    Text14.SelLength = Len(Text14.Text)
    Text14.SetFocus
End If

End Sub

Private Sub SSOption9_Click(Value As Integer)
If SSOption9.Value = -1 Then
    Text13.Text = ""
    Text14.Text = ""
    Text13.Enabled = False
    Text14.Enabled = False
    Text13.BackColor = &H80000005
    Text14.BackColor = &H80000005
    Text12.Enabled = True
    Text12.BackColor = &HC0FFFF
    Text12.SelStart = 0
    Text12.SelLength = Len(Text12.Text)
    Text12.SetFocus
End If

End Sub

Private Sub Text1_GotFocus()
Text1.BackColor = &HC0FFFF
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub
Private Sub Text1_LostFocus()
Text1.BackColor = &HFFFFFF
Text1.Text = UCase(Text1.Text)
End Sub

Private Sub Text10_GotFocus()
Text10.BackColor = &HC0FFFF
Text10.SelStart = 0
Text10.SelLength = Len(Text10.Text)
End Sub
Private Sub Text10_LostFocus()
Text10.BackColor = &HFFFFFF
Text10.Text = UCase(Text10.Text)
End Sub

Private Sub Text11_GotFocus()
Text11.BackColor = &HC0FFFF
Text11.SelStart = 0
Text11.SelLength = Len(Text11.Text)
End Sub
Private Sub Text11_LostFocus()
Text11.BackColor = &HFFFFFF
Text11.Text = UCase(Text11.Text)
End Sub

Private Sub Text15_GotFocus()
Text15.BackColor = &HC0FFFF
End Sub
Private Sub Text15_LostFocus()
Text15.BackColor = &HFFFFFF
Text15Text = UCase(Text15Text)
End Sub
Private Sub Text3_GotFocus()
Text3.BackColor = &HC0FFFF
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub
Private Sub Text3_LostFocus()
Text3.BackColor = &HFFFFFF
End Sub
Private Sub Text4_GotFocus()
Text4.BackColor = &HC0FFFF
End Sub
Private Sub Text4_LostFocus()
Text4.BackColor = &HFFFFFF
End Sub
Private Function limpiadatos()
Text1.Text = ""
'Text3.Text = ""
'Text4.Text = ""
Text5.Text = ""
'Text6.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Combo3.Text = ""
Combo2.Text = ""
SSCheck1.Value = 0
SSCheck1.Visible = False
SSCheck2.Value = 0
SSCheck4.Visible = True
'SSOption6.Value = 0
'SSDateCombo1.Text = Date
SSDateCombo2.Text = Date
SSDateCombo3.Text = Date
'SSDateCombo4.Text = Date
Label20.Caption = ""
seleccion = 0
SSFrame2.Visible = False
End Function

Private Sub Text5_GotFocus()
Text5.BackColor = &HC0FFFF
edad
End Sub
Private Sub Text5_LostFocus()
Text5.BackColor = &HFFFFFF
End Sub

Private Sub Text6_GotFocus()
If Len(Trim(Text3.Text)) > 0 Then
    If Len(Text3.Text) < 8 Then
        MsgBox "Número de digitos incorrecto", vbInformation, EmpresaBox
        Text3.SetFocus
    End If
End If
Text6.BackColor = &HC0FFFF
Text6.SelStart = 0
Text6.SelLength = Len(Text6.Text)
End Sub
Private Sub Text6_LostFocus()
Text6.BackColor = &HFFFFFF
Text6.Text = UCase(Text6.Text)
End Sub

Private Function edad()
'Dim vfecnac As Date
vmatr = Trim(SSDateCombo2.Text)
vyear = Right(vmatr, 2)
vday = Left(SSDateCombo2.Text, 2)
vmonth = Mid(SSDateCombo2.Text, 4, 2)
vfecnac = vday & "/" & vmonth & "/" & vyear
vfecnac = CDate(vfecnac)
vedad = (Date - vfecnac) / 365
Text5.Text = Int(vedad)
End Function
Private Sub Text7_GotFocus()
Text7.BackColor = &HC0FFFF
End Sub
Private Sub Text7_LostFocus()
Text7.BackColor = &HFFFFFF
End Sub
Private Sub Text2_GotFocus()
Text2.BackColor = &HC0FFFF
End Sub
Private Sub Text2_LostFocus()
Text2.BackColor = &HFFFFFF
Text2.Text = UCase(Text2.Text)
End Sub
Private Sub Text8_GotFocus()
Text8.BackColor = &HC0FFFF
End Sub
Private Sub Text8_LostFocus()
Text8.BackColor = &HFFFFFF
Text8.Text = UCase(Text8.Text)
End Sub
Private Function verificafecha()
vfechap = Format(SSDateCombo1.Text, "yyyy-mm-dd")
Adodc2.RecordSource = "SELECT * FROM nuevos WHERE NUeFeP = " & "'" & vfechap & "'"
Adodc2.Refresh

If Not Adodc2.Recordset.EOF Then
    If Adodc2.Recordset.RecordCount = 30 Then
        MsgBox "Se llegó al límite de programaciones para esta fecha", vbInformation, EmpresaBox
    End If
End If

End Function

Private Sub Text9_GotFocus()
Text9.BackColor = &HC0FFFF
Text9.SelStart = 0
Text9.SelLength = Len(Text9.Text)
End Sub
Private Sub Text9_LostFocus()
Text9.BackColor = &HFFFFFF
Text9.Text = UCase(Text9.Text)
End Sub

Private Function codigoQR()
  Dim Matrix() As Integer
  Dim QRSize As Integer
  Dim Y, X As Long
  Dim intScale As Integer
  Dim off As Single
  
  Dim vbQRObj As vbQRCode
  Dim strBanner As String
  Dim strValue As String
  Dim lngColor As Long
  Dim fb As Integer
  
  picCode.Cls
 
  Set vbQRObj = New vbQRCode

'ModCGS
    cmbError.ListIndex = 1
    chkBMP.Value = vbChecked
    chkClipboard = 0
  Select Case cmbError.ListIndex
    Case 0
      vbQRObj.ErrorLevel = qrLevelL
    Case 1
      vbQRObj.ErrorLevel = qrLevelM
    Case 2
      vbQRObj.ErrorLevel = qrLevelQ
    Case 3
      vbQRObj.ErrorLevel = qrLevelH
  End Select
  
  vbQRObj.FindBestMask = (chkFindBestMask.Value = vbChecked)
  vbQRObj.ShowMarkers = (chkShowMarkers.Value = vbChecked)
  vbQRObj.QuietZone = cmbQuietZone.ListIndex
  
  'strValue = txtCode.Text
  strValue = codigoTXT
  
  If (vbQRObj.Encode(strValue)) Then
    QRSize = vbQRObj.Size
    
    If cmbScale.ListIndex = 0 Then
      intScale = Int(picCode.ScaleWidth / QRSize)
    Else
      intScale = cmbScale.ListIndex * 5
    End If
    
    off = (picCode.ScaleWidth - intScale * QRSize) / 2
    
    Matrix() = vbQRObj.Matrix()
    For Y = 0 To QRSize - 1
      For X = 0 To QRSize - 1
        lngColor = vbWhite
        Select Case Matrix(Y, X)
          Case 1
            lngColor = vbBlack
          Case 2
            lngColor = vbRed
          Case 3
            lngColor = vbGreen
          Case 4
            lngColor = vbYellow
          Case 5
            lngColor = vbBlue
          Case Else
        End Select
        picCode.Line (off + X * intScale, off + Y * intScale)-Step(intScale, intScale), lngColor, BF
      Next
    Next
    
    '' Add a banner on top*******************************************
    'strBanner = "http://www.luigimicco.altervista.org"
    With picCode
      .FontSize = 13
      .ForeColor = vbRed
      .CurrentX = (.ScaleWidth - .TextWidth(strBanner)) / 2
      .CurrentY = .ScaleHeight / 2 - .TextHeight("A")
      picCode.Print strBanner
    End With
    ''
    
    '' Copy to clipboard
    If chkClipboard.Value = vbChecked Then
      Clipboard.Clear
      Clipboard.SetData picCode.Image, vbCFBitmap
    End If
 
    '' Save to EPS format file
    If (chkEPS.Value = vbChecked) Then
      fb = FreeFile
      Open App.Path & "\qrcode.eps" For Output As #fb
      Print #fb, vbQRObj.GetEPS(intScale)
      Close #fb
    End If
      
    '' Save to SVG format file
    If (chkSVG.Value = vbChecked) Then
      fb = FreeFile
      Open App.Path & "\qrcode.svg" For Output As #fb
      Print #fb, vbQRObj.GetSVG(intScale)
      Close #fb
    End If
 
    '' Save to HTML format file
    If (chkHTML.Value = vbChecked) Then
      fb = FreeFile
      Open App.Path & "\qrcode.html" For Output As #fb
      Print #fb, vbQRObj.GetHTML(intScale)
      Close #fb
    End If
      
    '' Save to BMP format file
    If (chkBMP.Value = vbChecked) Then
      fb = FreeFile
      Open App.Path & "\qrcode.bmp" For Output As #fb
      Print #fb, vbQRObj.GetBMP(intScale)
      Close #fb
    End If
 
 
  End If
  Set vbQRObj = Nothing

End Function

