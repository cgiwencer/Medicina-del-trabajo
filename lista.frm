VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3150
      Top             =   3555
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
      RecordSource    =   "Select * from implisd  ORDER BY NueNom"
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
      Bindings        =   "lista.frx":0000
      Height          =   4485
      Left            =   135
      TabIndex        =   1
      Top             =   765
      Width           =   11025
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
      ExtraHeight     =   106
      Columns.Count   =   4
      Columns(0).Width=   6641
      Columns(0).Caption=   "NueNom"
      Columns(0).Name =   "NueNom"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "NueNom"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   5265
      Columns(1).Caption=   "EmpDes"
      Columns(1).Name =   "EmpDes"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "EmpDes"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1958
      Columns(2).Caption=   "NueFeM"
      Columns(2).Name =   "NueFeM"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   1
      Columns(2).DataField=   "NueFeM"
      Columns(2).DataType=   7
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "NueTip"
      Columns(3).Name =   "NueTip"
      Columns(3).CaptionAlignment=   0
      Columns(3).DataField=   "NueTip"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      _ExtentX        =   19447
      _ExtentY        =   7911
      _StockProps     =   79
      Caption         =   "SSOleDBGrid1"
      ForeColor       =   16777215
      BackColor       =   16777215
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
   Begin VB.Image Image22 
      Height          =   600
      Left            =   45
      Picture         =   "lista.frx":0015
      Stretch         =   -1  'True
      Top             =   45
      Width           =   600
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "LISTADO PARA AFILIACIONES"
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
      Left            =   7950
      TabIndex        =   0
      Top             =   210
      Width           =   3255
   End
   Begin VB.Image Image14 
      Height          =   690
      Left            =   0
      Picture         =   "lista.frx":AA72
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11625
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
