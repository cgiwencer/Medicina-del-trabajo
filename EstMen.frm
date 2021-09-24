VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form EstMen 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2700
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
   Begin Threed.SSFrame SSFrame1 
      Height          =   1950
      Left            =   225
      TabIndex        =   4
      Top             =   855
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   3440
      _Version        =   196608
      BackStyle       =   1
      ClipControls    =   0   'False
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "EstMen.frx":0000
         Left            =   4050
         List            =   "EstMen.frx":0013
         TabIndex        =   2
         Top             =   630
         Width           =   1500
      End
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo2 
         Height          =   375
         Left            =   2100
         TabIndex        =   1
         Top             =   615
         Width           =   1635
         _Version        =   65537
         _ExtentX        =   2884
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483633
      End
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo1 
         Height          =   375
         Left            =   225
         TabIndex        =   0
         Top             =   615
         Width           =   1635
         _Version        =   65537
         _ExtentX        =   2884
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483633
      End
      Begin VB.Image Image5 
         Height          =   285
         Left            =   3285
         Picture         =   "EstMen.frx":0035
         Stretch         =   -1  'True
         Top             =   1350
         Width           =   345
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
         Left            =   630
         TabIndex        =   8
         Top             =   1365
         Width           =   1815
      End
      Begin VB.Image Image6 
         Height          =   285
         Left            =   540
         Picture         =   "EstMen.frx":27AC
         Stretch         =   -1  'True
         Top             =   1350
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione Gestión"
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
         Left            =   3930
         TabIndex        =   7
         Top             =   405
         Width           =   1665
      End
      Begin VB.Label Label7 
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
         TabIndex        =   6
         Top             =   390
         Width           =   1335
      End
      Begin VB.Label Label6 
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
         TabIndex        =   5
         Top             =   390
         Width           =   1005
      End
      Begin VB.Image Image7 
         Height          =   465
         Left            =   405
         Picture         =   "EstMen.frx":580E
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   1260
         Width           =   1965
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
         Left            =   3330
         TabIndex        =   9
         Top             =   1350
         Width           =   1920
      End
      Begin VB.Image Image8 
         Height          =   465
         Left            =   3150
         Picture         =   "EstMen.frx":7DC3
         Stretch         =   -1  'True
         ToolTipText     =   "Nuevo"
         Top             =   1260
         Width           =   1965
      End
   End
   Begin VB.Image Image22 
      Height          =   600
      Left            =   45
      Picture         =   "EstMen.frx":A504
      Stretch         =   -1  'True
      Top             =   45
      Width           =   600
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "REPORTE ESTADISTICO"
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
      Left            =   3525
      TabIndex        =   3
      Top             =   240
      Width           =   2595
   End
   Begin VB.Image Image14 
      Height          =   690
      Left            =   0
      Picture         =   "EstMen.frx":14F61
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "EstMen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
KeyPreview = True
Combo2.ListIndex = 0
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

Private Sub Label1_Click()
Dim total As Integer
total = 0
If Len(Trim(Combo2.Text)) > 0 Then
'******************************************************************************
    'Seleccion del Mes
   ' If Combo1.Text = "ENERO" Then
   '     vmes = 1
   ' ElseIf Combo1.Text = "FEBRERO" Then
   '     vmes = 2
   ' ElseIf Combo1.Text = "MARZO" Then
   '     vmes = 3
   ' ElseIf Combo1.Text = "ABRIL" Then
   '     vmes = 4
   ' ElseIf Combo1.Text = "MAYO" Then
   '     vmes = 5
   ' ElseIf Combo1.Text = "JUNIO" Then
   '     vmes = 6
   ' ElseIf Combo1.Text = "JULIO" Then
   '     vmes = 7
   ' ElseIf Combo1.Text = "AGOSTO" Then
   '     vmes = 8
   ' ElseIf Combo1.Text = "SEPTIEMBRE" Then
   '     vmes = 9
   ' ElseIf Combo1.Text = "OCTUBRE" Then
   '     vmes = 10
   ' ElseIf Combo1.Text = "NOVIEMBRE" Then
   '     vmes = 11
   ' ElseIf Combo1.Text = "DICIEMBRE" Then
   '     vmes = 12
   ' End If
'********************************************************************************
    vfechai = SSDateCombo1.Date
    vfechaf = SSDateCombo2.Date
    vfechai = Format(vfechai, "YYYY-MM-dd")
    vfechaf = Format(vfechaf, "yyyy-mm-dd")
    vfechair = SSDateCombo1.Date
    vfechafr = SSDateCombo2.Date
    vges = Combo2.Text
    
    'Inicia variables de suma
    vMeEMPr = 0
    vMeEMPO = 0
    vMeLaPr = 0
    vMeLaPrS = 0
    vMeLaPO = 0
    vMeLaPOS = 0
    vMeRxPr = 0
    vMeRxPrS = 0
    vMeRxPO = 0
    vMeRxPOS = 0
    vMeRLaPr = 0
    vMeRLaPrS = 0
    vMeRLaPO = 0
    vMeRLaPOS = 0
    vMeRRxPr = 0
    vMeRRxPrS = 0
    vMeRRxPO = 0
    vMeRRxPOS = 0
    vMeRx = 0
    vMeLab = 0
    
    Dim Cn As New ADODB.Connection
    Dim rsla As New ADODB.Recordset ' Recordset de reporte estadistico
    Cn.ConnectionString = Cadena
    Cn.Open
    
    rsla.CursorType = adOpenKeyset
    rsla.LockType = adLockOptimistic
    rsla.ActiveConnection = Cn
    rsla.Source = "Select * from nuevos where FecRev BETWEEN " & "'" & vfechai & "'" & " AND " & "'" & vfechaf & "' AND Year(FecRev) = " & vges
    'rsla.Source = "Select * from nuevos WHERE Month(FecRev) = " & vmes & " AND Year(FecRev) = " & vges
    rsla.Open
    
    If Not rsla.EOF Then
        Do While Not rsla.EOF
            If rsla!NueTip = "PREOCUPACIONAL" Or rsla!RepTip = "PREOCUPACIONAL" Then
                vMeEMPr = vMeEMPr + 1
            End If
            If rsla!NueTip = "POSTOCUPACIONAL" Or rsla!NueTip = "POSTCUPACIONAL" Then
                vMeEMPO = vMeEMPO + 1
            End If
            rsla.MoveNext
        Loop
    End If
    rsla.Close
    
    rsla.CursorType = adOpenKeyset
    rsla.LockType = adLockOptimistic
    rsla.ActiveConnection = Cn
    rsla.Source = "Select * from nuevos where NueFeS BETWEEN " & "'" & vfechai & "'" & " AND " & "'" & vfechaf & "' AND Year(NueFeS) = " & vges
    'rsla.Source = "Select * from nuevos WHERE Month(NueFeS) = " & vmes & " AND Year(NueFeS) = " & vges
    rsla.Open
    
    If Not rsla.EOF Then
        Do While Not rsla.EOF
            'Laboratorio Programacion
            If rsla!NueTip = "PREOCUPACIONAL" And Len(rsla!NueCobro) > 0 Then
                vMeLaPr = vMeLaPr + 1
            End If
            If rsla!NueTip = "PREOCUPACIONAL" And Len(rsla!NueRec) > 0 Then
                vMeLaPr = vMeLaPr + 1
            End If
            If rsla!NueTip = "POSTOCUPACIONAL" And Len(rsla!NueRec) > 0 Then
                vMeLaPO = vMeLaPO + 1
            End If
            If rsla!NueTip = "POSTOCUPACIONAL" And Len(rsla!NueCobro) > 0 Then
                vMeLaPO = vMeLaPO + 1
            End If
            If rsla!NueTip = "PREOCUPACIONAL" And Len(rsla!NueSigep) > 0 Then
                vMeLaPrS = vMeLaPrS + 1
            End If
            If rsla!NueTip = "POSTOCUPACIONAL" And Len(rsla!NueSigep) > 0 Then
                vMeLaPOS = vMeLaPOS + 1
            End If
                        
            'Rayos X Programacion
            If rsla!NueTip = "PREOCUPACIONAL" And Len(rsla!NueRec) > 0 And rsla!NueEmb = 0 Then
                vMeRxPr = vMeRxPr + 1
            End If
            If rsla!NueTip = "PREOCUPACIONAL" And Len(rsla!NueCobro) > 0 And rsla!NueEmb = 0 Then
                vMeRxPr = vMeRxPr + 1
            End If
            If rsla!NueTip = "POSTOCUPACIONAL" And Len(rsla!NueRec) > 0 And rsla!NueEmb = 0 Then
                vMeRxPO = vMeRxPO + 1
            End If
            If rsla!NueTip = "POSTOCUPACIONAL" And Len(rsla!NueCobro) > 0 And rsla!NueEmb = 0 Then
                vMeRxPO = vMeRxPO + 1
            End If
            
            If rsla!NueTip = "PREOCUPACIONAL" And Len(rsla!NueSigep) > 0 And rsla!NueEmb = 0 Then
                vMeRxPrS = vMeRxPrS + 1
            End If
            If rsla!NueTip = "POSTOCUPACIONAL" And Len(rsla!NueSigep) > 0 And rsla!NueEmb = 0 Then
                vMeRxPOS = vMeRxPOS + 1
            End If
            
            'Laboratorio ReProgramacion
            
            If rsla!NueTip = "REPROGRAMACION" And rsla!RepTip = "PREOCUPACIONAL" And Len(rsla!NueCobro) > 0 Then
                vMeRLaPr = vMeRLaPr + 1
            End If
            If rsla!NueTip = "REPROGRAMACION" And rsla!RepTip = "PREOCUPACIONAL" And Len(rsla!NueRec) > 0 Then
                vMeRLaPr = vMeRLaPr + 1
            End If
            If rsla!NueTip = "REPROGRAMACION" And rsla!RepTip = "POSTOCUPACIONAL" And Len(rsla!NueRec) > 0 Then
                vMeRLaPO = vMeRLaPO + 1
            End If
            If rsla!NueTip = "REPROGRAMACION" And rsla!RepTip = "POSTOCUPACIONAL" And Len(rsla!NueCobro) > 0 Then
                vMeRLaPO = vMeRLaPO + 1
            End If
            If rsla!NueTip = "REPROGRAMACION" And rsla!RepTip = "PREOCUPACIONAL" And Len(rsla!NueSigep) > 0 Then
                vMeRLaPrS = vMeRLaPrS + 1
            End If
            If rsla!NueTip = "REPROGRAMACION" And rsla!RepTip = "POSTOCUPACIONAL" And Len(rsla!NueSigep) > 0 Then
                vMeRLaPOS = vMeRLaPOS + 1
            End If
                       
            'Rayos X ReProgramacion
            
            If rsla!NueTip = "REPROGRAMACION" And rsla!RepTip = "PREOCUPACIONAL" And Len(rsla!NueRec) > 0 And rsla!NueEmb = 0 Then
                vMeRRxPr = vMeRRxPr + 1
            End If
            If rsla!NueTip = "REPROGRAMACION" And rsla!RepTip = "PREOCUPACIONAL" And Len(rsla!NueCobro) > 0 And rsla!NueEmb = 0 Then
                vMeRRxPr = vMeRRxPr + 1
            End If
            If rsla!NueTip = "REPROGRAMACION" And rsla!RepTip = "POSTOCUPACIONAL" And Len(rsla!NueRec) > 0 And rsla!NueEmb = 0 Then
                vMeRRxPO = vMeRRxPO + 1
            End If
            If rsla!NueTip = "REPROGRAMACION" And rsla!RepTip = "POSTOCUPACIONAL" And Len(rsla!NueCobro) > 0 And rsla!NueEmb = 0 Then
                vMeRRxPO = vMeRRxPO + 1
            End If
            
            If rsla!NueTip = "REPROGRAMACION" And rsla!RepTip = "PREOCUPACIONAL" And Len(rsla!NueSigep) > 0 And rsla!NueEmb = 0 Then
                vMeRRxPrS = vMeRRxPrS + 1
            End If
            If rsla!NueTip = "REPROGRAMACION" And rsla!RepTip = "POSTOCUPACIONAL" And Len(rsla!NueSigep) > 0 And rsla!NueEmb = 0 Then
                vMeRRxPOS = vMeRRxPOS + 1
            End If
            
            rsla.MoveNext
        Loop
    End If
    rsla.Close
    
    rsla.CursorType = adOpenKeyset
    rsla.LockType = adLockOptimistic
    rsla.ActiveConnection = Cn
    rsla.Source = "Select * from nuevos where FecLab BETWEEN " & "'" & vfechai & "'" & " AND " & "'" & vfechaf & "' AND Year(FecLab) = " & vges
    'rsla.Source = "Select * from nuevos WHERE Month(FecLab) = " & vmes & " AND Year(FecLab) = " & vges
    rsla.Open
    
    If Not rsla.EOF Then
        Do While Not rsla.EOF
            If rsla!ProEsLa = -1 Then
                vMeLab = vMeLab + 1
            End If
        rsla.MoveNext
        Loop
    End If
    rsla.Close
    
    rsla.CursorType = adOpenKeyset
    rsla.LockType = adLockOptimistic
    rsla.ActiveConnection = Cn
    rsla.Source = "Select * from nuevos where FecRx BETWEEN " & "'" & vfechai & "'" & " AND " & "'" & vfechaf & "' AND Year(FecRx) = " & vges
    'rsla.Source = "Select * from nuevos WHERE Month(FecRx) = " & vmes & " AND Year(FecRx) = " & vges
    rsla.Open
    
    If Not rsla.EOF Then
        Do While Not rsla.EOF
            If rsla!ProEsRx = -1 Then
                vMeRx = vMeRx + 1
            End If
        rsla.MoveNext
        Loop
    End If
    rsla.Close
    
    
    'Borra reporte mensual
    borra = "DELETE FROM Mensual"
    Cn.Execute borra
    
    'Graba en tabla Mensual
    grabam = "INSERT INTO mensual SET MeEMPr = " & vMeEMPr & ", MeEMPO = " & vMeEMPO & ", MeLaPr = " & vMeLaPr & ", MeLaPrS = " & vMeLaPrS & _
    ", MeLaPO = " & vMeLaPO & ", MeLaPOS = " & vMeLaPOS & ", MeRxPr = " & vMeRxPr & ", MeRxPrS = " & vMeRxPrS & ", MeRxPO = " & vMeRxPO & _
    ", MeRxPOS = " & vMeRxPOS & ", MeRLaPr = " & vMeRLaPr & ", MeRLaPrS = " & vMeRLaPrS & ", MeRLaPO = " & vMeRLaPO & ", MeRLaPOS = " & vMeRLaPOS & _
    ", MeRRxPr = " & vMeRRxPr & ", MeRRxPrS = " & vMeRRxPrS & ", MeRRxPO = " & vMeRRxPO & ", MeRRxPOS = " & vMeRRxPOS & ", MeLab = " & vMeLab & ", MeRx = " & vMeRx & _
    ", MeEMPrT = 'REVISIÓN MEDICA PRE OCUPACIONAL', MeEMPOT = 'REVISION MEDICA POST OCUPACIONAL', MeLaPrT = 'PROGRAMACIÓN LABORATORIO PRE OCUPACIONAL'" & _
    ", MeLaPrST = 'PROGRAMACIÓN DE LABORATORIO PRE OCUPACIONAL SIGEP', MeLaPOT = 'PROGRAMACIÓN DE LABORATORIO POST OCUPACIONAL', MeLaPOST = 'PROGRAMACIÓN DE LABORATORIO POST OCUPACIONAL SIGEP', MeRxPrT = 'PROGRAMACIÓN DE RX PRE OCUPACIONAL', MeRxPrST = 'PROGRAMACIÓN DE RX PRE OCUPACIONAL SIGEP', MeRxPOT = 'PROGRAMACIÓN DE RX POST OCUPACIONAL', MeRxPOST = 'PROGRAMACIÓN DE RX POST OCUPACIONAL SIGEP', MeRLaPrT='REPROGRAMACIÓN LABORATORIO PRE OCUPACIONAL', MeRLaPrST='REPROGRAMACIÓN DE LABORATORIO PRE OCUPACIONAL SIGEP', MeRLaPOT = 'REPROGRAMACIÓN DE LABORATORIO POST OCUPACIONAL', MeRLaPOST = 'REPROGRAMACIÓN DE LABORATORIO POST OCUPACIONAL SIGEP', MeRRxPrT = 'REPROGRAMACIÓN DE RX PRE OCUPACIONAL', MeRRxPrST= 'REPROGRAMACIÓN DE RX PRE OCUPACIONAL SIGEP', MeRRxPOT = 'REPROGRAMACIÓN DE RX POST OCUPACIONAL', MeRRxPOST = 'REPROGRAMACIÓN DE RX POST OCUPACIONAL SIGEP'" & _
    ", MeLabt = 'REGISTRO DE EXAMENES DE LABORATORIO', MeRxT = 'REGISTRO DE EXAMENES DE RX'"
    Cn.Execute grabam
    
    CrystalReport1.ReportFileName = App.Path & "\Mensual.rpt"
    CrystalReport1.Formulas(1) = "FechaI = " & "'" & vfechair & "'"
    CrystalReport1.Formulas(0) = "FechaF = " & "'" & vfechafr & "'"
    CrystalReport1.Action = 1
Else
    MsgBox "Debe seleccionar un mes y un año", vbInformation, empresa
End If

End Sub

Private Sub Label10_Click()
Principal.Enabled = True
Unload EstMen
Set EstMen = Nothing
End Sub
