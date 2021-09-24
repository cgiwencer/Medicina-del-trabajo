Attribute VB_Name = "Variables"
'Variables de coneccion
Public Cn As New ADODB.Connection
'Public Const Cadena = "server = localhost;driver=MySQL ODBC 3.51 Driver;db=MedTrabajo;UID=root;PWD=cagisa"
Public Const Cadena = "server = 172.0.1.36;driver=MySQL ODBC 3.51 Driver;db=MedTrabajo;UID=root;PWD=cagisa"
'Public Const Cadena = "server = 192.168.200.15;driver=MySQL ODBC 3.51 Driver;db=MedTrabajo;UID=root;PWD=cagisa"
Public Const EmpresaBox = "CAJA PETROLERA DE SALUD"
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public resp As Long

'Variable de caracter
Public modo, modop, vUsuario, vusuariot, vusuacc, origen, vtipo, varea, vnivel, vtipove, codigoTXT As String
Public textomarquesina, direccionRx, direccionlab As String

'Variables numericas
Public seleccion, seleccioni, selecciont, vusucod, siaccesos, siexiste, venccod, vnueid, vnueids, vexiste As Integer
Public uactivado, vconini, vflag, vlcocod1, vEquId, vid As Integer
Public vnumero, literal, vempresades As String
Public pnumero As Double

'Variables Booleanas
Public VarOrder As Boolean


