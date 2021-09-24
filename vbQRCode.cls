VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbQRCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================
' Descrizione.....: Classe per la generazione di QRCode
'                   Class to create QRCode
' Nome dei File...: vbQRCode.cls
' Data............: 21/06/2014
' Versione........: 2013.2d
' Sistema.........: Visual Basic 6.0 Pro - SP 6
' Scritto da......: Luigi Micco
' E-Mail..........: l.micco(at)tiscali.it
'==================================
'==================================
' (C) 2012  - L'uso di questo software è consentito solo su espressa
'     2013    autorizzazione dell'autore. Non puo' essere copiato o
'             ridistribuito, ne' integralmente, ne' parzialmente.
'
'             The use of this software is allowed only on express
'             authorization of the author. It's cannot be copied or
'             redistributed, neither integrally, neither partially.
'==================================

Option Explicit

Public Enum vbQRCodeLevel
  qrLevelL = 0 ' Error correction level L : About 7% or less errors can be corrected.
  qrLevelM = 1 ' Error correction level M : About 15% or less errors can be corrected.
  qrLevelQ = 2 ' Error correction level Q : About 25% or less errors can be corrected.
  qrLevelH = 3 ' Error correction level H : About 30% or less errors can be corrected.
End Enum

Public Enum vbQREncodingMode
  qr8BitMode = 0  '' Encoding mode 8bit byte data. In theory, 2953 characters or less can be stored in a QRcode.
  qrKanjiMode = 1 '' Encoding mode KANJI. A KANJI character (multibyte character) is encoded to 13bit length. In theory, 1817 characters or less can be stored in a QRcode.
End Enum


Private Const qrModeNL As Integer = -1 '' Encoding mode
Private Const qrModeNM As Integer = 0 '' Encoding mode numeric (0-9). 3 characters are encoded to 10bit length. In theory, 7089 characters or less can be stored in a QRcode.
Private Const qrModeAN As Integer = 1  '' Encoding mode alphanumeric (0-9A-Z %*+-./:) 45characters. 2 characters are encoded to 11bit length. In theory, 4296 characters or less can be stored in a QRcode.
Private Const qrMode8B As Integer = 2 '' Encoding mode 8bit byte data. In theory, 2953 characters or less can be stored in a QRcode.
Private Const qrModeKJ As Integer = 3  '' Encoding mode KANJI. A KANJI character (multibyte character) is encoded to 13bit length. In theory, 1817 characters or less can be stored in a QRcode.

Private Const mvarVersionMax As Integer = 40 '' Maximum QR Code version.
Private Const mvarWidthMax As Integer = 177 '' Maximum matrix size for maximum version (version 40 is 177*177 matrix).

Private Const qrCapwidth As Integer = 0   '' Matrix index to get width from capacity array.
Private Const qrCapWords As Integer = 1    '' Matrix index to get number of words from capacity array.
Private Const qrCapReminder As Integer = 2 '' Matrix index to get remainder from capacity array.
Private Const qrCapEc As Integer = 3       '' Matrix index to get error correction level from capacity array.

Private Const qrMaskN1 As Integer = 3   '' Down point base value for case 1 mask pattern (concatenation of same color in a line or a column)
Private Const qrMaskN2 As Integer = 3  '' Down point base value for case 2 mask pattern (module block of same color)
Private Const qrMaskN3 As Integer = 40 '' Down point base value for case 3 mask pattern (1:1:3:1:1(dark:bright:dark:bright:dark)pattern in a line or a column)
Private Const qrMaskN4 As Integer = 10 '' Down point base value for case 4 mask pattern (ration of dark modules in whole)

Private Const qrDefaultMask As Integer = 2     '' used when mvarFindBestMask = false

Private mvarFindBestMask As Boolean '' default is true and estimates best mask (but extremally slow), set to false to significant performance boost but (propably) worst quality code

Private Const qrFindFromRandom As Integer = 0  '' if 0, checks all masks available, otherwise value tells count of masks need to be checked, mask id are got randomly

Private mvarSize As Integer

Private mvarDataStr As String '' Input data string.
Private mvarDataCode() As Byte '' Data code.
Private mvarVersion As Integer  '' QR code version. Size of QRcode is defined as version. Version is from 1 to 40. Version 1 is 21*21 matrix. And 4 modules increases whenever 1 version increases. So version 40 is 177*177 matrix.

Private mvarLengthTableBits() As Variant '' Array Length indicator.
Private mvarAlTable() As Variant  '' Alphabet-numeric convesion table.
Private mvarAlignPattern() As Variant  '' This array includes only the second and the third position of the alignment patterns. Rest of them can be calculated from the distance between them.
Private mvarVersionPattern() As Variant '' Array Version information pattern (BCH coded).
Private mvarFormatInfo() As Variant '' Array Format information
Private mvarEccTable(40) As Variant '' Array Table of the error correction code (Reed-Solomon block).
Private mvarCapacity(40) As Variant '' Array Table of the mvarCapacity of symbols.
Private mvarMarker(4) As Variant '' Array Table of alignment marker
Private mvarFinder(6) As Variant '' Array Table of finder pattern
Private mvarQuietZone As Integer  '' size, in module, of Quiet Zone
Private mvarFinal() As Integer '' Barcode array to be returned which is readable by VB.

Private mvarLevel As Integer '' Levels of error correction. See definitions for possible values.
Private mvarEncodingMode As vbQREncodingMode  '' Encoding mode.
Private mvarCaseSensitive As Boolean '' Boolean flag, if true the input string will be converted to uppercase.
Private mvarShowMarkers As Boolean '' Boolean flag, if true all markers will have a specific value.

Private mvarData() As Byte '' Mask data.
Private mvarFrame() As Byte  '' mvarFrame.

Private Type varRSItems
  Pad As Integer
  NRoots As Integer
  mm As Integer
  nn As Integer
  GFpoly As Integer
  Fcr As Integer
  Prim As Integer
  iprim As Integer
  Index_of() As Integer
  Alpha_to() As Integer
  genpoly() As Integer
End Type

Private mvarRSItems() As varRSItems '= array()'' Reed-Solomon items.

Private Type varInputItems
  mode As Integer
  Size As Integer
  data() As Byte
  bstream() As Byte
End Type

Private mvarItems() As varInputItems '' Input items.

Private Type varRSBlocks
  DataLength As Integer
  data() As Byte
  EccLength As Integer
  Ecc() As Byte
End Type

Private mvarRsBlocks() As varRSBlocks ''of RSblock'' Reed-Solomon blocks.

'' Encode a text
Public Function Encode(ByVal Value As String) As Boolean

  Dim i As Integer
  Dim k As Integer
  Dim Size As Integer
  
  If (Trim(Value) = "") Or _
     ((mvarEncodingMode <> qr8BitMode) And (mvarEncodingMode <> qrKanjiMode)) Or _
     ((mvarVersion < 0) Or (mvarVersion > mvarVersionMax)) Then Exit Function
  
  mvarDataStr = Value
  If (Not mvarCaseSensitive) Then Call zzToUpper
  
  If zzStr2DataCode() Then
    If zzEncodeData() Then
      Size = UBound(mvarData)
      
      mvarSize = Size + 1 + 2 * mvarQuietZone
      ReDim mvarFinal(mvarSize - 1, mvarSize - 1)
      
      For i = 0 To Size
        For k = 0 To Size
          mvarFinal(mvarQuietZone + k, mvarQuietZone + i) = IIf((mvarData(k, i) And 1) <> 0, 1, 0)
        Next
      Next
      If mvarShowMarkers Then Call zzShowMarkers
      Encode = True
    End If
  End If
End Function

Public Function GetSVG(ByVal ModuleSize As Integer) As String
  Dim strSVG As String
  Dim Y As Integer
  Dim X As Integer
  
  strSVG = "<?xml version='1.0' standalone='no'?>" & vbCrLf
  strSVG = strSVG & "<!DOCTYPE svg PUBLIC '-//W3C//DTD SVG 1.1//EN' 'http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd'>" & vbCrLf
  strSVG = strSVG & "<svg width='" & CStr(mvarSize * ModuleSize) & "' height='" & CStr(mvarSize * ModuleSize) & "' version='1.1' xmlns='http://www.w3.org/2000/svg'>" & vbCrLf
  strSVG = strSVG & "  <g id='qrcode'>" & vbCrLf
  strSVG = strSVG & "    <g id='background' fill='#ffffff' stroke='#ffffff'>" & vbCrLf
  strSVG = strSVG & "      <rect x='0' y='0' width='" & (mvarSize * ModuleSize) & "' height='" & CStr(mvarSize * ModuleSize) & "' />" & vbCrLf
  strSVG = strSVG & "    </g>" & vbCrLf
  strSVG = strSVG & "    <g id='pattern' fill='#000000' stroke='#000000'>" & vbCrLf
  '' print barcode elements
  For Y = 0 To mvarSize - 1 '' for each row
    For X = 0 To mvarSize - 1 '' for each column
      If mvarFinal(Y, X) >= 1 Then '' draw a single barcode cell
        strSVG = strSVG & "    <g id='R" & CStr(Y) & "C" & CStr(X) & "' >" & vbCrLf
        strSVG = strSVG & "      <rect x='" & CStr(X * ModuleSize) & "' y='" & CStr(Y * ModuleSize) & "' width='" & CStr(ModuleSize) & "' height='" & CStr(ModuleSize) & "' />" & vbCrLf
        strSVG = strSVG & "    </g>" & vbCrLf
      End If
    Next
  Next
  strSVG = strSVG & "    </g>" & vbCrLf
  strSVG = strSVG & "  </g>" & vbCrLf
  strSVG = strSVG & "</svg>" & vbCrLf
  
  GetSVG = strSVG

End Function

Public Function GetEPS(ByVal ModuleSize As Integer) As String

  Dim strEPS As String
  Dim Y As Integer
  Dim X As Integer
  Dim n As Integer


  strEPS = "%!PS-Adobe-3.0 EPSF-3.0" & vbCrLf
  strEPS = strEPS & "%%Creator: vbQRCode 2013.2" & vbCrLf
  strEPS = strEPS & "%%Title: qrcode" & vbCrLf
  strEPS = strEPS & "%%CreationDate: " & Replace(Format(Now(), "ddd mmm d hh:mm:ss yyyy"), ".", ":") & vbCrLf
  strEPS = strEPS & "%%BoundingBox: 0 0 " & CStr(mvarSize * ModuleSize) & " " & CStr(mvarSize * ModuleSize) & vbCrLf
  strEPS = strEPS & "%%DocumentData: Clean7Bit" & vbCrLf
  strEPS = strEPS & "%%EndComments" & vbCrLf
  strEPS = strEPS & "%%BeginProlog" & vbCrLf
  strEPS = strEPS & "/bd { bind def } bind def" & vbCrLf
  strEPS = strEPS & "/incompound false def" & vbCrLf
  strEPS = strEPS & "/m { moveto } bd" & vbCrLf
  strEPS = strEPS & "/l { lineto } bd" & vbCrLf
  strEPS = strEPS & "/*U { /incompound false def f} bd" & vbCrLf
  strEPS = strEPS & "/k { setcmykcolor } bd" & vbCrLf
  
  strEPS = strEPS & "/c { curveto } bd" & vbCrLf
  strEPS = strEPS & "/F { incompound not {fill} if } bd" & vbCrLf
  strEPS = strEPS & "/f { closepath F } bd" & vbCrLf
  strEPS = strEPS & "/S { stroke } bd" & vbCrLf
  strEPS = strEPS & "/*u { /incompound true def } bd" & vbCrLf
  strEPS = strEPS & "/K { k } bd" & vbCrLf
  
  strEPS = strEPS & "%%EndProlog" & vbCrLf
  strEPS = strEPS & "%%BeginSetup" & vbCrLf
  strEPS = strEPS & "%%EndSetup" & vbCrLf
  strEPS = strEPS & "0.000 0.000 0.000 0.000 k" & vbCrLf
  strEPS = strEPS & "0 0 m" & vbCrLf
  strEPS = strEPS & "0 0 l" & vbCrLf
  strEPS = strEPS & CStr(mvarSize * ModuleSize) & " 0 l" & vbCrLf
  strEPS = strEPS & CStr(mvarSize * ModuleSize) & " " & CStr(mvarSize * ModuleSize) & " l" & vbCrLf
  strEPS = strEPS & "0 " & CStr(mvarSize * ModuleSize) & " l" & vbCrLf
  strEPS = strEPS & "f" & vbCrLf
  strEPS = strEPS & "*U" & vbCrLf
  
  n = mvarSize * ModuleSize
  strEPS = strEPS & "0.000 0.000 0.000 1.000 k" & vbCrLf
  For Y = 0 To mvarSize - 1
    For X = 0 To mvarSize - 1
      If mvarFinal(Y, X) >= 1 Then
        strEPS = strEPS & CStr(X * ModuleSize) & " " & CStr(n - Y * ModuleSize) & " m" & vbCrLf
        strEPS = strEPS & CStr(X * ModuleSize) & " " & CStr(n - Y * ModuleSize) & " l" & vbCrLf
        strEPS = strEPS & CStr((X + 1) * ModuleSize) & " " & CStr(n - Y * ModuleSize) & " l" & vbCrLf
        strEPS = strEPS & CStr((X + 1) * ModuleSize) & " " & CStr(n - (Y + 1) * ModuleSize) & " l" & vbCrLf
        strEPS = strEPS & CStr((X) * ModuleSize) & " " & CStr(n - (Y + 1) * ModuleSize) & " l" & vbCrLf
        strEPS = strEPS & "f" & vbCrLf
      End If
    Next
  Next
  
  strEPS = strEPS & "*U" & vbCrLf
  strEPS = strEPS & "%%Trailer" & vbCrLf
  strEPS = strEPS & "%%EOF" & vbCrLf


  GetEPS = strEPS
End Function

Public Function GetHTML(ByVal ModuleSize As Integer) As String

  Dim strHTML As String
  Dim Y As Integer
  Dim X As Integer


  strHTML = "<div style='font-size:0;background-color:#ffffff;position:relative;width:" & CStr(mvarSize * ModuleSize) & "px;height:" & CStr(mvarSize * ModuleSize) & "px;'>" & vbCrLf
  For Y = 0 To mvarSize - 1 '' for each row
    For X = 0 To mvarSize - 1 '' for each column
      If mvarFinal(Y, X) >= 1 Then '' draw a single barcode cell
        strHTML = strHTML & "<div style='background-color:#000000;width:" & CStr(ModuleSize) & "px;height:" & CStr(ModuleSize) & "px;position:absolute;left:" & CStr(X * ModuleSize) & "px;top:" & CStr(Y * ModuleSize) & "px;'>&nbsp;</div>" & vbCrLf
      End If
    Next
  Next
  strHTML = strHTML & "</div>" & vbCrLf

  GetHTML = strHTML
End Function

Public Function GetBMP(ByVal ModuleSize As Integer) As String

  Dim strBMP As String
  Dim Y As Integer
  Dim X As Integer
  Dim k As Integer
  Dim strRow As String
  Dim intTemp As Integer
  Dim strPad As String
  
  Dim Size As Long
  
  Size = mvarSize * ModuleSize
  intTemp = Size Mod 4
  strPad = ""
  If intTemp <> 0 Then
    strPad = String(intTemp, Chr(&H0))
  End If

  strBMP = "BM" & String(8, Chr$(&H0))
  strBMP = strBMP & Chr$(&H36) & String(3, Chr$(&H0))
  strBMP = strBMP & Chr$(&H28) & String(3, Chr$(&H0))
  strBMP = strBMP & Chr$(Size Mod 256) & Chr$(Int(Size / 256)) & String(2, Chr$(&H0)) '' width
  strBMP = strBMP & Chr$(Size Mod 256) & Chr$(Int(Size / 256)) & String(2, Chr$(&H0)) '' height
  strBMP = strBMP & Chr$(1) & Chr$(&H0)
  strBMP = strBMP & Chr$(24) & Chr$(&H0)
  strBMP = strBMP & String(4, Chr$(&H0))
  strBMP = strBMP & String(4, Chr$(&H0))
  
  strBMP = strBMP & Chr$(&H12) & Chr$(&HB) & Chr$(&H0) & Chr$(&H0)
  strBMP = strBMP & Chr$(&H12) & Chr$(&HB) & Chr$(&H0) & Chr$(&H0)
  strBMP = strBMP & String(8, Chr$(&H0))

  For Y = mvarSize - 1 To 0 Step -1 '' for each row
    strRow = ""
    For X = 0 To mvarSize - 1 '' for each column
      If mvarFinal(Y, X) >= 1 Then '' draw a single barcode cell
        strRow = strRow & String(3 * ModuleSize, Chr$(&H0))
      Else
        strRow = strRow & String(3 * ModuleSize, Chr$(&HFF))
      End If
    Next
    
    strRow = strRow & strPad
    For k = 1 To ModuleSize
      strBMP = strBMP & strRow
    Next
  Next

  GetBMP = strBMP
End Function

'---------------------------------------------------------------------------------------
' Size (Read-Only)
'---------------------------------------------------------------------------------------
Public Property Get Size() As Long
  Size = mvarSize
End Property

'---------------------------------------------------------------------------------------
' ErrorLevel
'---------------------------------------------------------------------------------------
Public Property Let ErrorLevel(Value As vbQRCodeLevel)
  mvarLevel = Value
End Property

'---------------------------------------------------------------------------------------
' FindBestMask
'---------------------------------------------------------------------------------------
Public Property Let FindBestMask(Value As Boolean)
  mvarFindBestMask = Value
End Property

'---------------------------------------------------------------------------------------
' ShowMarkers
'---------------------------------------------------------------------------------------
Public Property Let ShowMarkers(Value As Boolean)
  mvarShowMarkers = Value
End Property

'---------------------------------------------------------------------------------------
' EncodingMode
'---------------------------------------------------------------------------------------
Public Property Let EncodingMode(Value As vbQREncodingMode)
  mvarEncodingMode = Value
End Property

'---------------------------------------------------------------------------------------
' CaseSensitive
'---------------------------------------------------------------------------------------
Public Property Let CaseSensitive(Value As Boolean)
  mvarCaseSensitive = Value
End Property

'---------------------------------------------------------------------------------------
' QuietZone
'---------------------------------------------------------------------------------------
Public Property Let QuietZone(Value As Integer)
  If Value >= 0 Then mvarQuietZone = Value
End Property

'---------------------------------------------------------------------------------------
' Matrix (Read-Only)
'---------------------------------------------------------------------------------------
Public Property Get Matrix() As Integer()
  Matrix = mvarFinal
End Property

Private Sub Class_Initialize()

   '' Alphabet-numeric convesion table.
  mvarAlTable = Array(-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
                  -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
                  36, -1, -1, -1, 37, 38, -1, -1, -1, -1, 39, 40, -1, 41, 42, 43, _
                   0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 44, -1, -1, -1, -1, -1, _
                  -1, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, _
                  25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, -1, -1, -1, -1, -1, _
                  -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
                  -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1)

   
  mvarLengthTableBits = Array(Array(10, 12, 14), Array(9, 11, 13), Array(8, 16, 16), Array(8, 10, 12))

   '' Array Positions of alignment patterns.
   '' This array includes only the second and the third position of the alignment patterns. Rest of them can be calculated from the distance between them.
   '' See Table 1 in Appendix E (pp.71) of JIS X0510:2004.
  mvarAlignPattern = Array(Array(0, 0), _
                           Array(0, 0), Array(18, 0), Array(22, 0), Array(26, 0), Array(30, 0), _
                           Array(34, 0), Array(22, 38), Array(24, 42), Array(26, 46), Array(28, 50), _
                           Array(30, 54), Array(32, 58), Array(34, 62), Array(26, 46), Array(26, 48), _
                           Array(26, 50), Array(30, 54), Array(30, 56), Array(30, 58), Array(34, 62), _
                           Array(28, 50), Array(26, 50), Array(30, 54), Array(28, 54), Array(32, 58), _
                           Array(30, 58), Array(34, 62), Array(26, 50), Array(30, 54), Array(26, 52), _
                           Array(30, 56), Array(34, 60), Array(30, 58), Array(34, 62), Array(30, 54), _
                           Array(24, 50), Array(28, 54), Array(32, 58), Array(26, 54), Array(30, 58))

   '' Array Version information pattern (BCH coded).
   '' See Table 1 in Appendix D (pp.68) of JIS X0510:2004.
   '' size: (mvarVersionMax - 6)
  mvarVersionPattern = Array(&H7C94, &H85BC, &H9A99, &HA4D3, &HBBF6, &HC762, &HD847, &HE60D, _
                             &HF928, &H10B78, &H1145D, &H12A17, &H13532, &H149A6, &H15683, &H168C9, _
                             &H177EC, &H18EC4, &H191E1, &H1AFAB, &H1B08E, &H1CC1A, &H1D33F, &H1ED75, _
                             &H1F250, &H209D5, &H216F0, &H228BA, &H2379F, &H24B0B, &H2542E, &H26A64, _
                             &H27541, &H28C69)
 
   '' Array Format information
  mvarFormatInfo = Array(Array(&H77C4, &H72F3, &H7DAA, &H789D, &H662F, &H6318, &H6C41, &H6976), _
                         Array(&H5412, &H5125, &H5E7C, &H5B4B, &H45F9, &H40CE, &H4F97, &H4AA0), _
                         Array(&H355F, &H3068, &H3F31, &H3A06, &H24B4, &H2183, &H2EDA, &H2BED), _
                         Array(&H1689, &H13BE, &H1CE7, &H19D0, &H762, &H255, &HD0C, &H83B))

   '' Array Table of the error correction code (Reed-Solomon block).
   '' See Table 12-16 (pp.30-36), JIS X0510:2004.
  mvarEccTable(0) = Array(Array(0, 0), Array(0, 0), Array(0, 0), Array(0, 0))
  mvarEccTable(1) = Array(Array(1, 0), Array(1, 0), Array(1, 0), Array(1, 0))           ''  1
  mvarEccTable(2) = Array(Array(1, 0), Array(1, 0), Array(1, 0), Array(1, 0))           ''
  mvarEccTable(3) = Array(Array(1, 0), Array(1, 0), Array(2, 0), Array(2, 0))           ''
  mvarEccTable(4) = Array(Array(1, 0), Array(2, 0), Array(2, 0), Array(4, 0))           ''
  mvarEccTable(5) = Array(Array(1, 0), Array(2, 0), Array(2, 2), Array(2, 2))           ''  5
  mvarEccTable(6) = Array(Array(2, 0), Array(4, 0), Array(4, 0), Array(4, 0))           ''
  mvarEccTable(7) = Array(Array(2, 0), Array(4, 0), Array(2, 4), Array(4, 1))           ''
  mvarEccTable(8) = Array(Array(2, 0), Array(2, 2), Array(4, 2), Array(4, 2))           ''
  mvarEccTable(9) = Array(Array(2, 0), Array(3, 2), Array(4, 4), Array(4, 4))           ''
  mvarEccTable(10) = Array(Array(2, 2), Array(4, 1), Array(6, 2), Array(6, 2))           '' 10
  mvarEccTable(11) = Array(Array(4, 0), Array(1, 4), Array(4, 4), Array(3, 8))           ''
  mvarEccTable(12) = Array(Array(2, 2), Array(6, 2), Array(4, 6), Array(7, 4))           ''
  mvarEccTable(13) = Array(Array(4, 0), Array(8, 1), Array(8, 4), Array(12, 4))          ''
  mvarEccTable(14) = Array(Array(3, 1), Array(4, 5), Array(11, 5), Array(11, 5))         ''
  mvarEccTable(15) = Array(Array(5, 1), Array(5, 5), Array(5, 7), Array(11, 7))          '' 15
  mvarEccTable(16) = Array(Array(5, 1), Array(7, 3), Array(15, 2), Array(3, 13))         ''
  mvarEccTable(17) = Array(Array(1, 5), Array(10, 1), Array(1, 15), Array(2, 17))        ''
  mvarEccTable(17) = Array(Array(5, 1), Array(9, 4), Array(17, 1), Array(2, 19))         ''
  mvarEccTable(19) = Array(Array(3, 4), Array(3, 11), Array(17, 4), Array(9, 16))        ''
  mvarEccTable(20) = Array(Array(3, 5), Array(3, 13), Array(15, 5), Array(15, 10))       '' 20
  mvarEccTable(21) = Array(Array(4, 4), Array(17, 0), Array(17, 6), Array(19, 6))        ''
  mvarEccTable(22) = Array(Array(2, 7), Array(17, 0), Array(7, 16), Array(34, 0))        ''
  mvarEccTable(23) = Array(Array(4, 5), Array(4, 14), Array(11, 14), Array(16, 14))      ''
  mvarEccTable(24) = Array(Array(6, 4), Array(6, 14), Array(11, 16), Array(30, 2))       ''
  mvarEccTable(25) = Array(Array(8, 4), Array(8, 13), Array(7, 22), Array(22, 13))       '' 25
  mvarEccTable(26) = Array(Array(10, 2), Array(19, 4), Array(28, 6), Array(33, 4))       ''
  mvarEccTable(27) = Array(Array(8, 4), Array(22, 3), Array(8, 26), Array(12, 28))       ''
  mvarEccTable(28) = Array(Array(3, 10), Array(3, 23), Array(4, 31), Array(11, 31))      ''
  mvarEccTable(29) = Array(Array(7, 7), Array(21, 7), Array(1, 37), Array(19, 26))       ''
  mvarEccTable(30) = Array(Array(5, 10), Array(19, 10), Array(15, 25), Array(23, 25))    '' 30
  mvarEccTable(31) = Array(Array(13, 3), Array(2, 29), Array(42, 1), Array(23, 28))      ''
  mvarEccTable(32) = Array(Array(17, 0), Array(10, 23), Array(10, 35), Array(19, 35))    ''
  mvarEccTable(33) = Array(Array(17, 1), Array(14, 21), Array(29, 19), Array(11, 46))    ''
  mvarEccTable(34) = Array(Array(13, 6), Array(14, 23), Array(44, 7), Array(59, 1))      ''
  mvarEccTable(35) = Array(Array(12, 7), Array(12, 26), Array(39, 14), Array(22, 41))    '' 35
  mvarEccTable(36) = Array(Array(6, 14), Array(6, 34), Array(46, 10), Array(2, 64))      ''
  mvarEccTable(37) = Array(Array(17, 4), Array(29, 14), Array(49, 10), Array(24, 46))    ''
  mvarEccTable(38) = Array(Array(4, 18), Array(13, 32), Array(48, 14), Array(42, 32))    ''
  mvarEccTable(39) = Array(Array(20, 4), Array(40, 7), Array(43, 22), Array(10, 67))     ''
  mvarEccTable(40) = Array(Array(19, 6), Array(18, 31), Array(34, 34), Array(20, 61))     '' 40

   '' Array Table of the capacity of symbols.
   '' See Table 1 (pp.13) and Table 12-16 (pp.30-36), JIS X0510:2004.
  mvarCapacity(0) = Array(0, 0, 0, Array(0, 0, 0, 0))
  mvarCapacity(1) = Array(21, 26, 0, Array(7, 10, 13, 17))
  mvarCapacity(2) = Array(25, 44, 7, Array(10, 16, 22, 28))
  mvarCapacity(3) = Array(29, 70, 7, Array(15, 26, 36, 44))
  mvarCapacity(4) = Array(33, 100, 7, Array(20, 36, 52, 64))
  mvarCapacity(5) = Array(37, 134, 7, Array(26, 48, 72, 88))
  mvarCapacity(6) = Array(41, 172, 7, Array(36, 64, 96, 112))
  mvarCapacity(7) = Array(45, 196, 0, Array(40, 72, 108, 130))
  mvarCapacity(8) = Array(49, 242, 0, Array(48, 88, 132, 156))
  mvarCapacity(9) = Array(53, 292, 0, Array(60, 110, 160, 192))
  mvarCapacity(10) = Array(57, 346, 0, Array(72, 130, 192, 224))
  mvarCapacity(11) = Array(61, 404, 0, Array(80, 150, 224, 264))
  mvarCapacity(12) = Array(65, 466, 0, Array(96, 176, 260, 308))
  mvarCapacity(13) = Array(69, 532, 0, Array(104, 198, 288, 352))
  mvarCapacity(14) = Array(73, 581, 3, Array(120, 216, 320, 384))
  mvarCapacity(15) = Array(77, 655, 3, Array(132, 240, 360, 432))
  mvarCapacity(16) = Array(81, 733, 3, Array(144, 280, 408, 480))
  mvarCapacity(17) = Array(85, 815, 3, Array(168, 308, 448, 532))
  mvarCapacity(18) = Array(89, 901, 3, Array(180, 338, 504, 588))
  mvarCapacity(19) = Array(93, 991, 3, Array(196, 364, 546, 650))
  mvarCapacity(20) = Array(97, 1085, 3, Array(224, 416, 600, 700))
  mvarCapacity(21) = Array(101, 1156, 4, Array(224, 442, 644, 750))
  mvarCapacity(22) = Array(105, 1258, 4, Array(252, 476, 690, 816))
  mvarCapacity(23) = Array(109, 1364, 4, Array(270, 504, 750, 900))
  mvarCapacity(24) = Array(113, 1474, 4, Array(300, 560, 810, 960))
  mvarCapacity(25) = Array(117, 1588, 4, Array(312, 588, 870, 1050))
  mvarCapacity(26) = Array(121, 1706, 4, Array(336, 644, 952, 1110))
  mvarCapacity(27) = Array(125, 1828, 4, Array(360, 700, 1020, 1200))
  mvarCapacity(28) = Array(129, 1921, 3, Array(390, 728, 1050, 1260))
  mvarCapacity(29) = Array(133, 2051, 3, Array(420, 784, 1140, 1350))
  mvarCapacity(30) = Array(137, 2185, 3, Array(450, 812, 1200, 1440))
  mvarCapacity(31) = Array(141, 2323, 3, Array(480, 868, 1290, 1530))
  mvarCapacity(32) = Array(145, 2465, 3, Array(510, 924, 1350, 1620))
  mvarCapacity(33) = Array(149, 2611, 3, Array(540, 980, 1440, 1710))
  mvarCapacity(34) = Array(153, 2761, 3, Array(570, 1036, 1530, 1800))
  mvarCapacity(35) = Array(157, 2876, 0, Array(570, 1064, 1590, 1890))
  mvarCapacity(36) = Array(161, 3034, 0, Array(600, 1120, 1680, 1980))
  mvarCapacity(37) = Array(165, 3196, 0, Array(630, 1204, 1770, 2100))
  mvarCapacity(38) = Array(169, 3362, 0, Array(660, 1260, 1860, 2220))
  mvarCapacity(39) = Array(173, 3532, 0, Array(720, 1316, 1950, 2310))
  mvarCapacity(40) = Array(177, 3706, 0, Array(750, 1372, 2040, 2430))


  mvarMarker(0) = Array(&HA1, &HA1, &HA1, &HA1, &HA1)
  mvarMarker(1) = Array(&HA1, &HA0, &HA0, &HA0, &HA1)
  mvarMarker(2) = Array(&HA1, &HA0, &HA1, &HA0, &HA1)
  mvarMarker(3) = Array(&HA1, &HA0, &HA0, &HA0, &HA1)
  mvarMarker(4) = Array(&HA1, &HA1, &HA1, &HA1, &HA1)

  mvarFinder(0) = Array(&HC1, &HC1, &HC1, &HC1, &HC1, &HC1, &HC1)
  mvarFinder(1) = Array(&HC1, &HC0, &HC0, &HC0, &HC0, &HC0, &HC1)
  mvarFinder(2) = Array(&HC1, &HC0, &HC1, &HC1, &HC1, &HC0, &HC1)
  mvarFinder(3) = Array(&HC1, &HC0, &HC1, &HC1, &HC1, &HC0, &HC1)
  mvarFinder(4) = Array(&HC1, &HC0, &HC1, &HC1, &HC1, &HC0, &HC1)
  mvarFinder(5) = Array(&HC1, &HC0, &HC0, &HC0, &HC0, &HC0, &HC1)
  mvarFinder(6) = Array(&HC1, &HC1, &HC1, &HC1, &HC1, &HC1, &HC1)

  mvarLevel = qrLevelL
  mvarEncodingMode = qr8BitMode
  mvarCaseSensitive = True
  mvarQuietZone = 0
  mvarShowMarkers = False
  
  mvarFindBestMask = True
   
End Sub

'' Append data to an input object.
Private Sub zzAppendInputItem(ByRef Items() As varInputItems, _
                              ByVal mode As Integer, _
                              ByVal Size As Integer, _
                              ByRef data() As Byte)
                              
  Dim newitem As varInputItems
  If zzNewInputItem(newitem, mode, Size, data) Then
    Call zzPutStackVariant(Items, newitem)
  End If
End Sub

'' Append one bitstream created from number to another
Private Sub zzAppendNum(ByRef BitStream() As Byte, ByVal bits As Integer, ByVal num As Integer)
  Dim b() As Byte
  Dim Temp As Long
  Dim i As Integer
  
  If (bits <> 0) Then
    '' Return new bitstream from number
    ReDim b(bits - 1)
    Temp = zzShiftLeft(1, (bits - 1))
    For i = 0 To bits - 1
      b(i) = IIf((num And Temp) <> 0, 1, 0)
      Temp = zzShiftRight(Temp, 1)
    Next
    
    BitStream = zzArrayMerge(BitStream, b)
  End If
End Sub

'' Append Padding Bit to bitstream
Private Sub zzAppendPaddingBit(ByRef bstream() As Byte)
  Dim bits As Integer
  Dim maxwords As Integer
  Dim maxbits As Integer
  Dim words As Integer
  Dim padding() As Byte
  Dim padlen As Integer
  Dim padbuf() As Byte
  Dim i As Integer
  Dim j As Integer
  Dim p As Integer
  Dim mask As Integer
  Dim b() As Byte

  If (zzUBound(bstream) >= 0) Then
    bits = UBound(bstream) + 1
    maxwords = mvarCapacity(mvarVersion)(qrCapWords) - mvarCapacity(mvarVersion)(qrCapEc)(mvarLevel)
    maxbits = maxwords * 8
    If (maxbits <> bits) Then
      If (maxbits - bits < 5) Then
        Call zzAppendNum(bstream, maxbits - bits, 0)
      Else
        bits = bits + 4
        words = Int((bits + 7) / 8)
        Call zzAppendNum(padding, words * 8 - bits + 4, 0)
        padlen = maxwords - words
        If (padlen > 0) Then
          ReDim padbuf(padlen - 1)
          For i = 0 To padlen - 1
            padbuf(i) = IIf((i And 1) <> 0, &H11, &HEC)
          Next
          '' Append one bitstream created from bytes to another
          If (padlen <> 0) Then
            '' Return new bitstream from bytes
            ReDim b(padlen * 8 - 1)
            p = 0
            For i = 0 To padlen - 1
              mask = &H80
              For j = 0 To 7
                b(p) = IIf((padbuf(i) And mask) <> 0, 1, 0)
                p = p + 1
                mask = zzShiftRight(mask, 1)
              Next
            Next
            padding = zzArrayMerge(padding, b)
          End If
        End If
        bstream = zzArrayMerge(bstream, padding)
      End If
    End If
  End If
End Sub

'' Merge of two array
Private Function zzArrayMerge(ByRef Data1() As Byte, ByRef Data2() As Byte) As Byte()
  Dim Temp() As Byte
  Dim i As Integer
  Dim l1 As Integer
  Dim l2 As Integer
  
  l1 = zzUBound(Data1)
  l2 = zzUBound(Data2)
  
  If (l2 = -1) Then
    Temp = Data1
  Else
    If (l1 = -1) Then
      Temp = Data2
    Else
      ReDim Temp(l1 + l2 + 1)
      For i = 0 To l1
        Temp(i) = Data1(i)
      Next
    
      For i = 0 To l2
        Temp(l1 + 1 + i) = Data2(i)
      Next
    End If
  End If
  
  zzArrayMerge = Temp
End Function

'' Shift down first value of array
Private Function zzArrayShift(ByRef data() As Byte) As Byte
  Dim Temp() As Byte
  Dim i As Integer
  Dim n As Integer
  
  Temp = data
  zzArrayShift = data(0)
  n = UBound(data)
  
  If n > 0 Then
    ReDim Preserve data(n - 1)
    For i = 1 To n
      data(i - 1) = Temp(i)
    Next
  Else
    Erase data
  End If
End Function

'' Slice of an array
Private Function zzArraySlice(ByRef data() As Byte, _
                              ByVal offset As Integer, _
                              Optional ByVal Length As Integer = 0) As Byte()
  Dim Temp() As Byte
  Dim i As Integer
  Dim l As Integer
  
  l = UBound(data) - offset
  If (Length > 0) And (Length < l) Then l = Length
  
  ReDim Temp(l)
  For i = 0 To l
    Temp(i) = data(offset + i)
  Next
  zzArraySlice = Temp
End Function

Private Sub zzArrayUnset(ByRef Stack() As Integer, ByRef Index As Integer)
  
  Dim i As Integer
  Dim k As Integer
  Dim Temp() As Integer
  
  Dim intBound As Integer
  intBound = UBound(Stack) - 1
  
  If intBound <> -1 Then
    ReDim Temp(intBound)
    k = 0
    For i = 0 To intBound
      Temp(i) = Stack(k)
      k = k + 1
      If k = Index Then k = k + 1
    Next
  End If
  Stack = Temp
  
End Sub

Private Function zzBitsModeAn(ByVal Size As Integer) As Integer
  Dim bits As Integer
  
  bits = Int(Size * 5.5)
  If (Size And 1) <> 0 Then bits = bits + 6
  
  zzBitsModeAn = bits
End Function

Private Function zzBitsModeNum(ByVal Size As Integer) As Integer
  Dim w As Integer
  Dim bits As Integer
  
  w = Int(Size / 3)
  bits = (w * 10)
  Select Case Size - (w * 3)
    Case 1:
      bits = bits + 4
    
    Case 2:
      bits = bits + 7
  End Select
  zzBitsModeNum = bits
End Function

'' Convert bitstream to bytes
Private Function zzBitstream2Byte(ByRef bstream() As Byte) As Byte()
  Dim data() As Byte
  Dim bytes As Integer
  Dim p As Integer
  Dim i As Integer
  Dim j As Integer
  Dim v As Integer
  Dim Size As Integer
  
  If (zzUBound(bstream) >= 0) Then
    Size = zzUBound(bstream) + 1
    If (Size > 0) Then
      ReDim data(Int((Size + 7) / 8) - 1)
      bytes = Int(Size / 8)
      p = 0
      For i = 0 To bytes - 1
        v = 0
        For j = 0 To 7
          v = zzShiftLeft(v, 1) Or bstream(p)
          p = p + 1
        Next
        data(i) = v
      Next
      If (Size And 7) <> 0 Then
        v = 0
        For j = 0 To (Size And 7) - 1
          v = zzShiftLeft(v, 1) Or bstream(p)
          p = p + 1
        Next
        data(bytes) = v
      End If
    End If
    zzBitstream2Byte = data
  End If
End Function

Private Function zzCalcN1N3(ByRef RunLength() As Integer, ByVal Length As Integer) As Integer
 
  Dim dmt As Integer
  Dim i As Integer
  Dim fact As Integer
  
  dmt = 0
  For i = 0 To Length - 1
    If (RunLength(i) >= 5) Then
      dmt = dmt + (qrMaskN1 + (RunLength(i) - 5))
    End If
    If (i And 1) <> 0 Then
      If ((i >= 3) And (i < (Length - 2)) And (RunLength(i) Mod 3 = 0)) Then
        fact = Int(RunLength(i) / 3)
        If ((RunLength(i - 2) = fact) _
          And (RunLength(i - 1) = fact) _
          And (RunLength(i + 1) = fact) _
          And (RunLength(i + 2) = fact)) Then
          If ((RunLength(i - 3) < 0) Or (RunLength(i - 3) >= (4 * fact))) Then
            dmt = dmt + qrMaskN3
          ElseIf (((i + 3) >= Length) Or (RunLength(i + 3) >= (4 * fact))) Then
            dmt = dmt + qrMaskN3
          End If
        End If
      End If
    End If
  Next
  zzCalcN1N3 = dmt
End Function

Private Function zzCheckModeAN(ByVal Size As Integer, _
                               ByRef data() As Byte) As Boolean
  Dim i As Integer
  
  For i = 0 To Size - 1
    If (zzLookAlphaNumTable(data(i)) = -1) Then
      zzCheckModeAN = False
      Exit Function
    End If
  Next
  zzCheckModeAN = True
End Function

Private Function zzCheckModeKanji(ByVal Size As Integer, _
                                  ByRef data() As Byte) As Boolean
  Dim i As Integer
  Dim Value As Long
  
  If (Size And 1) <> 0 Then
    zzCheckModeKanji = False
  Else
    For i = 0 To Size - 1 Step 2
      Value = (zzShiftLeft(data(i), 8)) Or data(i + 1)
      If ((Value < &H8140) Or ((Value > &H9FFC) And (Value < &HE040)) Or (Value > &HEBBF)) Then
        zzCheckModeKanji = False
        Exit Function
      End If
    Next
    zzCheckModeKanji = True
  End If
End Function

Private Function zzCheckModeNum(ByVal Size As Integer, _
                               ByRef data() As Byte) As Boolean
  Dim i As Integer
  
  For i = 0 To Size - 1
    If ((data(i) < 48) Or (data(i) > 57)) Then  '' 48 = Asc("0"), 57 = Asc("9")
      zzCheckModeNum = False
      Exit Function
    End If
  Next
  zzCheckModeNum = True
End Function

'' Create initialized frame.
Private Sub zzCreateFrame()
  Dim setPattern() As Variant
  Dim width As Integer
  Dim yOffset As Integer
  Dim Y As Integer
  Dim i As Integer
  Dim v As Long
  Dim X As Integer
  Dim d As Integer
  Dim w As Integer
  Dim j As Integer
  Dim cx As Integer
  Dim cy As Integer
  Dim wo As Integer
  
  width = mvarCapacity(mvarVersion)(qrCapwidth)
  ReDim mvarFrame(width - 1, width - 1)
  
  For Y = 0 To 6
    For i = 0 To 6
      mvarFrame(Y, i) = CByte(mvarFinder(Y)(i))
      mvarFrame(Y, width - 7 + i) = CByte(mvarFinder(Y)(i))
      mvarFrame(width - 7 + Y, i) = CByte(mvarFinder(Y)(i))
    Next
  Next
  
  '' Separator
  yOffset = width - 7
  For Y = 0 To 6
    mvarFrame(Y, 7) = &HC0
    mvarFrame(Y, width - 8) = &HC0
    mvarFrame(yOffset, 7) = &HC0
    yOffset = yOffset + 1
  Next
  
  setPattern = Array(&HC0, &HC0, &HC0, &HC0, &HC0, &HC0, &HC0, &HC0)
  
  For i = 0 To 7
    mvarFrame(7, i) = CByte(setPattern(i))
    mvarFrame(7, width - 8 + i) = CByte(setPattern(i))
    mvarFrame(width - 8, i) = CByte(setPattern(i))
  Next
  
  '' Format info
  setPattern = Array(&H84, &H84, &H84, &H84, &H84, &H84, &H84, &H84, &H84)
  
  For i = 0 To 8
    mvarFrame(8, i) = CByte(setPattern(i))
    If i < 8 Then mvarFrame(8, width - 8 + i) = CByte(setPattern(i))
  Next
  
  yOffset = width - 8
  For Y = 0 To 7
    mvarFrame(Y, 8) = &H84
    mvarFrame(yOffset, 8) = &H84
    yOffset = yOffset + 1
  Next
  
  '' Timing pattern
  For i = 1 To width - 15 - 1
    mvarFrame(6, 7 + i) = &H90 Or (i And 1)
    mvarFrame(7 + i, 6) = &H90 Or (i And 1)
  Next
  
  '' Alignment pattern
  If (mvarVersion >= 2) Then
    d = mvarAlignPattern(mvarVersion)(1) - mvarAlignPattern(mvarVersion)(0)
    If (d < 0) Then
      w = 2
    Else
      w = Int((width - mvarAlignPattern(mvarVersion)(0)) / d + 2)
    End If
    
    If (w * w - 3 = 1) Then
      X = mvarAlignPattern(mvarVersion)(0)
      Y = mvarAlignPattern(mvarVersion)(0)
      '' Put an alignment marker.
      For j = 0 To 4
        For i = 0 To 4
          mvarFrame(Y - 2 + j, X - 2 + i) = CByte(mvarMarker(j)(i))
        Next
      Next
    End If
    
    If (w * w - 3 <> 1) Then
      cx = mvarAlignPattern(mvarVersion)(0)
      wo = w - 1
      For X = 1 To wo - 1
        '' Put an alignment marker.
        For j = 0 To 4
          For i = 0 To 4
            mvarFrame(cx - 2 + j, 4 + i) = CByte(mvarMarker(j)(i))
            mvarFrame(4 + j, cx - 2 + i) = CByte(mvarMarker(j)(i))
          Next
        Next
        cx = cx + d
      Next
      cy = mvarAlignPattern(mvarVersion)(0)
      For Y = 0 To wo - 1
        cx = mvarAlignPattern(mvarVersion)(0)
        For X = 0 To wo - 1
          '' Put an alignment marker.
          For j = 0 To 4
            For i = 0 To 4
              mvarFrame(cy - 2 + j, cx - 2 + i) = CByte(mvarMarker(j)(i))
            Next
          Next
          cx = cx + d
        Next
        cy = cy + d
      Next
    End If
  End If
  
  '' Version information
  If (mvarVersion >= 7) Then
    '' Return BCH encoded version information pattern that is used for the symbol of version 7 or greater. Use lower 18 bits.
    v = 0
    If (mvarVersion <= mvarVersionMax) Then v = mvarVersionPattern((mvarVersion - 7))
    
    For X = 0 To 5
      For Y = 0 To 2
        mvarFrame((width - 11) + Y, X) = &H88 Or (v And 1)
        mvarFrame(X, Y + (width - 11)) = &H88 Or (v And 1)
        v = zzShiftRight(v, 1)
      Next
    Next
  End If
  '' and a little bit...
  mvarFrame(width - 8, 8) = &H81
End Sub

Private Function zzEat8() As Integer
  Dim p As Integer
  Dim q As Integer
  Dim dif As Integer
  Dim mode As Integer
  Dim intLen As Integer
  
  p = 1
  intLen = Len(mvarDataStr)
  Do While (p < intLen)
    mode = zzIdentifyMode(p + 1)
    If (mode = qrModeKJ) Then Exit Do
    
    If (mode = qrModeNM) Then
      q = p
      Do While (zzIsDigit(mvarDataStr, q + 1))
        q = q + 1
      Loop
      dif = Int(p * 8) + zzBitsModeNum(q - p) + 4 + _
            zzLengthIndicator(qrModeNM, mvarVersion) - Int(q * 8)
      If (dif < 0) Then
        Exit Do
      Else
        p = q
      End If
    ElseIf (mode = qrModeAN) Then
      q = p
      Do While (zzIsAlphaNum(mvarDataStr, q + 1))
        q = q + 1
      Loop
      dif = Int(p * 8) + zzBitsModeAn(q - p) + 4 + _
            zzLengthIndicator(qrModeAN, mvarVersion) - Int(q * 8)
      If (dif < 0) Then
        Exit Do
      Else
        p = q
      End If
    Else
      p = p + 1
    End If
  Loop

  Call zzAppendInputItem(mvarItems, qrMode8B, p, StrConv(mvarDataStr, vbFromUnicode))
  zzEat8 = p
End Function

Private Function zzEatAN() As Integer
 
  Dim p As Integer
  Dim q As Integer
  Dim dif As Integer
  Dim run As Integer

  p = 1
  Do While (zzIsAlphaNum(mvarDataStr, p + 1))
    If (zzIsDigit(mvarDataStr, p + 1)) Then
      q = p
      Do While (zzIsDigit(mvarDataStr, q + 1))
        q = q + 1
      Loop
      dif = zzBitsModeAn(p) + zzBitsModeNum(q - p) + 4 + _
            zzLengthIndicator(qrModeNM, mvarVersion) - zzBitsModeAn(q)
      If (dif < 0) Then
        Exit Do
      Else
        p = q
      End If
    Else
      p = p + 1
    End If
    
  Loop
  
  run = p
  If (Not zzIsAlphaNum(mvarDataStr, p + 1)) Then
    dif = zzBitsModeAn(run) + 4 + _
          zzLengthIndicator(qrModeAN, mvarVersion) + 8 - Int((run + 1) * 8)
    If (dif > 0) Then
      zzEatAN = zzEat8()
      Exit Function
    End If
  End If

  Call zzAppendInputItem(mvarItems, qrModeAN, run, StrConv(mvarDataStr, vbFromUnicode))
  zzEatAN = run
End Function

Private Function zzEatKanji() As Integer
  Dim p As Integer
  
  p = 1
  Do While (zzIdentifyMode(p + 1) = qrModeKJ)
    p = p + 2
  Loop
  Call zzAppendInputItem(mvarItems, qrModeKJ, p, StrConv(mvarDataStr, vbFromUnicode))
  zzEatKanji = p
End Function

Private Function zzEatNum() As Integer
  
  Dim p As Integer
  Dim run As Integer
  Dim mode As Integer
  Dim dif As Integer
  
  p = 0
  Do While (zzIsDigit(mvarDataStr, p + 1))
    p = p + 1
  Loop
  run = p
  mode = zzIdentifyMode(p + 1)
  If (mode = qrMode8B) Then
    dif = zzBitsModeNum(run) + 4 + _
          zzLengthIndicator(qrModeNM, mvarVersion) + 8 - Int((run + 1) * 8)
    If (dif > 0) Then
      zzEatNum = zzEat8()
      Exit Function
    End If
  End If
  If (mode = qrModeAN) Then
    dif = zzBitsModeNum(run) + 4 + _
          zzLengthIndicator(qrModeNM, mvarVersion) + zzBitsModeAn(1) - zzBitsModeAn(run + 1)
    If (dif > 0) Then
      zzEatNum = zzEatAN()
      Exit Function
    End If
  End If

  Call zzAppendInputItem(mvarItems, qrModeNM, run, StrConv(mvarDataStr, vbFromUnicode))
  zzEatNum = run
End Function

'' zzEncodeBitStream
Private Sub zzEncodeBitStream(ByRef InputItem As varInputItems)
  Dim words As Long
  Dim st1 As varInputItems
  Dim st2 As varInputItems
  Dim Temp() As Byte
  Dim i As Integer
  Dim tempNum() As Byte
  Dim Value As Integer
  Dim asc0 As Integer
  Dim l As Integer
  
  '' Return the maximum length for the mode and version.
  If (mvarVersion <= 9) Then
    l = 0
  ElseIf (mvarVersion <= 26) Then
    l = 1
  Else
    l = 2
  End If
  words = zzShiftLeft(1, mvarLengthTableBits(InputItem.mode)(l)) - 1
  If (InputItem.mode = qrModeKJ) Then words = words * 2  '' the number of bytes is required

  If (InputItem.Size > words) Then
    Temp = InputItem.data
    Call zzNewInputItem(st1, InputItem.mode, words, Temp)
    Call zzNewInputItem(st2, InputItem.mode, InputItem.Size - words, zzArraySlice(Temp, words))
    Call zzEncodeBitStream(st1)
    Call zzEncodeBitStream(st2)
    InputItem.bstream = zzArrayMerge(st1.bstream, st2.bstream)
  Else
    Select Case InputItem.mode
      Case qrModeNM:
        
        asc0 = Asc("0")
        words = Int(InputItem.Size / 3)
        Value = &H1
        Erase InputItem.bstream
        Call zzAppendNum(InputItem.bstream, 4, Value)
        Call zzAppendNum(InputItem.bstream, zzLengthIndicator(qrModeNM, mvarVersion), InputItem.Size)
        For i = 0 To words - 1
          Value = (InputItem.data(i * 3) - asc0) * 100
          Value = Value + (InputItem.data(i * 3 + 1) - asc0) * 10
          Value = Value + (InputItem.data(i * 3 + 2) - asc0)
          Call zzAppendNum(InputItem.bstream, 10, Value)
        Next
        If (InputItem.Size - words * 3 = 1) Then
          Value = InputItem.data(words * 3) - asc0
          Call zzAppendNum(InputItem.bstream, 4, Value)
        ElseIf ((InputItem.Size - (words * 3)) = 2) Then
          Value = (InputItem.data(words * 3) - asc0) * 10
          Value = Value + (InputItem.data(words * 3 + 1) - asc0)
          Call zzAppendNum(InputItem.bstream, 7, Value)
        End If
      
      Case qrModeAN:
        words = Int(InputItem.Size / 2)
        Erase InputItem.bstream
        Call zzAppendNum(InputItem.bstream, 4, &H2)
        Call zzAppendNum(InputItem.bstream, zzLengthIndicator(qrModeAN, mvarVersion), InputItem.Size)
        For i = 0 To words - 1
          Value = Int(zzLookAlphaNumTable(InputItem.data(i * 2)) * 45)
          Value = Value + Int(zzLookAlphaNumTable(InputItem.data((i * 2) + 1)))
          Call zzAppendNum(InputItem.bstream, 11, Value)
        Next
        If (InputItem.Size And 1) Then
          Value = zzLookAlphaNumTable(InputItem.data((words * 2)))
          Call zzAppendNum(InputItem.bstream, 6, Value)
        End If
      
      Case qrMode8B:
        Call zzAppendNum(tempNum, 4, &H4)
        Call zzAppendNum(tempNum, zzLengthIndicator(qrMode8B, mvarVersion), InputItem.Size)
        For i = 0 To InputItem.Size - 1
          Call zzAppendNum(tempNum, 8, InputItem.data(i))
        Next
        InputItem.bstream = tempNum
      
      Case qrModeKJ:
        Erase InputItem.bstream
        Call zzAppendNum(InputItem.bstream, 4, &H8)
        Call zzAppendNum(InputItem.bstream, zzLengthIndicator(qrModeKJ, mvarVersion), Int(InputItem.Size / 2))
        For i = 0 To InputItem.Size - 1 Step 2
          Value = zzShiftLeft(InputItem.data(i), 8) Or InputItem.data(i + 1)
          If (Value <= &H9FFC) Then
            Value = Value - &H8140
          Else
            Value = Value - &HC140
          End If
          Value = (Value And &HFF) + (zzShiftRight(Value, 8)) * &HC0
          Call zzAppendNum(InputItem.bstream, 13, Value)
        Next
      
      Case Else:
    
    End Select
  End If
End Sub

'' Encode mask
Private Function zzEncodeData() As Boolean
  Dim Spec(4) As Integer
  Dim ret As Integer
  Dim i As Integer
  Dim j As Integer
  Dim code As Integer
  Dim TempBit As Integer
  
  Dim posX As Integer
  Dim posY As Integer
  
  Dim b2 As Integer
  Dim Ecc As Integer
  Dim data As Integer
  Dim masked() As Byte
  Dim Blocks As Integer
  Dim EccCode() As Byte  ' Error correction code.
  Dim width As Integer
  Dim Bit As Integer
  Dim Direction As Integer
  Dim BitX As Integer '' X position of bit.
  Dim BitY As Integer '' Y position of bit.
  Dim EccLength As Integer '' Error correction length.
  Dim Count As Integer
  Dim DataLength As Integer
  Dim B1 As Integer
  
  '' Calc array of ECC specification.
  B1 = mvarEccTable(mvarVersion)(mvarLevel)(0)
  b2 = mvarEccTable(mvarVersion)(mvarLevel)(1)
  data = mvarCapacity(mvarVersion)(qrCapWords) - mvarCapacity(mvarVersion)(qrCapEc)(mvarLevel)
  Ecc = mvarCapacity(mvarVersion)(qrCapEc)(mvarLevel)
  If (b2 = 0) Then
    Spec(0) = B1
    Spec(1) = Int(data / B1)
    Spec(2) = Int(Ecc / B1)
    Spec(3) = 0
    Spec(4) = 0
  Else
    Spec(0) = B1
    Spec(1) = Int(data / (B1 + b2))
    Spec(2) = Int(Ecc / (B1 + b2))
    Spec(3) = b2
    Spec(4) = Spec(1) + 1
  End If
  
  DataLength = (Spec(0) * Spec(1)) + (Spec(3) * Spec(4))
  EccLength = (Spec(0) + Spec(3)) * Spec(2)
  ReDim EccCode(EccLength - 1)
  Blocks = Spec(0) + Spec(3)
  ret = zzInitCode(Spec, EccCode)
  If (ret < 0) Then Exit Function
  
  Count = 0
  width = mvarCapacity(mvarVersion)(qrCapwidth) 'getWidth(mvarVersion)
  
  If ((mvarVersion < 1) Or (mvarVersion > mvarVersionMax)) Then Exit Function
  Call zzCreateFrame
  
  BitX = width - 1
  BitY = width - 1
  Direction = -1
  Bit = -1

  '' inteleaved data and ecc codes
  For i = 0 To (DataLength + EccLength) - 1
    code = zzGetCode(Count, Blocks, EccLength, DataLength, B1)
    TempBit = &H80
    For j = 0 To 7
      If zzNextPosition(posX, posY, Bit, BitX, BitY, Direction, width) Then
        mvarFrame(posY, posX) = CInt(&H2 Or IIf((TempBit And code) <> 0, 1, 0)) '' Set frame value at specified position
      End If
      TempBit = zzShiftRight(TempBit, 1)
    Next
  Next
  
  '' remainder bits
  j = mvarCapacity(mvarVersion)(qrCapReminder)
  For i = 0 To j - 1
    If zzNextPosition(posX, posY, Bit, BitX, BitY, Direction, width) Then
      mvarFrame(posY, posX) = &H2 '' Set frame value at specified position
    End If
  Next
  
  '' masking
  If (mvarFindBestMask) Then
    masked = zzMask(width, mvarFrame)
  Else
    masked = zzMakeMask(width, mvarFrame, (Int(qrDefaultMask) Mod 8))
  End If
  If zzUBound(masked) = -1 Then Exit Function
  mvarData = masked
  zzEncodeData = True
End Function

'' Encode a Reed-Solomon codec and returns the parity array
Private Sub zzEncodeRsChar(ByRef Rs As varRSItems, ByRef data() As Byte, ByRef Parity() As Byte)
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  Dim feedback As Integer
  
  ReDim Parity(Rs.NRoots - 1)
  For i = 0 To (Rs.nn - Rs.NRoots - Rs.Pad) - 1
    feedback = Rs.Index_of(data(i) Xor Parity(0))
    If (feedback <> Rs.nn) Then
      ''  feedback term is non-zero
      '' This line is unnecessary when GENPOLY(NROOTS) is unity, as it must
      '' always be for the polynomials constructed by zzInitRs()
      feedback = zzModNN(Rs, Rs.nn - Rs.genpoly(Rs.NRoots) + feedback)
      For j = 1 To Rs.NRoots - 1
        Parity(j) = Parity(j) Xor Rs.Alpha_to(zzModNN(Rs, feedback + Rs.genpoly((Rs.NRoots - j))))
      Next
    End If
    Call zzArrayShift(Parity)
    k = UBound(Parity)
    ReDim Preserve Parity(k + 1)
    If (feedback <> Rs.nn) Then
      Parity(k + 1) = Rs.Alpha_to(zzModNN(Rs, feedback + Rs.genpoly(0)))
    Else
      Parity(k + 1) = 0
    End If
  Next
  
End Sub

'' zzEstimateBitStreamSize
Private Function zzEstimateBitStreamSize(ByRef Items() As varInputItems, ByVal Version As Integer) As Integer
  Dim bits As Integer
  Dim num As Integer
  Dim i As Integer
  Dim l As Integer
  Dim m As Long
  
  bits = 0
  If (Version = 0) Then Version = 1
  
  num = UBound(Items)
  For i = 0 To num
    Select Case Items(i).mode
      Case qrModeNM:
        bits = zzBitsModeNum(Items(i).Size)
      
      Case qrModeAN:
        bits = zzBitsModeAn(Items(i).Size)
      
      Case qrMode8B:
        bits = Int(Items(i).Size * 8)
      
      Case qrModeKJ:
        bits = Int(Items(i).Size * 6.5)
      
      Case Else:
        zzEstimateBitStreamSize = 0
        Exit Function
    
    End Select
    l = zzLengthIndicator(Items(i).mode, Version)
    m = zzShiftLeft(1, l)
    num = Int((Items(i).Size + m - 1) / m)
    bits = bits + num * (4 + l)
  Next
  zzEstimateBitStreamSize = bits
End Function

'' zzEstimateVersion
Private Function zzEstimateVersion(ByRef Items() As varInputItems) As Integer
  Dim Version As Integer
  Dim prev As Integer
  Dim bits As Integer
  
  Version = 0
  prev = 0
  Do
    prev = Version
    bits = zzEstimateBitStreamSize(Items, prev)
    Version = zzMinimumVersion(Int((bits + 7) / 8))
    If (Version < 0) Then
      zzEstimateVersion = -1
      Exit Function
    End If
  Loop While (Version > prev)
  zzEstimateVersion = Version
End Function

'' zzEvaluateSymbol
Private Function zzEvaluateSymbol(ByVal width As Integer, ByRef MaskTable() As Byte) As Integer
  
  Dim head As Integer
  Dim dmt As Integer
  Dim Y As Integer
  Dim X As Integer
  Dim b22 As Integer
  Dim w22 As Integer
  Dim RunLength() As Integer
 
  ReDim RunLength(mvarWidthMax)
  head = 0
  dmt = 0
  For Y = 0 To width - 1
    head = 0
    RunLength(0) = 1
    For X = 0 To width - 1
      If ((X > 0) And (Y > 0)) Then
        b22 = MaskTable(Y, X) And MaskTable(Y, X - 1) And MaskTable(Y - 1, X) And MaskTable(Y - 1, X - 1)
        w22 = MaskTable(Y, X) Or MaskTable(Y, X - 1) Or MaskTable(Y - 1, X) Or MaskTable(Y - 1, X - 1)
        If ((b22 Or (w22 Xor 1)) And 1) Then
          dmt = dmt + qrMaskN2
        End If
      End If
      If ((X = 0) And (MaskTable(Y, X) And 1) <> 0) Then
        RunLength(0) = -1
        head = 1
        RunLength(head) = 1
      ElseIf (X > 0) Then
        If ((MaskTable(Y, X) Xor MaskTable(Y, X - 1)) And 1) <> 0 Then
          head = head + 1
          RunLength(head) = 1
        Else
          RunLength(head) = RunLength(head) + 1
        End If
      End If
    Next
    dmt = dmt + zzCalcN1N3(RunLength, head + 1)
  Next
  For X = 0 To width - 1
    head = 0
    RunLength(0) = 1
    For Y = 0 To width - 1
      If ((Y = 0) And (MaskTable(Y, X) And 1) <> 0) Then
        RunLength(0) = -1
        head = 1
        RunLength(head) = 1
      ElseIf (Y > 0) Then
        If ((MaskTable(Y, X) Xor MaskTable(Y - 1, X)) And 1) <> 0 Then
          head = head + 1
          RunLength(head) = 1
        Else
          RunLength(head) = RunLength(head) + 1
        End If
      End If
    Next
    dmt = dmt + zzCalcN1N3(RunLength, head + 1)
  Next
  zzEvaluateSymbol = dmt
End Function

'' Return Reed-Solomon block code.
Private Function zzGetCode(ByRef Count As Integer, _
                           ByVal Blocks As Integer, _
                           ByVal EccLength As Integer, _
                           ByVal DataLength As Integer, _
                           ByVal B1 As Integer) As Integer
  Dim row As Integer
  Dim col As Integer
  Dim ret As Integer
  
  If (Count < DataLength) Then
    row = Count Mod Blocks
    col = Int(Count / Blocks)
    If (col >= mvarRsBlocks(0).DataLength) Then row = row + B1
    ret = mvarRsBlocks(row).data(col)
  ElseIf (Count < DataLength + EccLength) Then
    row = (Count - DataLength) Mod Blocks
    col = Int((Count - DataLength) / Blocks)
    ret = mvarRsBlocks(row).Ecc(col)
  Else
    zzGetCode = 0
    Exit Function
  End If
  Count = Count + 1
  zzGetCode = ret
End Function

'' zzIdentifyMode
Private Function zzIdentifyMode(ByVal Pos As Integer) As Integer
  Dim word As Long
  Dim result As Integer
  
  If (Pos > Len(mvarDataStr)) Then
    result = qrModeNL
  Else
    result = qrMode8B
    
    If (zzIsDigit(mvarDataStr, Pos)) Then
      result = qrModeNM
    Else
      If (zzIsAlphaNum(mvarDataStr, Pos)) Then
        result = qrModeAN
      Else
        If (mvarEncodingMode = qrKanjiMode) Then
          If (Pos + 1 <= Len(mvarDataStr)) Then
            word = (zzShiftLeft(Asc(Mid$(mvarDataStr, Pos, 1)), 8)) Or Asc(Mid$(mvarDataStr, Pos + 1, 1))
            If ((word >= &H8140 And word <= &H9FFC) Or (word >= &HE040 And word <= &HEBBF)) Then
              result = qrModeKJ
            End If
          End If
        End If
      End If
    End If
  End If
  
  zzIdentifyMode = result
End Function

'' Initialize code.
Private Function zzInitCode(ByRef Spec() As Integer, ByRef EccCode() As Byte) As Integer

  Dim blockNo As Integer
  Dim dataPos As Integer
  Dim eccpos As Integer
  Dim i As Integer
  Dim Ecc() As Byte
  Dim Rs As varRSItems
  
  Call zzInitRs(Rs, 8, &H11D, 0, 1, Spec(2), 255 - Spec(1) - Spec(2))
  
  blockNo = 0
  dataPos = 0
  eccpos = 0
  ReDim mvarRsBlocks(Spec(0) - 1)
  
  For i = 0 To Spec(0) - 1
    Ecc = zzArraySlice(EccCode, eccpos)
    mvarRsBlocks(blockNo).DataLength = Spec(1)
    mvarRsBlocks(blockNo).data = zzArraySlice(mvarDataCode, dataPos)
    mvarRsBlocks(blockNo).EccLength = Spec(2)
    Call zzEncodeRsChar(Rs, mvarRsBlocks(blockNo).data, Ecc)
    mvarRsBlocks(blockNo).Ecc = Ecc
    
    EccCode = zzArrayMerge(zzArraySlice(EccCode, 0, eccpos), Ecc)
    dataPos = dataPos + Spec(1)
    eccpos = eccpos + Spec(2)
    blockNo = blockNo + 1
  Next
  
  If (Spec(3) = 0) Then
    zzInitCode = 0
  ElseIf Not zzInitRs(Rs, 8, &H11D, 0, 1, Spec(2), 255 - Spec(4) - Spec(2)) Then
    zzInitCode = -1
  Else
    ReDim Preserve mvarRsBlocks(UBound(mvarRsBlocks) + Spec(3))
    For i = 0 To Spec(3) - 1
      Ecc = zzArraySlice(EccCode, eccpos)
      
      mvarRsBlocks(blockNo).DataLength = Spec(4)
      mvarRsBlocks(blockNo).data = zzArraySlice(mvarDataCode, dataPos)
      mvarRsBlocks(blockNo).EccLength = Spec(2)
      Call zzEncodeRsChar(Rs, mvarRsBlocks(blockNo).data, Ecc)
      mvarRsBlocks(blockNo).Ecc = Ecc
      
      EccCode = zzArrayMerge(zzArraySlice(EccCode, 0, eccpos), Ecc)
      dataPos = dataPos + Spec(4)
      eccpos = eccpos + Spec(2)
      blockNo = blockNo + 1
    Next
    zzInitCode = 0
  End If
End Function

'' Initialize a Reed-Solomon codec and add it to existing rsitems
Private Function zzInitRs(ByRef Rs As varRSItems, _
                          ByVal SymSize As Integer, _
                          ByVal GFpoly As Integer, _
                          ByVal Fcr As Integer, _
                          ByVal Prim As Integer, _
                          ByVal NRoots As Integer, _
                          ByVal Pad As Integer) As Boolean
                      
  Dim l As Integer
  Dim i As Integer

  zzInitRs = False
  l = zzUBoundRS(mvarRSItems)
  If l <> -1 Then
    For i = 0 To l
      If ((mvarRSItems(i).Pad = Pad) And _
          (mvarRSItems(i).NRoots = NRoots) And _
          (mvarRSItems(i).mm = SymSize) And _
          (mvarRSItems(i).GFpoly = GFpoly) And _
          (mvarRSItems(i).Fcr = Fcr) And _
          (mvarRSItems(i).Prim = Prim)) Then
        Rs = mvarRSItems(i)
        zzInitRs = True
        Exit Function
      End If
    Next
  End If
  ReDim Preserve mvarRSItems(l + 1)
  mvarRSItems(l + 1) = zzInitRsChar(SymSize, GFpoly, Fcr, Prim, NRoots, Pad)
  Rs = mvarRSItems(l + 1)
  zzInitRs = True
End Function

'' Initialize a Reed-Solomon codec and returns an array of values.
Private Function zzInitRsChar(ByVal SymSize As Integer, ByVal GFpoly As Integer, ByVal Fcr As Integer, ByVal Prim As Integer, ByVal NRoots As Integer, ByVal Pad As Integer) As varRSItems
  '' Based on Reed solomon encoder by Phil Karn, KA9Q (GNU-LGPLv2)
  Dim Rs As varRSItems
  Dim sr As Integer
  Dim i As Integer
  Dim iprim As Integer
  Dim root As Integer
  Dim j As Integer
  
  '' Check parameter ranges
  If ((SymSize < 0) Or (SymSize > 8)) Then
    zzInitRsChar = Rs
    Exit Function
  ElseIf ((Fcr < 0) Or (Fcr >= (zzShiftLeft(1, SymSize)))) Then
    zzInitRsChar = Rs
    Exit Function
  ElseIf ((Prim <= 0) Or (Prim >= (zzShiftLeft(1, SymSize)))) Then
    zzInitRsChar = Rs
    Exit Function
  ElseIf ((NRoots < 0) Or (NRoots >= (zzShiftLeft(1, SymSize)))) Then
    zzInitRsChar = Rs
    Exit Function
  ElseIf ((Pad < 0) Or (Pad >= ((zzShiftLeft(1, SymSize)) - 1 - NRoots))) Then
    zzInitRsChar = Rs
    Exit Function
  End If

  Rs.mm = SymSize
  Rs.nn = (zzShiftLeft(1, SymSize)) - 1
  Rs.Pad = Pad
  ReDim Rs.Alpha_to(Rs.nn)
  ReDim Rs.Index_of(Rs.nn)
  '' Generate Galois field lookup tables
  Rs.Index_of(0) = Rs.nn
  Rs.Alpha_to(Rs.nn) = 0
  sr = 1
  For i = 0 To Rs.nn - 1
    Rs.Index_of(sr) = i
    Rs.Alpha_to(i) = sr
    sr = zzShiftLeft(sr, 1)
    If (sr And (zzShiftLeft(1, SymSize))) <> 0 Then
      sr = sr Xor GFpoly
    End If
    sr = sr And Rs.nn
  Next
  If (sr <> 1) Then '' field generator polynomial is not primitive!
    Exit Function
  End If
  '' Form RS code generator polynomial from its roots
  ReDim Rs.genpoly(NRoots)
  Rs.Fcr = Fcr
  Rs.Prim = Prim
  Rs.NRoots = NRoots
  Rs.GFpoly = GFpoly
  '' Find prim-th root of 1, used in decoding
  iprim = 1
  Do While (iprim Mod Prim) <> 0
    iprim = iprim + Rs.nn
  Loop

  Rs.iprim = Int(iprim / Prim)
  Rs.genpoly(0) = 1
  root = Fcr * Prim
  For i = 0 To NRoots - 1
    Rs.genpoly(i + 1) = 1 '' Multiply genpoly() by  @**(root + x)
    For j = i To 1 Step -1
      If (Rs.genpoly(j) <> 0) Then
        Rs.genpoly(j) = Rs.genpoly(j - 1) Xor Rs.Alpha_to(zzModNN(Rs, Rs.Index_of(Rs.genpoly(j)) + root))
      Else
        Rs.genpoly(j) = Rs.genpoly(j - 1)
      End If
    Next
    '' genpoly(0) can never be zero
    Rs.genpoly(0) = Rs.Alpha_to(zzModNN(Rs, Rs.Index_of(Rs.genpoly(0)) + root))
    root = root + Prim
  Next
  
  '' convert genpoly() to index form for quicker encoding
  For i = 0 To NRoots
    Rs.genpoly(i) = Rs.Index_of(Rs.genpoly(i))
  Next
  zzInitRsChar = Rs
End Function

'' Return true if the character at specified position is an alphanumeric character
Private Function zzIsAlphaNum(ByRef Value As String, ByVal Pos As Integer) As Boolean
  
  If (Pos > Len(Value)) Then
    zzIsAlphaNum = False
  Else
    zzIsAlphaNum = (zzLookAlphaNumTable(Asc(Mid$(Value, Pos, 1))) >= 0)
  End If
End Function

'' Return true if the character at specified position is a number
Private Function zzIsDigit(ByRef Value As String, ByVal Pos As Integer) As Boolean
  Dim c As String
  
  If (Pos > Len(Value)) Then
    zzIsDigit = False
  Else
    c = Mid$(Value, Pos, 1)
    zzIsDigit = (c >= "0") And (c <= "9")
  End If
End Function

'' Return the size of length indicator for the mode and version.
Private Function zzLengthIndicator(ByVal mode As Integer, _
                                 ByVal Version As Integer) As Integer
  Dim l As Integer
  
  If (Version <= 9) Then
    l = 0
  ElseIf (Version <= 26) Then
    l = 1
  Else
    l = 2
  End If
  zzLengthIndicator = mvarLengthTableBits(mode)(l)
End Function

'' 2013.2c

'' Look up the alphabet-numeric convesion table (see JIS X0510:2004, pp.19).
Private Function zzLookAlphaNumTable(ByVal c As Integer) As Integer
  If c > 127 Then
    zzLookAlphaNumTable = -1
  Else
    zzLookAlphaNumTable = mvarAlTable(c)
  End If
End Function

'' zzMakeMask
Private Function zzMakeMask(ByVal width As Integer, _
                            ByRef Frame() As Byte, _
                            ByVal MaskNo As Integer) As Byte()
 
  Dim masked() As Byte
  
  ReDim masked(width - 1, width - 1)
  Call zzMakeMaskNo(MaskNo, width, Frame, masked)
  Call zzWriteFormatInfo(width, masked, MaskNo)
  zzMakeMask = masked
End Function

'' zzMakeMask
Private Function zzMakeMaskNo(ByVal MaskNo As Integer, _
                             ByVal width As Integer, _
                             ByRef s() As Byte, _
                             ByRef d() As Byte) As Integer
  Dim b As Integer
  Dim Y As Integer
  Dim X As Integer
  Dim bitMask() As Integer
  Dim maskFunc As Integer
  
  b = 0
  ReDim bitMask(width - 1, width - 1)

  For Y = 0 To width - 1
    For X = 0 To width - 1
      If (s(Y, X) And &H80) <> 0 Then
        bitMask(Y, X) = 0
      Else
        Select Case MaskNo
          Case 0
            maskFunc = (X + Y) And 1
          Case 1
            maskFunc = (Y And 1)
          Case 2
            maskFunc = (X Mod 3)
          Case 3
            maskFunc = (X + Y) Mod 3
          Case 4
            maskFunc = ((Int(Y / 2)) + (Int(X / 3))) And 1
          Case 5
            maskFunc = ((X * Y) And 1) + ((X * Y) Mod 3)
          Case 6
            maskFunc = (((X * Y) And 1) + ((X * Y) Mod 3)) And 1
          Case 7
            maskFunc = ((((X * Y) Mod 3)) + ((X + Y) And 1)) And 1
        End Select
        bitMask(Y, X) = IIf(maskFunc = 0, 1, 0)
      End If
    Next
  Next
  
  d = s
  For Y = 0 To width - 1
    For X = 0 To width - 1
      If (bitMask(Y, X) = 1) Then d(Y, X) = s(Y, X) Xor 1
      b = b + Int(d(Y, X) And 1)
    Next
  Next
  zzMakeMaskNo = b
End Function

'' zzMask
Private Function zzMask(ByVal width As Integer, ByRef Frame() As Byte) As Byte()
  
  Dim dmt As Integer
  Dim blacks As Integer
  Dim mindmt As Long
  Dim bestMaskNum As Integer
  Dim k As Integer
  Dim i As Integer
  Dim bestMask() As Byte
  Dim checkedmasks() As Integer
  Dim MaskTable() As Byte
  Dim howManuOut As Integer
  Dim remPos As Integer
  
  mindmt = 2147483647 'INT_MAX
  bestMaskNum = 0
  
  ReDim checkedmasks(7)
  For i = 0 To 7
    checkedmasks(i) = i
  Next
  If (qrFindFromRandom <> 0) Then
    howManuOut = 8 - (qrFindFromRandom Mod 9)
    Randomize
    For i = 0 To howManuOut - 1
      remPos = (UBound(checkedmasks) + 1) * Rnd()
      Call zzArrayUnset(checkedmasks, remPos)
    Next
  End If
  bestMask = Frame
  For k = 0 To UBound(checkedmasks)
    i = checkedmasks(k)
    
    ReDim MaskTable(width - 1, width - 1)
  
    blacks = zzMakeMaskNo(i, width, Frame, MaskTable)
    blacks = blacks + zzWriteFormatInfo(width, MaskTable, i)
    blacks = 100 * Int(blacks / (width * width))
    dmt = Int(Int(Abs(blacks - 50) / 5) * qrMaskN4) + zzEvaluateSymbol(width, MaskTable)
    If (dmt < mindmt) Then
      mindmt = dmt
      bestMask = MaskTable
      bestMaskNum = i
    End If
  Next
  zzMask = bestMask
End Function

'' zzMergeBitStream
Private Function zzMergeBitStream(ByRef Items() As varInputItems) As Byte()
  Dim l As Integer
  Dim i As Integer
  Dim Temp() As Byte
  Dim bstream() As Byte
  
  Dim ver As Integer
  Dim bits As Integer
  Dim bFlag As Boolean
  
  bFlag = True
  ver = zzEstimateVersion(Items)
  If (ver > mvarVersion) Then mvarVersion = ver
  
  Do While (bFlag)
    bits = 0
    For i = 0 To UBound(Items)
      Call zzEncodeBitStream(Items(i))
      bits = bits + UBound(Items(i).bstream) + 1
    Next
 
    If (bits < 0) Then
      bFlag = False
    Else
      ver = zzMinimumVersion(Int((bits + 7) / 8))
      If (ver < 0) Then
        bFlag = False
      ElseIf (ver > mvarVersion) Then
         mvarVersion = ver
      Else
        bFlag = False
      End If
    End If
  Loop
  
  l = zzUBoundVar(Items)
  If (l = -1) Then Exit Function
  
  For i = 0 To l
    Temp = Items(i).bstream
    bstream = zzArrayMerge(bstream, Temp)
  Next
  zzMergeBitStream = bstream
End Function

'' Return a version number that satisfies the input code length.
Private Function zzMinimumVersion(ByVal Size As Integer) As Integer
  Dim words As Integer
  Dim i As Integer
  
  For i = 1 To mvarVersionMax
    words = mvarCapacity(i)(qrCapWords) - mvarCapacity(i)(qrCapEc)(mvarLevel)
    If (words >= Size) Then
      zzMinimumVersion = i
      Exit Function
    End If
  Next
  zzMinimumVersion = -1
End Function

'' zzModNN
Private Function zzModNN(ByRef Rs As varRSItems, ByVal X As Integer) As Integer
  Do While (X >= Rs.nn)
    X = X - Rs.nn
    X = (zzShiftRight(X, Rs.mm)) + (X And Rs.nn)
  Loop
  zzModNN = X
End Function

'' zzNewInputItem
Private Function zzNewInputItem(ByRef Item As varInputItems, _
                                ByVal mode As Integer, _
                                ByVal Size As Integer, _
                                ByRef data() As Byte) As Boolean
                               
  Dim bFlag As Boolean
  Dim SetData() As Byte
                               
  SetData = data
  ReDim Preserve SetData(Size - 1)
  
  '' Validate the input data.
  If (Size <= 0) Then
    bFlag = False
  Else
    Select Case mode
      Case qrModeNM:
        bFlag = zzCheckModeNum(Size, SetData)
      
      Case qrModeAN:
        bFlag = zzCheckModeAN(Size, SetData)
      
      Case qrModeKJ:
        bFlag = zzCheckModeKanji(Size, SetData)
      
      Case qrMode8B:
        bFlag = True
      
      Case Else:
        bFlag = False
    End Select
  End If
  
  If bFlag Then
    Item.mode = mode
    Item.Size = Size
    Item.data = SetData
  End If
  zzNewInputItem = bFlag
End Function

'' Return the next frame position
Private Function zzNextPosition(ByRef X As Integer, _
                                ByRef Y As Integer, _
                                ByRef Bit As Integer, _
                                ByRef BitX As Integer, _
                                ByRef BitY As Integer, _
                                ByRef Direction As Integer, _
                                ByVal width As Integer _
                                ) As Boolean
  Dim w As Integer
  
  Do
    If (Bit = -1) Then
      Bit = 0
      X = BitX
      Y = BitY
      zzNextPosition = True
      Exit Function
    End If
    X = BitX
    Y = BitY
    w = width
    If (Bit = 0) Then
      X = X - 1
      Bit = Bit + 1
    Else
      X = X + 1
      Y = Y + Direction
      Bit = Bit - 1
    End If
    If (Direction < 0) Then
      If (Y < 0) Then
        Y = 0
        X = X - 2
        Direction = 1
        If (X = 6) Then
          X = X - 1
          Y = 9
        End If
      End If
    Else
      If (Y = w) Then
        Y = w - 1
        X = X - 2
        Direction = -1
        If (X = 6) Then
          X = X - 1
          Y = Y - 8
        End If
      End If
    End If
    If ((X < 0) Or (Y < 0)) Then
      zzNextPosition = False
      Exit Function
    End If
    BitX = X
    BitY = Y
  Loop While ((mvarFrame(Y, X) And &H80) <> 0)
  zzNextPosition = True
End Function

Private Function zzPutStackVariant(ByRef Stack() As varInputItems, ByRef Value As varInputItems) As Integer
  
  Dim intBound As Integer
  
  intBound = 0
  On Error Resume Next
  intBound = UBound(Stack) + 1
  
  ReDim Preserve Stack(intBound)
  Stack(intBound) = Value
  
  zzPutStackVariant = intBound
End Function

''' Replace a value on the array at the specified position
'Private Sub zzQRSet(ByVal X As Integer, _
'                    ByVal Y As Integer, _
'                    ByVal row As Integer, _
'                    ByRef Repl() As Variant, _
'                    Optional Length As Integer = 0)
'
'  Dim l As Integer
'  Dim i As Integer
'
'  l = UBound(Repl(row))
'  If Length <> 0 Then l = Length - 1
'  For i = 0 To l
'    mvarFrame(Y, X + i) = CByte(Repl(row)(i))
'  Next
'
'End Sub

'' SHIFT binario verso sinistra
Private Static Function zzShiftLeft(ByVal Value As Long, ByVal ShiftCount As Long) As Long
' by Jost Schwider, jost@schwider.de, 20010928
  Dim Pow2(0 To 31) As Long
  Dim i As Long
  Dim mask As Long
  
  Select Case ShiftCount
  Case 1 To 31
  
    'Ggf. Initialisieren:
    If i = 0 Then
      Pow2(0) = 1
      For i = 1 To 30
        Pow2(i) = 2 * Pow2(i - 1)
      Next i
    End If
    
    'Los gehts:
    mask = Pow2(31 - ShiftCount)
    If Value And mask Then
      zzShiftLeft = (Value And (mask - 1)) * Pow2(ShiftCount) Or &H80000000
    Else
      zzShiftLeft = (Value And (mask - 1)) * Pow2(ShiftCount)
    End If
  
    
  Case 0
  
    zzShiftLeft = Value
  
  End Select
End Function

'' SHIFT binario verso destra
Private Function zzShiftRight(ByVal Value As Long, ByVal ShiftCount As Long) As Long
' by Jost Schwider, jost@schwider.de, 20011010
  Select Case ShiftCount
    Case 0&:  zzShiftRight = Value
    Case 1&:  zzShiftRight = (Value And &HFFFFFFFE) \ &H2&
    Case 2&:  zzShiftRight = (Value And &HFFFFFFFC) \ &H4&
    Case 3&:  zzShiftRight = (Value And &HFFFFFFF8) \ &H8&
    Case 4&:  zzShiftRight = (Value And &HFFFFFFF0) \ &H10&
    Case 5&:  zzShiftRight = (Value And &HFFFFFFE0) \ &H20&
    Case 6&:  zzShiftRight = (Value And &HFFFFFFC0) \ &H40&
    Case 7&:  zzShiftRight = (Value And &HFFFFFF80) \ &H80&
    Case 8&:  zzShiftRight = (Value And &HFFFFFF00) \ &H100&
    Case 9&:  zzShiftRight = (Value And &HFFFFFE00) \ &H200&
    Case 10&: zzShiftRight = (Value And &HFFFFFC00) \ &H400&
    Case 11&: zzShiftRight = (Value And &HFFFFF800) \ &H800&
    Case 12&: zzShiftRight = (Value And &HFFFFF000) \ &H1000&
    Case 13&: zzShiftRight = (Value And &HFFFFE000) \ &H2000&
    Case 14&: zzShiftRight = (Value And &HFFFFC000) \ &H4000&
    Case 15&: zzShiftRight = (Value And &HFFFF8000) \ &H8000&
    Case 16&: zzShiftRight = (Value And &HFFFF0000) \ &H10000
  End Select
End Function

'' zzStr2DataCode
Private Function zzStr2DataCode() As Boolean
  Dim mode As Integer
  Dim Length As Integer
  Dim bstream() As Byte
  
  Do While (Len(mvarDataStr) > 0)
    If (mvarDataStr = "") Then Exit Do
    
    mode = zzIdentifyMode(1)
    Select Case mode
      Case qrModeNM:
        Length = zzEatNum()
      Case qrModeAN:
        Length = zzEatAN()
      Case qrModeKJ:
        If (mvarEncodingMode = qrKanjiMode) Then
          Length = zzEatKanji()
        Else
          Length = zzEat8()
        End If
      Case Else:
        Length = zzEat8()
    End Select
    If (Length = 0) Then Exit Do
    
    If (Length < 0) Then
      zzStr2DataCode = False
      Exit Function
    End If
    If Len(mvarDataStr) > Length Then
      mvarDataStr = Mid$(mvarDataStr, Length + 1)
    Else
      mvarDataStr = ""
    End If
  Loop
  
  bstream = zzMergeBitStream(mvarItems)
  Call zzAppendPaddingBit(bstream)
  mvarDataCode = zzBitstream2Byte(bstream)
  If zzUBound(mvarDataCode) = -1 Then
    zzStr2DataCode = False
    Exit Function
  End If

  zzStr2DataCode = True
End Function

'' toUpper
Private Sub zzToUpper()
  Dim intLen As Integer
  Dim p As Integer
  Dim mode As Integer
  
  intLen = Len(mvarDataStr)
  p = 1
  Do While (p <= intLen)
    mode = zzIdentifyMode(Mid$(mvarDataStr, p, 1))
    If (mode = qrModeKJ) Then
      p = p + 2
    Else
      If ((Mid$(mvarDataStr, p, 1) >= "a") And (Mid$(mvarDataStr, p, 1) <= "z")) Then
        Mid$(mvarDataStr, p, 1) = UCase(Mid$(mvarDataStr, p, 1))
      End If
      p = p + 1
    End If
  Loop
End Sub

Private Function zzUBound(ByRef Var() As Byte) As Integer
  Dim intBound As Integer
  
  intBound = -1
  On Error Resume Next
  intBound = UBound(Var)
  zzUBound = intBound
End Function

Private Function zzUBoundRS(ByRef Var() As varRSItems) As Integer
  Dim intBound As Integer
  
  intBound = -1
  On Error Resume Next
  intBound = UBound(Var)
  zzUBoundRS = intBound
End Function

Private Function zzUBoundVar(ByRef Var() As varInputItems) As Integer
  Dim intBound As Integer
  
  intBound = -1
  On Error Resume Next
  intBound = UBound(Var)
  zzUBoundVar = intBound
End Function

'' Write Format Information on frame and returns the number of black bits
Private Function zzWriteFormatInfo(ByVal width As Integer, _
                                   ByRef Frame() As Byte, _
                                   ByVal mask As Integer) As Integer
                                   
  Dim blacks As Integer
  Dim Format As Integer
  Dim i As Integer
  Dim v As Integer
  
  '' Calc BCH encoded format information pattern.
  If ((mask < 0) Or (mask > 7)) Or ((mvarLevel < 0) Or (mvarLevel > 3)) Then
    Format = 0
  Else
    Format = mvarFormatInfo(mvarLevel)(mask)
  End If
  
  blacks = 0
  For i = 0 To 7
    If (Format And 1) Then
      blacks = blacks + 2
      v = &H85
    Else
      v = &H84
    End If
    Frame(8, width - 1 - i) = v
    If (i < 6) Then
      Frame(i, 8) = v
    Else
      Frame(i + 1, 8) = v
    End If
    Format = zzShiftRight(Format, 1)
  Next
  For i = 0 To 6
    If (Format And 1) Then
      blacks = blacks + 2
      v = &H85
    Else
      v = &H84
    End If
    Frame(width - 7 + i, 8) = v
    If (i = 0) Then
      Frame(8, 7) = v
    Else
      Frame(8, 6 - i) = v
    End If
    Format = zzShiftRight(Format, 1)
  Next
  zzWriteFormatInfo = blacks
End Function

'' Show markers.
Private Sub zzShowMarkers()
  Dim width As Integer
  Dim Y As Integer
  Dim i As Integer
  Dim X As Integer
  Dim d As Integer
  Dim w As Integer
  Dim j As Integer
  Dim cx As Integer
  Dim cy As Integer
  Dim wo As Integer
  
  Dim Value As Integer
  
  width = mvarCapacity(mvarVersion)(qrCapwidth)
  
  Value = 0
  For Y = 0 To 6
    For i = 0 To 6
      If mvarFinder(Y)(i) = &HC1 Then
        Value = 2
        If (Y >= 2 And Y <= 4) And (i >= 2 And i <= 4) Then Value = 3
        mvarFinal(mvarQuietZone + Y, mvarQuietZone + i) = Value
        mvarFinal(mvarQuietZone + Y, mvarQuietZone + width - 7 + i) = Value
        mvarFinal(mvarQuietZone + width - 7 + Y, mvarQuietZone + i) = Value
      End If
    Next
  Next
  
  '' Alignment pattern
  If (mvarVersion >= 2) Then
    d = mvarAlignPattern(mvarVersion)(1) - mvarAlignPattern(mvarVersion)(0)
    If (d < 0) Then
      w = 2
    Else
      w = Int((width - mvarAlignPattern(mvarVersion)(0)) / d + 2)
    End If
    
    If (w * w - 3 = 1) Then
      X = mvarAlignPattern(mvarVersion)(0)
      Y = mvarAlignPattern(mvarVersion)(0)
      '' Put an alignment marker.
      Value = 0
      For j = 0 To 4
        For i = 0 To 4
          If mvarMarker(j)(i) = &HA1 Then
            Value = 4
            If (j = 2) And (i = 2) Then Value = 5
            mvarFinal(mvarQuietZone + Y - 2 + j, mvarQuietZone + X - 2 + i) = Value
          End If
        Next
      Next
    End If
    
    If (w * w - 3 <> 1) Then
      cx = mvarAlignPattern(mvarVersion)(0)
      wo = w - 1
      For X = 1 To wo - 1
        '' Put an alignment marker.
        Value = 0
        For j = 0 To 4
          For i = 0 To 4
            If mvarMarker(j)(i) = &HA1 Then
              Value = 4
              If (j = 2) And (i = 2) Then Value = 5
              mvarFinal(mvarQuietZone + cx - 2 + j, mvarQuietZone + 4 + i) = Value
              mvarFinal(mvarQuietZone + 4 + j, mvarQuietZone + cx - 2 + i) = Value
            End If
          Next
        Next
        cx = cx + d
      Next
      cy = mvarAlignPattern(mvarVersion)(0)
      For Y = 0 To wo - 1
        cx = mvarAlignPattern(mvarVersion)(0)
        For X = 0 To wo - 1
          '' Put an alignment marker.
          Value = 0
          For j = 0 To 4
            For i = 0 To 4
              If mvarMarker(j)(i) = &HA1 Then
                Value = 4
                If (j = 2) And (i = 2) Then Value = 5
                mvarFinal(mvarQuietZone + cy - 2 + j, mvarQuietZone + cx - 2 + i) = Value
              End If
            Next
          Next
          cx = cx + d
        Next
        cy = cy + d
      Next
    End If
  End If
  
End Sub


