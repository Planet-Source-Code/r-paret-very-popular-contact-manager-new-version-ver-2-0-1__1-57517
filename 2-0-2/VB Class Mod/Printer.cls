VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Printer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'MS Windows API Function Prototypes
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

'---------------------------------------------------------------
' Retreive the vb object "printer" corresponding to the window's
' default printer.
'---------------------------------------------------------------
Public Function GetDefaultPrinter() As Printer

    Dim strBuffer As String * 254
    Dim iRetValue As Long
    Dim strDefaultPrinterInfo As String
    Dim tblDefaultPrinterInfo() As String
    Dim objPrinter As Printer

    ' Retreive current default printer information
    iRetValue = GetProfileString("windows", "device", ",,,", strBuffer, 254)
    strDefaultPrinterInfo = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
    tblDefaultPrinterInfo = Split(strDefaultPrinterInfo, ",")
    For Each objPrinter In Printers
        If objPrinter.DeviceName = tblDefaultPrinterInfo(0) Then
            ' Default printer found !
            Exit For
        End If
    Next

    ' If not found, return nothing
    If objPrinter.DeviceName <> tblDefaultPrinterInfo(0) Then
        Set objPrinter = Nothing
    End If

    Set GetDefaultPrinter = objPrinter
End Function



