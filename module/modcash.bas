Attribute VB_Name = "modcash"
    Private Declare Function ClosePrinter Lib "winspool.drv" _
     (ByVal hPrinter As Long) As Long

    Private Declare Function EndDocPrinter Lib "winspool.drv" _
     (ByVal hPrinter As Long) As Long

    Private Declare Function EndPagePrinter Lib "winspool.drv" _
     (ByVal hPrinter As Long) As Long

    Private Declare Function OpenPrinter Lib "winspool.drv" _
     Alias "OpenPrinterA" _
     (ByVal pPrinterName As String, phPrinter As Long, _
     ByVal pDefault As Long) As Long

    Private Declare Function StartDocPrinter Lib "winspool.drv" _
      Alias "StartDocPrinterA" _
     (ByVal hPrinter As Long, ByVal Level As Long, _
     pDocInfo As DOCINFO) As Long

    Private Declare Function StartPagePrinter Lib "winspool.drv" _
     (ByVal hPrinter As Long) As Long

    Private Declare Function WritePrinter Lib "winspool.drv" _
     (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, _
     pcWritten As Long) As Long

    Private Type DOCINFO
     pDocName As String
     pOutputFile As String
     pDatatype As String
    End Type

    Public Sub openTillDrawerUsb(ByVal sUsbPrinterName As String, _
     ByVal sOpenCodes As String)

     Dim lPrinterHandle As Long
     Dim lpcWritten As Long
     Dim lRet As Long
     Dim sWriteData As String
     Dim MyDocInfo As DOCINFO
     Dim sCodeArray() As String
     Dim i As Integer

     On Error GoTo errError1

     If OpenPrinter(sUsbPrinterName, lPrinterHandle, 0) = 0 Then
         err.Raise 1, , "USB Printer Name specified [" & _
         sUsbPrinterName & _
         "] " & "when trying to open the till drawer wasn't valid"
     End If
     On Error GoTo errError2

     With MyDocInfo
         .pDocName = "DRAWERKICK"
         .pOutputFile = vbNullString
         .pDatatype = vbNullString
     End With

     lRet = StartDocPrinter(lPrinterHandle, 1, MyDocInfo)
     Call StartPagePrinter(lPrinterHandle)

     ' Split cash drawer code list into array
     sCodeArray = Split(sOpenCodes, ",")

     ' Convert array into actual characters to send to printer
     For i = 0 To UBound(sCodeArray)
         sWriteData = sWriteData & Chr$(val(sCodeArray(i)))
     Next

     lRet = WritePrinter(lPrinterHandle, ByVal sWriteData, _
         Len(sWriteData), lpcWritten)
     
     lRet = EndPagePrinter(lPrinterHandle)
     lRet = EndDocPrinter(lPrinterHandle)

     lRet = ClosePrinter(lPrinterHandle)
     On Error GoTo errError1

     Exit Sub

errError2:
     lRet = ClosePrinter(lPrinterHandle)
errError1:
     err.Raise err.Number, , err.Description
    End Sub


