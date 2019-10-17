Attribute VB_Name = "C4ISR_Inventory"
Option Explicit
'Password to unlock page: pwd123

Private Const cmshInventory As String = "Inventory"
Private Const cmshFullInventory As String = "Full Inventory"
Private Const cmshScan As String = "Scan"
Private Const cmshCoverPage As String = "Cover Page"
Private Const cmlScanListEndRow As Long = 5000
Private Const cmlInventoryStartRow As Long = 2
Private Const cmlInventoryEndRow As Long = 5000
Private clLogger As New LoggerClass
Private clMotorola As New MotorolaCS3070Class

Sub C4ISR_Inventory_Start()
  'Description
  'Parameters:
  'Created by: Laszlo Tamas


  On Error GoTo PROC_ERR

  clLogger.logINFO "Program start", "C4ISR_Inventory.C4ISR_Inventory_Start"

  '---------------
PROC_EXIT:
  On Error GoTo 0
  Exit Sub
PROC_ERR:
  Debug.Print "Error in Procedure C4ISR_Inventory.C4ISR_Inventory_Start"
  If Err.Number Then
    clLogger.logERROR Err.Description, "C4ISR_Inventory.C4ISR_Inventory_Start"
  End If
  Resume PROC_EXIT
End Sub

Private Sub C4ISR_InventoryTest()
  'Test procedure for C4ISR_Inventory
  C4ISR_Inventory
End Sub
Private Function WantToDeleteScanned() As Boolean
  'Want ot delete scanned data
  'Returns:{Boolean}
  'Created by: Laszlo Tamas
  'Licence: MIT
  
  Dim bRes As Boolean
  Dim mgConfirm As Long
  
  On Error GoTo FUNC_ERR
  
  bRes = False
  
  mgConfirm = MsgBox("Are you sure?", 36, "Confirm delete scanned data")
  Select Case mgConfirm
    Case vbYes
      bRes = True
    Case vbNo
      bRes = False
  End Select
  WantToDeleteScanned = bRes
FUNC_EXIT:
  On Error GoTo 0
  Exit Function
FUNC_ERR:
  Debug.Print "Error in Function C4ISR_Inventory.WantToDeleteScanned"
  If Err.Number Then
    clLogger.logERROR Err.Description, "C4ISR_Inventory.WantToDeleteScanned"
  End If
  Resume FUNC_EXIT
End Function
Private Sub WantToDeleteScannedTest()
  'Test procedure For WantToDeleteScanned
  'Want ot delete scanned data
  clLogger.logDEBUG WantToDeleteScanned(), "WantToDeleteScannedTest"
End Sub
Private Function WanToDeleteInventory() As Boolean
  'Want ot delete inventory data
  'Returns:{Boolean}
  'Created by: Laszlo Tamas
  'Licence: MIT
  
  Dim bRes As Boolean
  Dim mgConfirm As Long
  
  On Error GoTo FUNC_ERR
  
  bRes = False
  
  mgConfirm = MsgBox("Are you sure?", 36, "Confirm delete inventory data")
  Select Case mgConfirm
    Case vbYes
      bRes = True
    Case vbNo
      bRes = False
  End Select
  WanToDeleteInventory = bRes
FUNC_EXIT:
  On Error GoTo 0
  Exit Function
FUNC_ERR:
  Debug.Print "Error in Function C4ISR_Inventory.WanToDeleteInventory"
  If Err.Number Then
    clLogger.logERROR Err.Description, "C4ISR_Inventory.WanToDeleteInventory"
  End If
  Resume FUNC_EXIT
End Function

Private Sub WanToDeleteInventoryTest()
  'Test procedure For WanToDeleteInventory
  'Want ot delete inventory data
  clLogger.logDEBUG WanToDeleteInventory(), "WanToDeleteInventoryTest"
End Sub
  

Private Function DeleteInventory() As Boolean
  'Delete the content of Inventory sheet
  'Created by: Laszlo Tamas
  'Licence: MIT
  
  Dim lLastRowInventory As Long
  Dim sRowString As String
  Dim bDelete As Boolean
  
  On Error GoTo FUNC_ERROR
  
  bDelete = False
  If WanToDeleteInventory() Then
'    lLastRowInventory = GetLastRow(cmshInventory, 1, 2, 50000, False)
    lLastRowInventory = GetLastRow(cmshInventory, 1)
    sRowString = Trim(CStr(cmlInventoryStartRow)) & ":" & Trim(CStr(lLastRowInventory))
    Sheets(cmshInventory).Select
    Rows(sRowString).Select
    Selection.ClearContents
    Range("A2").Select
    bDelete = True
  End If
  
FUNC_EXIT:
  On Error GoTo 0
  DeleteInventory = bDelete
  Exit Function
FUNC_ERROR:
  Debug.Print "Error in Function C4ISR_Inventory.DeleteInventory"
  If Err.Number Then
    clLogger.logERROR Err.Description, "C4ISR_Inventory.DeleteInventory"
  End If
  Resume FUNC_EXIT
End Function

Sub AddToFullInventory()
  'Add data to Full Inventory
  'Created by: Laszlo Tamas
  'Licence: MIT

  Dim sPath2Inventory() As String
  Dim wbkADSInventory As Workbook
  Dim wbkCurr As Workbook
  Dim lLastRowFullInventory As Long
  Dim ws As Worksheet
  Dim pos As Long
  Dim shInventorySheet As String
  Dim sPackageName As String
  Dim i As Long
  Dim k As Long
  Dim sCell As String
  Dim mgConfirm As Long

  On Error GoTo PROC_ERR

  Set wbkCurr = Application.ActiveWorkbook
  shInventorySheet = vbNullString
  lLastRowFullInventory = GetLastRow(cmshFullInventory, 1)
  clLogger.logDEBUG "Last row of " & cmshFullInventory & " sheet: " & lLastRowFullInventory, "C4ISR_Inventory.Add2FullInventory"
  
  sPath2Inventory = clMotorola.SelectFile(False, "Select Inventory file", "Inventory files", "*.xlsx,*.xlsm")
  If sPath2Inventory(1) <> vbNullString Then
    Set wbkADSInventory = Workbooks.Open(Filename:=sPath2Inventory(1))
    sPackageName = ScanForPackageName(wbkADSInventory)
    ' Check if package is in the Full Inventory already
    For i = 2 To lLastRowFullInventory
      sCell = Trim(CStr(wbkCurr.Sheets(cmshFullInventory).Cells(i, 1)))
      If sCell = sPackageName Then
        mgConfirm = MsgBox("Are you sure to continue?", 36, sPackageName & " is already in the list")
        Select Case mgConfirm
          Case vbYes
            clLogger.logDEBUG "Continue, but " & sPackageName & " is in the list.", "C4ISR_Inventory.Add2FullInventory"
          Case vbNo
            clLogger.logDEBUG "Interrupted because " & sPackageName & " is in the list.", "C4ISR_Inventory.Add2FullInventory"
            wbkADSInventory.Close
            GoTo PROC_EXIT
        End Select
        Exit For
      End If
    Next i
    For Each ws In wbkADSInventory.Worksheets
      clLogger.logDEBUG "Sheet name: " & ws.Name, "C4ISR_Inventory.AddToFullInventory"
      pos = InStr(ws.Name, "Inventory")
      If pos <> 0 Then
        shInventorySheet = ws.Name
        clLogger.logDEBUG "Inventory sheet: " & sPath2Inventory(1) & ">" & shInventorySheet, "C4ISR_Inventory.AddToFullInventory"
        'Copy data
        For i = 2 To 3000
          lLastRowFullInventory = lLastRowFullInventory + 1
          sCell = Trim(CStr(wbkADSInventory.Sheets(shInventorySheet).Cells(i, 1)))
          If sCell = "" Then
            Exit For
          End If
          wbkCurr.Sheets(cmshFullInventory).Cells(lLastRowFullInventory, 1) = sPackageName
          For k = 1 To 8
            wbkCurr.Sheets(cmshFullInventory).Cells(lLastRowFullInventory, k + 1) = Trim(CStr(wbkADSInventory.Sheets(shInventorySheet).Cells(i, k)))
          Next k
        Next i
      End If
    Next
    wbkADSInventory.Close
    clLogger.logDEBUG "Inventory items copied into database", "C4ISR_Inventory.AddToFullInventory"
  End If

PROC_EXIT:
  On Error GoTo 0
  Set wbkADSInventory = Nothing
  Set wbkCurr = Nothing
  Exit Sub
PROC_ERR:
  Debug.Print "Error in Procedure C4ISR_Inventory.AddToFullInventory"
  If Err.Number Then
    clLogger.logERROR Err.Description, "C4ISR_Inventory.AddToFullInventory"
  End If
  Resume PROC_EXIT
End Sub
Private Function ScanForPackageName(ByRef InventoryBook As Workbook) As String
  'Scan for Package name
  'Parameters:
  ' {Workbook} InvetoryBook
  'Returns:{String}
  'Created by: Laszlo Tamas
  Dim pos As Long
  Dim sRes As String
  Dim i As Long
  Dim sCell As String
  
  On Error GoTo FUNC_ERR
  
  sRes = "text"
  'Code here
  For i = 1 To 50
    sCell = Trim(CStr(InventoryBook.Sheets(cmshCoverPage).Cells(i, 1)))
    pos = InStr(sCell, "INVENTORY")
    If pos <> 0 Then
      sRes = Trim(CStr(Replace(sCell, "INVENTORY", "")))
      Exit For
    End If
  Next i
  ScanForPackageName = sRes
FUNC_EXIT:
  On Error GoTo 0
  Exit Function
FUNC_ERR:
  Debug.Print "Error in Function C4ISR_Inventory.ScanForPackageName"
  If Err.Number Then
    clLogger.logERROR Err.Description, "C4ISR_Inventory.ScanForPackageName"
  End If
  Resume FUNC_EXIT
End Function
Private Sub ScanForPackageNameTest()
  'Test procedure for ScanForPackageName
  'Scan for Package name
  Dim wbkInventoryBook As Workbook
  Dim sPath2ADSInventoryFile() As String
  
  ' wbkInvetoryBook = ActiveWorkbook
  sPath2ADSInventoryFile = clMotorola.SelectFile(False, "Select Inventory file", "Inventory files", "*.xlsx,*.xlsm")
  Set wbkInventoryBook = Workbooks.Open(Filename:=sPath2ADSInventoryFile(1))
  
  clLogger.logDEBUG "Package name: " & ScanForPackageName(wbkInventoryBook), "C4ISR_Inventory.ScanForPackageNameTest"
  wbkInventoryBook.Close
  Set wbkInventoryBook = Nothing

End Sub
Sub AddToInventory()
  'Add data to Inventory
  'Created by: Laszlo Tamas
  'Licence: MIT

  Dim sPath2Inventory() As String
  Dim wbkADSInventory As Workbook
  Dim wbkCurr As Workbook
  Dim lLastRowInventory As Long
  Dim ws As Worksheet
  Dim pos As Long
  Dim shInventorySheet As String
  Dim i As Long
  Dim k As Long
  Dim sCell As String
  
  On Error GoTo PROC_ERR

  Set wbkCurr = Application.ActiveWorkbook
  lLastRowInventory = 1
  shInventorySheet = vbNullString
  
  If DeleteInventory() Then
    sPath2Inventory = clMotorola.SelectFile(False, "Select Inventory file", "Inventory files", "*.xlsx,*.xlsm")
    If sPath2Inventory(1) <> vbNullString Then
      Set wbkADSInventory = Workbooks.Open(Filename:=sPath2Inventory(1))
      For Each ws In wbkADSInventory.Worksheets
        clLogger.logDEBUG "Sheet name: " & ws.Name, "C4ISR_Inventory.Add2Inventory"
        pos = InStr(ws.Name, "Inventory")
        If pos <> 0 Then
          shInventorySheet = ws.Name
          clLogger.logDEBUG "Inventory sheet: " & sPath2Inventory(1) & ">" & shInventorySheet, "C4ISR_Inventory.Add2Inventory"
          'Copy data
          For i = cmlInventoryStartRow To cmlInventoryEndRow
            lLastRowInventory = lLastRowInventory + 1
            sCell = Trim(CStr(wbkADSInventory.Sheets(shInventorySheet).Cells(i, 1)))
            If sCell = "" Then
              Exit For
            End If
            For k = 1 To 8
              wbkCurr.Sheets(cmshInventory).Cells(lLastRowInventory, k) = Trim(CStr(wbkADSInventory.Sheets(shInventorySheet).Cells(i, k)))
            Next k
          Next i
        End If
      Next
      wbkADSInventory.Close
      clLogger.logDEBUG "Inventory items copied into database", "C4ISR_Inventory.AddToInventory"
    End If
  End If

PROC_EXIT:
  On Error GoTo 0
  Set wbkADSInventory = Nothing
  Set wbkCurr = Nothing
  Exit Sub
PROC_ERR:
  Debug.Print "Error in Procedure C4ISR_Inventory.AddToInventory"
  If Err.Number Then
    clLogger.logERROR Err.Description, "C4ISR_Inventory.AddToInventory"
  End If
  Resume PROC_EXIT
End Sub

Private Sub DeleteScannedData()
  'Delete scanned datat from sheet
  'Created by: Laszlo Tamas
  'Licence: MIT
  
  Dim sRange As String
  
  On Error GoTo PROC_ERR
  Sheets(cmshScan).Activate
  If WantToDeleteScanned() Then
    Sheets(cmshScan).Select
    sRange = "A2:A" & Trim(CStr(cmlScanListEndRow))
    Range(sRange).Select
    Selection.ClearContents
    Range("A2").Select
  End If
PROC_EXIT:
  On Error GoTo 0
  Exit Sub
PROC_ERR:
  Debug.Print "Error in Procedure C4ISR_Inventory.DeleteScannedData"
  If Err.Number Then
    clLogger.logERROR Err.Description, "C4ISR_Inventory.DeleteScannedData"
  End If
  Resume PROC_EXIT
End Sub
  


Sub ReadFromFile()
  'Read scanned datat from TXT file stored On scanner
  'Created by: Laszlo Tamas
  'Licence: MIT
  Dim textdata() As String
  Dim i As Long
  
  On Error GoTo PROC_ERR
  
  If WantToDeleteScanned() Then
    textdata = clMotorola.GetTextData()
    clLogger.logDEBUG "Scanner data TXT file: " & clMotorola.PathTextData, "C4ISR_Inventory.ReadFromFile"
    If textdata(1) <> "" Then
      For i = 1 To UBound(textdata)
        Sheets(cmshScan).Cells(i + 1, 1) = clMotorola.GetBarcodeDataFromBarcodeLine(textdata(i))
        clLogger.logDEBUG Trim(CStr(i)) & ": " & clMotorola.GetBarcodeDataFromBarcodeLine(textdata(i)), "C4ISR_Inventory.ReadFromFile"
      Next i
    Else
      clLogger.logDEBUG "File is EMPTY", "C4ISR_Inventory.ReadFromFile"
    End If
  End If
PROC_EXIT:
  On Error GoTo 0
  Exit Sub
PROC_ERR:
  Debug.Print "Error in Procedure C4ISR_Inventory.ReadFromFile"
  If Err.Number Then
    clLogger.logERROR Err.Description, "C4ISR_Inventory.ReadFromFile"
  End If
  Resume PROC_EXIT
End Sub

Private Function GetLastRow(ByRef SheetName As String, ByRef CheckColumn As Long, Optional ByRef BackwardCheck As Boolean = False, Optional ByRef FirstRow As Long = 2, Optional ByRef LastRow As Long = 600000) As Long
  'Get last not empty row number
  'Parameters:
  '  {String} SheetName: Sheet name
  '  {Long} CheckColumn: Column check is based on
  '  {Optional Boolean} BackwardCheck: Check is executed backwards
  '  {Optional Long} FirstRow: First checked row
  '  {Optional Long} LastRow: Last checked row
  'Returns:
  '  {Long} Last not empty row of checked column
  'Created by: Laszlo Tamas
  'Licence: MIT
  Dim i As Long
  Dim sCell As String
  Dim sSheet As String
  Dim lStart As Long
  Dim lEnd As Long
  Dim lStep As Long
  Dim lDiff As Long
  Dim bEmpty As Boolean
  Dim bIsEmpty As Boolean
    
  On Error GoTo FUNC_ERR
    
  sSheet = Trim(CStr(SheetName))
  GetLastRow = 0
  lStart = FirstRow
  lEnd = LastRow
  lStep = 1
  lDiff = -1
  bEmpty = True
  If BackwardCheck Then
    lStart = LastRow
    lEnd = FirstRow
    lStep = -1
    lDiff = 0
    bEmpty = False
  End If
  For i = lStart To lEnd Step lStep
    bIsEmpty = False
    sCell = Trim(CStr(Sheets(sSheet).Cells(i, CheckColumn)))
    If sCell = "" Then bIsEmpty = True
    If bEmpty = bIsEmpty Then
      GetLastRow = i + lDiff
      Exit For
    End If
  Next i
FUNC_EXIT:
  On Error GoTo 0
  If GetLastRow < FirstRow Then
    GetLastRow = FirstRow
  End If
  Exit Function
FUNC_ERR:
  If Err.Number Then
    Debug.Print "Error in Function GetLastRow"
    clLogger.logERROR Err.Description, "GetLastRow"
  End If
  Resume FUNC_EXIT
End Function

Private Sub GetLastRowTest()
  'Test procedure for GetLastRow
  'Get last not empty row number
  Dim testSheet As String
  Dim testCol As Long
  Dim testBackward As Boolean
  
  On Error GoTo PROC_ERR
  
  testSheet = "Inventory"
  testCol = 1
  testBackward = False
  clLogger.logDEBUG Trim(CStr(testCol)) & " " & testSheet & " " & CStr(testBackward) & " >> " & Trim(CStr(GetLastRow(testSheet, testCol, testBackward))), "GetLastRowTest"
PROC_EXIT:
  On Error GoTo 0
  Exit Sub
PROC_ERR:
  If Err.Number Then
    Debug.Print "Error in Sub GetLastRowTest >> " & Err.Description
    clLogger.logERROR Err.Description, "GetLastRowTest"
  End If
  Resume PROC_EXIT
End Sub

Sub SwitchToENG()
  Call clMotorola.SwitchToENG
End Sub
Sub SwitchToHUN()
  Call clMotorola.SwitchToHUN
End Sub






