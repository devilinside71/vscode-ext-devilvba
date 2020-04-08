Private Function SampleFunc(SamplePar As String) As String
  'Description
  'Parameters:
  '           {String} SamplePar
  'Returns:{String}
  'Created by: Laszlo Tamas
  'Licence: MIT

  Dim retVal As String

  On Error GoTo FUNC_ERR

  'Code here

  SampleFunc = retVal
  '---------------
FUNC_EXIT:
  On Error GoTo 0
  Exit Function
FUNC_ERR:
  Debug.Print "Error in Function SampleFunc"
  If Err.Number Then
    Call clLogger.logERROR(Err.Description, "SampleFunc")
  End If
  Resume FUNC_EXIT
End Function
Private Sub SampleFuncTest
  'Test procedure for SampleFunc
  'Description
  Dim dtmStartTime As Date
  Dim testVal As String
  testVal = Nothing
  dtmStartTime = Now()
  Call clLogger.logDEBUG("Function SampleFunc test: >> " & SampleFunc(testVal), "SampleFuncTest")
End Sub

Private Sub SampleSub(SamplePar As String)
  'Description
  'Parameters:
  '           {String} SamplePar
  'Created by: Laszlo Tamas
  'Licence: MIT

  On Error GoTo PROC_ERR

  'Code here

  '---------------
PROC_EXIT:
  On Error GoTo 0
  Exit Sub
PROC_ERR:
  Debug.Print "Error in Sub SampleSub"
  If Err.Number Then
    Call clLogger.logERROR(Err.Description, "SampleSub")
  End If
  Resume PROC_EXIT
End Sub
Private Sub SampleSubTest
  'Test procedure for SampleSub
  'Description
  Dim testVal As String
  Dim dtmStartTime As Date
  testVal = Nothing
  dtmStartTime = Now()
  Call SampleSub(testVal)
End Sub



Private Function GetPackName(FrenchPackName As String) As String
  'Get Hungarian packName
  'Parameters:
  '           {String} FrenchPackName
  'Returns:{String}
  'Created by: Laszlo Tamas
  'Licence: MIT

  Dim retVal As String

  On Error GoTo FUNC_ERR

  'Code here

  GetPackName = retVal
  '---------------
FUNC_EXIT:
  On Error GoTo 0
  Exit Function
FUNC_ERR:
  Debug.Print "Error in Function GetPackName"
  If Err.Number Then
    Call clLogger.logERROR(Err.Description, "GetPackName")
  End If
  Resume FUNC_EXIT
End Function
Private Sub GetPackNameTest
  'Test procedure for GetPackName
  'Get Hungarian packName
  Dim dtmStartTime As Date
  Dim testVal As String
  testVal = Nothing
  dtmStartTime = Now()
  Call clLogger.logDEBUG("Function GetPackName test: >> " & GetPackName(testVal), "GetPackNameTest")
End Sub

Private Sub Hwmon(Parameter As String)
  'Collect HardwareMonitor files
  'Parameters:
  '           {String} Parameter
  'Created by: Laszlo Tamas
  'Licence: MIT

  On Error GoTo PROC_ERR

  'Code here

  '---------------
PROC_EXIT:
  On Error GoTo 0
  Exit Sub
PROC_ERR:
  Debug.Print "Error in Sub Hwmon"
  If Err.Number Then
    Call clLogger.logERROR(Err.Description, "Hwmon")
  End If
  Resume PROC_EXIT
End Sub
Private Sub HwmonTest
  'Test procedure for Hwmon
  'Collect HardwareMonitor files
  Dim testVal As String
  Dim dtmStartTime As Date
  testVal = Nothing
  dtmStartTime = Now()
  Call Hwmon(testVal)
End Sub

Private Sub CheckMK(Parameter As String)
  'Collect Check_MK files
  'Parameters:
  '           {String} Parameter
  'Created by: Laszlo Tamas
  'Licence: MIT

  On Error GoTo PROC_ERR

  'Code here

  '---------------
PROC_EXIT:
  On Error GoTo 0
  Exit Sub
PROC_ERR:
  Debug.Print "Error in Sub CheckMK"
  If Err.Number Then
    Call clLogger.logERROR(Err.Description, "CheckMK")
  End If
  Resume PROC_EXIT
End Sub
Private Sub CheckMKTest
  'Test procedure for CheckMK
  'Collect Check_MK files
  Dim testVal As String
  Dim dtmStartTime As Date
  testVal = Nothing
  dtmStartTime = Now()
  Call CheckMK(testVal)
End Sub

Private Function SelctCSVFile(StartPath As String) As String
  'Select one CSV file
  'Parameters:
  '           {String} StartPath
  'Returns:{String}
  'Created by: Laszlo Tamas
  'Licence: MIT

  Dim retVal As String

  On Error GoTo FUNC_ERR

  'Code here

  SelctCSVFile = retVal
  '---------------
FUNC_EXIT:
  On Error GoTo 0
  Exit Function
FUNC_ERR:
  Debug.Print "Error in Function SelctCSVFile"
  If Err.Number Then
    Call clLogger.logERROR(Err.Description, "SelctCSVFile")
  End If
  Resume FUNC_EXIT
End Function
Private Sub SelctCSVFileTest
  'Test procedure for SelctCSVFile
  'Select one CSV file
  Dim dtmStartTime As Date
  Dim testVal As String
  testVal = Nothing
  dtmStartTime = Now()
  Call clLogger.logDEBUG("Function SelctCSVFile test: >> " & SelctCSVFile(testVal), "SelctCSVFileTest")
End Sub

Private Function SelectCSVFiles(StartPath As String) As String
  'Select multiple csv files
  'Parameters:
  '           {String} StartPath
  'Returns:{String}
  'Created by: Laszlo Tamas
  'Licence: MIT

  Dim retVal As String

  On Error GoTo FUNC_ERR

  'Code here

  SelectCSVFiles = retVal
  '---------------
FUNC_EXIT:
  On Error GoTo 0
  Exit Function
FUNC_ERR:
  Debug.Print "Error in Function SelectCSVFiles"
  If Err.Number Then
    Call clLogger.logERROR(Err.Description, "SelectCSVFiles")
  End If
  Resume FUNC_EXIT
End Function
Private Sub SelectCSVFilesTest
  'Test procedure for SelectCSVFiles
  'Select multiple csv files
  Dim dtmStartTime As Date
  Dim testVal As String
  testVal = Nothing
  dtmStartTime = Now()
  Call clLogger.logDEBUG("Function SelectCSVFiles test: >> " & SelectCSVFiles(testVal), "SelectCSVFilesTest")
End Sub


Function SelectFile(ByVal Multiselect As Boolean, _
  ByVal DialogTitle As String, _
        ParamArray FileFilter() As Variant) As Variant
'Open file dialog to select file(s)
'Parameters:
' {Boolean} Multiselect: Allow multiselect
' {String} DialogTitle: Title of the dialog box
' {Optional Variant()} FileFilter: Filter array, 1:Name, 2:Extensions
' Example1: SelectFile(False, "Select File")
' Example2: SelectFile(True, "Select Files")
' Example3: SelectFile(True, "Select Files", "Excel files", "*.xlsx,*.xls,*.xlsm")
'Returns:
' {Variant()} Path(s) to selected file(s)
'Created by: Laszlo Tamas
'Licence: MIT

Dim sPath() As String
Dim iChoice As Long
Dim dialogBox As FileDialog
Dim i As Long

On Error GoTo FUNC_ERR

clLogger.logDEBUG "Select file WIN mode", "MotorolaCS3070Class.SelectFile"
Set dialogBox = Application.FileDialog(msoFileDialogOpen)
dialogBox.AllowMultiSelect = Multiselect
dialogBox.Title = DialogTitle
dialogBox.Filters.Clear
If UBound(FileFilter) = 1 Then
dialogBox.Filters.Add FileFilter(0), FileFilter(1)
End If
iChoice = dialogBox.Show
If iChoice <> 0 Then
For i = 1 To dialogBox.SelectedItems.Count
  ReDim Preserve sPath(i)
  sPath(i) = dialogBox.SelectedItems.Item(i)
Next i
Else
ReDim Preserve sPath(1)
sPath(1) = vbNullString
clLogger.logDEBUG "Nothing selected", "MotorolaCS3070Class.SelectFile"
End If
SelectFile = sPath
Set dialogBox = Nothing
FUNC_EXIT:
On Error GoTo 0
Exit Function
FUNC_ERR:
If Err.Number Then
Debug.Print "Error in Function MotorolaCS3070Class.SelectFile >> " & _
    Err.Description
clLogger.logERROR Err.Description, "MotorolaCS3070Class.SelectFile"
End If
Resume FUNC_EXIT
End Function
