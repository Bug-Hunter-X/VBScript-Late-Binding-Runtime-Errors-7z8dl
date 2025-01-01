Option Explicit

Dim objExcel

'Early Binding (preferred): Add a reference to the Excel object library
Set objExcel = CreateObject("Excel.Application")

'Check if the object is created successfully
If objExcel Is Nothing Then
    MsgBox "Error: Could not create Excel object.", vbCritical
    WScript.Quit
End If

'Now you can safely use objExcel methods
MsgBox "Excel object created successfully."

'Clean up
objExcel.Quit
Set objExcel = Nothing

'Alternative: Runtime Check (for situations where early binding is not feasible)
'Dim objFileSystem
'On Error Resume Next 'Handle potential errors gracefully
'Set objFileSystem = CreateObject("Scripting.FileSystemObject")
'If Err.Number <> 0 Then
'    MsgBox "Error creating FileSystemObject: " & Err.Description, vbCritical
'    Err.Clear
'Else
'    'Use objFileSystem
'    MsgBox "FileSystemObject created successfully."
'    Set objFileSystem = Nothing
'End If
'On Error GoTo 0