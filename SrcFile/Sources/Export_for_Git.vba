Attribute VB_Name = "Export_for_Git"

' Export all vba components in external files in order to be able to git them
'
' How does it work :
'  1 - Export all vba components in temporary folder (pathTmp)
'  2 - Compare each files exported in temporary folder (pathTmp) with the files exported previously (path)
'         -> if file in pathTemp = path --> no modification in the module since last export --> no action
'         -> if file in pathTemp != path --> module modified --> old version of module in "path" is replaced by new version of "pathTmp"
'

Sub exportComponents()
    Const exportDir As String = "Sources"
    Const tempDir As String = "tmp"

    Dim fso As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim count As Integer
    Dim path As String
    Dim pathTmp As String
    Dim VBComponent As Object
    Dim vbComponentExt As String
    Dim compareResult As Boolean

    path = ActiveWorkbook.path & "\" & exportDir & "\"
    pathTmp = ActiveWorkbook.path & "\" & tempDir & "\"
    compareResult = False
    count = 0

    Debug.Print "====================================================="
    Debug.Print " Starting export..."

    ' Create folders (if needed)
    If Not fso.FolderExists(path) Then
        MkDir (path)
    End If
    If Not fso.FolderExists(pathTmp) Then
        MkDir (pathTmp)
    End If
    Set fso = Nothing

    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        If VBComponent.CodeModule.CountOfLines > 0 Then
            ' Get extension
            Select Case VBComponent.Type
                Case ClassModule, Document
                    vbComponentExt = ".cls"
                Case Form
                    vbComponentExt = ".frm"
                Case Module
                    vbComponentExt = ".bas"
                Case Else
                    vbComponentExt = ".vba"
            End Select

            On Error Resume Next
            Err.Clear

            ' Export in temp folder
            VBComponent.Export (pathTmp & VBComponent.Name & vbComponentExt)

            If Err.Number <> 0 Then
                MsgBox "Failed to export " & VBComponent.Name & " to " & path, vbCritical
                Debug.Print " --> Failed    " & Left$(VBComponent.Name & " :" & Space(30), 30) & path
            Else
                count = count + 1
            End If

            On Error GoTo 0

            ' Compare file exported in temporary folder (pathTmp) with file exported previously (path)
            compareResult = compareFiles(pathTmp & VBComponent.Name & vbComponentExt, path & VBComponent.Name & vbComponentExt)
            If Not compareResult Then
                VBComponent.Export (path & VBComponent.Name & vbComponentExt)
            End If

            Debug.Print " --> Exported  " & Left$(VBComponent.Name & " :" & Space(30), 30) & path

            ' Delete temporary file
            Call DeleteFile(pathTmp & VBComponent.Name & vbComponentExt)
        Else
            Debug.Print " --> Discarded " & Left$(VBComponent.Name & " :" & Space(30), 30) & path
        End If
    Next

    Application.StatusBar = "Successfully exported " & CStr(count) & " components"
End Sub


Public Function compareFiles(ByVal File1 As String, ByVal File2 As String, Optional StringentCheck As Boolean = False) As Boolean
    On Error GoTo ErrorHandler

    If Dir(File1) = "" Or Dir(File2) = "" Then Exit Function

    Dim lLen1 As Long, lLen2 As Long
    Dim iFileNum1 As Integer, iFileNum2 As Integer
    Dim bytArr1() As Byte, bytArr2() As Byte
    Dim lCtr As Long, lStart As Long
    Dim bAns As Boolean

    lLen1 = FileLen(File1)
    lLen2 = FileLen(File2)

    If lLen1 <> lLen2 Then Exit Function

    If StringentCheck = False Then
        compareFiles = True
        Exit Function
    Else
        iFileNum1 = FreeFile
        Open File1 For Binary Access Read As #iFileNum1
        iFileNum2 = FreeFile
        Open File2 For Binary Access Read As #iFileNum2

        'put contents of both into byte Array
        bytArr1() = InputB(LOF(iFileNum1), #iFileNum1)
        bytArr2() = InputB(LOF(iFileNum2), #iFileNum2)
        lLen1 = UBound(bytArr1)
        lStart = LBound(bytArr1)

        bAns = True
        For lCtr = lStart To lLen1
            If bytArr1(lCtr) <> bytArr2(lCtr) Then
                bAns = False
                Exit For
            End If
        Next
        compareFiles = bAns
    End If

ErrorHandler:
    If iFileNum1 > 0 Then Close #iFileNum1
    If iFileNum2 > 0 Then Close #iFileNum2
End Function

Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function

Sub DeleteFile(ByVal File As String)
   If FileExists(File) Then 'See above
      ' First remove readonly attribute, if set
      SetAttr File, vbNormal
      ' Then delete the file
      Kill File
   End If
End Sub

