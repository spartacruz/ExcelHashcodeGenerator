Public Function FileToSHA256(sFullPath As String, Optional bB64 As Boolean = False) As String
    'parameter full path with name of file returned in the function as an SHA2-256 hash
    'Set a reference to mscorlib 4.0 64-bit
    
    Dim enc, bytes, outstr As String, pos As Integer
    
    Set enc = CreateObject("System.Security.Cryptography.SHA256Managed")
    'Convert the string to a byte array and hash it
    bytes = GetFileBytes(sFullPath) 'returned as a byte array
    bytes = enc.ComputeHash_2((bytes))
    
    If bB64 = True Then
       FileToSHA256 = ConvToBase64String(bytes)
    Else
       FileToSHA256 = ConvToHexString(bytes)
    End If
        
    Set enc = Nothing

End Function

Private Function GetFileBytes(ByVal sPath As String) As Byte()
    'makes byte array from file
    'Set a reference to mscorlib 4.0 64-bit
    
    Dim lngFileNum As Long, bytRtnVal() As Byte, bTest
    
    lngFileNum = FreeFile
    
    If LenB(Dir(sPath)) Then ''// Does file exist?
        
        Open sPath For Binary Access Read As lngFileNum
        
        'a zero length file content will give error 9 here
        
        ReDim bytRtnVal(0 To LOF(lngFileNum) - 1&) As Byte
        Get lngFileNum, , bytRtnVal
        Close lngFileNum
    Else
        Err.Raise 53 'File not found
    End If
    
    GetFileBytes = bytRtnVal
    
    Erase bytRtnVal

End Function

Function ConvToBase64String(vIn As Variant) As Variant
    'used to produce a base-64 output
    'Set a reference to mscorlib 4.0 64-bit
    
    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.base64"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToBase64String = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function

Function ConvToHexString(vIn As Variant) As Variant
    'used to produce a hex output
    'Set a reference to mscorlib 4.0 64-bit
    
    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToHexString = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function

Function GetFileSize(sFilePath As String, nSize As Double) As Boolean
    'use this to test for a zero file size
    'takes full path as string in sFileSize
    'returns file size in bytes in nSize
    
    Dim fs As FileSystemObject, f As File
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If fs.FileExists(sFilePath) Then
        Set f = fs.GetFile(sFilePath)
        nSize = f.Size
        GetFileSize = True
        Exit Function
    End If

End Function
