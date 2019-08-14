Option Explicit
Public Sub HitungFileCRC32Checksum()
    Dim banyakBaris As Long
    Dim kolomNameFile As Long
    Dim barisNameFile As Long
    Dim kolomStatus As Long
    Dim kolomCRC32 As Long
    Dim kolomSHA256 As Long
    Dim i As Long
    Dim fileNamePhoto As String
    Dim dirFotoFile As String
    Dim dirFoto As String
    Dim varBinary  As Variant
    Dim bytArray()  As Byte
                    
    dirFoto = Cells(1, 5).Value 'direktori foto
    kolomNameFile = 2 'kolom yang dicari udah pasti (di B)
    banyakBaris = Cells(2, 5).Value
    banyakBaris = banyakBaris + 3
    
    For i = 4 To banyakBaris
        'method starts here
        
        barisNameFile = i
        kolomStatus = kolomNameFile + 1
        
        If (apakahSelKosong(barisNameFile, kolomNameFile, 1)) = False Then
            
            fileNamePhoto = Cells(barisNameFile, kolomNameFile).Value
            If (apakahAdaExtensiJpg(barisNameFile, kolomNameFile, fileNamePhoto)) Then

                If (apakahFileExist(barisNameFile, kolomNameFile, dirFoto, fileNamePhoto)) Then

                    
                        dirFotoFile = dirFoto & fileNamePhoto
                        varBinary = ReadBinaryFile(dirFotoFile)
                        
                        If Not NoOfDimensionsInArray(varBinary) = 1 Then
                            'notify user if unable to calc checksum
                            Call messageStatus(4, barisNameFile, kolomStatus) 'Status "Crc32 cant be calculate"
                            
                        Else
                            '  calc checksum
                            bytArray = varBinary
                            Erase varBinary
                            
                            kolomCRC32 = kolomStatus + 1
                            Cells(barisNameFile, kolomCRC32).Value = Hex(CRC32(bytArray)) 'nulis crc32 nya
                            
                            kolomSHA256 = kolomCRC32 + 1
                            Cells(barisNameFile, kolomSHA256).Value = FileToSHA256(dirFotoFile, False) 'nulis sha256 nya
                            
                            Call messageStatus(5, barisNameFile, kolomStatus) 'Status "Complete"
                        End If
                    
                Else
                    Call messageStatus(3, barisNameFile, kolomStatus) 'Status "File Not Exist"
                End If
            Else
                Call messageStatus(2, barisNameFile, kolomStatus) 'Status "Tanpa Jpg"
            End If
            
        Else
            Call messageStatus(1, barisNameFile, kolomStatus) 'Status "Selkosong"
        End If
    Next

End Sub

Function messageStatus(angka As Long, baris As Long, kolom As Long)
Dim statusnye As String
Dim warnanye As Long
Dim barisStatus As Long
Dim kolomStatus As Long
Dim angkaPilihan As Long

angkaPilihan = angka

    Select Case angkaPilihan
        Case 1
            warnanye = 3 'red
            statusnye = "Selkosong"
        Case 2
            warnanye = 3 'red
            statusnye = "Tak Ada ekstensi JPG"
        Case 3
            warnanye = 3 'red
            statusnye = "File Not Exist"
        Case 4
            warnanye = 3 'red
            statusnye = "Unable to calculate CRC32 checksum of file. (Value most likely Null or Empty)"
        Case 5
            warnanye = 4 'green
            statusnye = "OK"
    End Select

barisStatus = baris
kolomStatus = kolom
Cells(barisStatus, kolomStatus).Interior.ColorIndex = warnanye
Cells(barisStatus, kolomStatus).Value = statusnye
End Function


Private Sub CalculateFileCRC32Checksum()
    Dim strFilePath As String
    Dim varBinary  As Variant
    Dim bytArray()  As Byte
    Dim Msg        As String

'  select file
 '   strFilePath = Application.GetOpenFilename(FileFilter:="All files, *.*", _
                                        Title:="Pick a file")
 '   If strFilePath = "False" Then 'user cancelled
'        Msg = "No file selected. Goodbye"
'        GoTo NotifyUser
'    End If

'  evaluate binary data
    varBinary = ReadBinaryFile(strFilePath)
    If Not NoOfDimensionsInArray(varBinary) = 1 Then
        'notify user if unable to calc checksum
        Msg = "Unable to calculate CRC32 checksum of file. (Value most likely Null or Empty)"
        GoTo NotifyUser
    End If

'  calc checksum
    bytArray = varBinary
    Erase varBinary
    Msg = "CRC32 Checksum = " & Hex(CRC32(bytArray))

NotifyUser:
    MsgBox Msg
End Sub

Private Function ReadBinaryFile(ByRef strFilePath As String) As Variant
  ' Requires a reference to Microsoft ActiveX Data Objects
    Rem MS ActiveX Data Objects library needs to be at least 2.5 to avoid "user defined type" error on ADODB.Stream
    Rem changed function to read as variant to prevent type mismatch errors if .read results in Empty or Null

    With New ADODB.Stream
        .Type = adTypeBinary
        .Open
        .LoadFromFile strFilePath
        ReadBinaryFile = .Read
    End With
End Function

Private Function CRC32(ByRef aiBuf() As Byte) As Long
    ' Adapted from http://www.vbaccelerator.com/home/VB/Code/Libraries/CRC32/article.asp
    Static aiCRC()  As Long
    Static bInit    As Boolean
    Dim i          As Long
    Dim j          As Long
    Dim iLookup    As Integer

    If Not bInit Then
        Const iPoly As Long = &HEDB88320
        Dim dwCrc  As Long
        ReDim aiCRC(0 To 255)

        For i = 0 To 255
            dwCrc = i

            For j = 8 To 1 Step -1
                If (dwCrc And 1) Then
                    dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                    dwCrc = dwCrc Xor iPoly
                Else
                    dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                End If
            Next j

            aiCRC(i) = dwCrc
        Next i
        bInit = True
    End If

    CRC32 = &HFFFFFFFF

    For i = LBound(aiBuf) To UBound(aiBuf)
        iLookup = (CRC32 And &HFF) Xor aiBuf(i)
        ' shift right 8 bits:
        CRC32 = ((CRC32 And &HFFFFFF00) \ &H100) And &HFFFFFF
        CRC32 = CRC32 Xor aiCRC(iLookup)
    Next i

    CRC32 = Not CRC32
End Function

Private Function NoOfDimensionsInArray(ByVal varArray As Variant) As Byte
'\ returns number of dimensions as 0 - 4
'\ 0 = not an array, 4 = anything above 3 dimensions
    Dim bytDimNum      As Byte
    Dim varErrorCheck  As Variant

    On Error GoTo FinalDimension
    For bytDimNum = 1 To 4
        varErrorCheck = LBound(varArray, bytDimNum)
    Next
FinalDimension:
    On Error GoTo 0
    NoOfDimensionsInArray = bytDimNum - 1
End Function
