

Function apakahSelKosong(baris, kolom, perluMerah As Long) As Boolean 'fungsi untuk ngecheck file kosong

    If IsEmpty(Cells(baris, kolom)) = True Then 'jika sel memang kosong maka
        If perluMerah > 0 Then 'jika sel yang kosong perlu dimerahi maka
            Cells(baris, kolom).Interior.ColorIndex = 3 'isi sel yang kosong dengan warna merah
        End If
        apakahSelKosong = True
        
    Else 'Jika sel tidak kosong
        apakahSelKosong = False
        
    End If

End Function

Function apakahAdaExtensiJpg(baris As Long, kolom As Long, fileName As String) As Boolean 'fungsi untuk ngecheck extensi foto
Dim extensi As String
Dim extensi2 As String
extensi = Right(fileName, 4)
extensi2 = Right(fileName, 5)
    
    If (StrComp(extensi, ".jpg", vbTextCompare) = 0) Or (StrComp(extensi, ".JPG", vbTextCompare) = 0) Or (StrComp(extensi2, ".jpeg", vbTextCompare) = 0) Or (StrComp(extensi2, ".JPEG", vbTextCompare) = 0) Or (StrComp(extensi, ".png", vbTextCompare) = 0) Then 'Jika ada extensi (".jpg") foto maka
        apakahAdaExtensiJpg = True
    Else
        Cells(baris, kolom).Interior.ColorIndex = 3 'isi sel yang kosong dengan warna merah
        apakahAdaExtensiJpg = False
    End If
End Function

Function apakahFileExist(baris As Long, kolom As Long, direktoriFoto As String, fileName As String)
Dim direktoriDanNama As String, garisMiring As String 'inisialisasi variabel

    garisMiring = Right(direktoriFoto, 1) 'mengambil 1 character dari kanan untuk variabel direktoriFoto

    If ((StrComp((garisMiring), "\", vbTextCompare) > 0)) Then 'jika di akhir direktori TIDAK ada slash ("/") maka
        direktoriFoto = direktoriFoto & "/"
    End If

    direktoriDanNama = direktoriFoto & fileName 'penggabungan direktori dan nama file

    If Len(Dir(direktoriDanNama)) > 0 Then 'pengecekan lokasi file, jika ada filenya, maka
        Cells(baris, kolom).Interior.ColorIndex = 4 'sel tsb dikasi warna hijau
        apakahFileExist = True 'file ada bang
    Else 'kalo file ga ada, maka
        Cells(baris, kolom).Interior.ColorIndex = 7 'sel tsb dikasi warna magenta
        apakahFileExist = False 'File ga ada bro
    End If
    
End Function
