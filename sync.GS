Option Explicit

Sub pushDataToGoogleSheet()

Dim nama As String
Dim userGender As String
Dim kelas As String
Dim alamat As String
Dim email As String
Dim kontak As String

Dim Gender As String

Dim xmlHTTP As MSXML2.ServerXMLHTTP60
Dim xmlHTTPHeader As String
Dim xmlID As String

Dim URL_firstPart As String
Dim URL_secPart As String

Dim googleSheetURL As String


If frmStudent.Laki_laki.Value Then
    Gender = "Laki-laki"
Else
    Gender = "Perempuan"
End If

With frmStudent
    nama = .txtNama
    userGender = Gender
    kelas = .cbKelas
    alamat = .txtAlamat
    email = .txtEmail
    kontak = .txtKontak
End With

xmlHTTPHeader = "Content-Type"
xmlID = "application/x-www.form-urlencoded;charset=utf-8"

URL_firstPart = "https://docs.google.com/forms/d/e/1FAIpQLScwC8vU6PHsL0rtBU-ho6FXNfumoSDRQQOkmo7hKMRWLrBBuA/formResponse?ifq"
URL_secPart = "&entry.3939263=" & nisn & "&entry.1905983660=" & nama & "&entry.785752111=" & Gender & "&entry.1554037124=" & kelas & "&entry.796307780=" & alamat & "&entry.1522035242=" & email & "&entry.1560752089=" & kontak & "&submit=submit"

googleSheetURL = URL_firstPart & URL_secPart

Set xmlHTTP = New ServerXMLHTTP60

xmlHTTP.Open "Post", googleSheetURL, False

xmlHTTP.setRequestHeader xmlHTTPHeader, xmlID

xmlHTTP.send

If xmlHTTP.statusText = "OK" Then
    MsgBox "Data berhasil disimpan di google sheet!.", vbInformation
Else
    MsgBox "Periksa Koneksi Internet!", vbCritical
End If

End Sub


