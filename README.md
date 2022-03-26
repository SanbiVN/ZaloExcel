# ZaloExcel
 Gửi tin nhắn Zalo từ Ứng dụng Excel


Bài viết này sẽ chia sẻ cho các bạn ứng dụng gửi tin nhắn Zalo trong ứng dụng Excel cho bất kì một người bạn của bạn hoặc một số điện thoại đã đăng ký Zalo.

Ứng dụng yêu cầu cài đặt Trình Duyệt Chrome và SeleniumBasic


1. Để ứng dụng hoạt động cần tải SeleniumBasic
2. Cần cài đặt ứng dụng Chrome và cập nhật chrome
3. Chạy cập nhật ChromeDriver trước khi đăng nhập
4. Cần đăng nhập Zalo bằng tay (click nút Đăng nhập để mở trình duyệt)

*** Rủi ro:
Nếu gửi quá nhiều tin cho nhiều số điện thoại, tài khoản có thể bị Zalo khóa nếu bị phát hiện có hành vi Spam tin nhắn.


Ứng dụng sử dụng Shell và Api để tự động cập nhật driver điều khiển Chrome nên trình duyệt xem là virus, vấn đề này đã nói ở bài viết này

Vì quá trình cài đặt và cập nhật bằng tay rất vất vả nên cần tự động tác vụ để giảm gánh nặng công việc, nên khó tránh thao tác với System, mà thao tác với System thì liên quan đến vấn đề an toàn, nên Trình quét sẽ nhận diện ứng dụng có nguy cơ gây nguy hiểm cho máy tính của bạn.


#### AN TOÀN VÀ BẢO MẬT
######Các dòng lệnh trình quét xem là Virus hay mã nguy hiểm bao gồm:
1. Shell """" & chromePath & """" & CmdLn, vbHide
2. URLDownloadToFile(0, eURL, temp & ZIP, 0, 0)
3. FSO.CopyFile temp & EXE, sb2 & EXE
4. VBA.CreateObject("Shell.Application").Namespace(temp & "\").CopyHere .Namespace(temp & ZIP).items

######Các API truy cập bộ nhớ System cũng xem là mã tìm tàng:
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function CloseClipboard Lib "USER32" () As Long
    Private Declare PtrSafe Function OpenClipboard Lib "USER32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function EmptyClipboard Lib "USER32" () As Long
    Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
    Private Declare PtrSafe Function SetClipboardData Lib "USER32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As LongPtr)
  
