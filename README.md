# ZaloExcel v2.0
 Gửi hoặc đặt lịch gửi tin nhắn, hình ảnh, bảng, biểu đồ, tập tin hoặc bộ nhớ tạm với ứng dụng ZaloExcel

Ứng dụng an toàn và miễn phí 100%

Tải về

![zalo_sending](https://user-images.githubusercontent.com/58664571/160992415-44d80779-dc56-4b38-ab25-c68699378309.gif)


Vấn đề bảo mật tài khoản Zalo không bị ảnh hưởng khi sử dụng ứng dụng này.

## Ứng dụng yêu cầu cài đặt Trình Duyệt Chrome và SeleniumBasic

1. Thao tác tay tải và cài đặt SeleniumBasic
 https://github.com/florentbr/SeleniumBasic
2. Thao tác tay cài đặt ứng dụng Chrome và cập nhật Chrome
 https://www.google.com/intl/vi_vn/chrome/
 
 - Cách cập nhật Chrome, (Đóng tất cả Chrome, mở lại gõ chrome://settings/help để cập nhật):
 
![update_chrome](https://user-images.githubusercontent.com/58664571/160245788-15983109-eaca-44dd-a78d-815493e2f7e6.gif)


3. Chạy cập nhật ChromeDriver trước khi đăng nhập (Nút tự động cập nhật)
4. Cần đăng nhập Zalo bằng tay (click nút Đăng nhập để mở trình duyệt)

## HƯỚNG DẪN

Để gửi ảnh chụp màn hình: chỉ cần nhấn nút chụp màn hình [PS] (Print Screen), sau đó chọn ô cột D và nhấn nút BNT để đặt "[Bộ nhớ tạm]" để gửi. Để quá trình gửi thành công, vui lòng không thao tác sao chép (không nhấn Ctrl+C).

Chọn kiểu dữ liệu để gửi (hình ảnh hướng dẫn):

![zalo_pick](https://user-images.githubusercontent.com/58664571/160544552-41b74783-6fe4-44f8-aa0d-b8c28ffb0df1.gif)


## Rủi ro:
Nếu gửi quá nhiều tin cho nhiều số điện thoại, tài khoản có thể bị Zalo khóa nếu bị phát hiện có hành vi Spam tin nhắn.


Ứng dụng sử dụng Shell và Api để tự động cập nhật driver điều khiển Chrome nên trình duyệt xem là virus, vấn đề này đã nói ở bài viết này
Vì quá trình cài đặt và cập nhật bằng tay rất vất vả nên cần tự động tác vụ để giảm gánh nặng công việc, nên khó tránh thao tác với System, mà thao tác với System thì liên quan đến vấn đề an toàn, nên Trình quét sẽ nhận diện ứng dụng có nguy cơ gây nguy hiểm cho máy tính của bạn.



## AN TOÀN VÀ BẢO MẬT
### Các dòng lệnh trình quét xem là Virus hay mã nguy hiểm bao gồm:
1. Shell """" & chromePath & """" & CmdLn, vbHide
2. URLDownloadToFile(0, eURL, temp & ZIP, 0, 0)
3. FSO.CopyFile temp & EXE, sb2 & EXE
4. VBA.CreateObject("Shell.Application").Namespace(temp & "\").CopyHere .Namespace(temp & ZIP).items

### Các API truy cập bộ nhớ System cũng xem là mã tìm tàng:
- Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
- Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
- Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
- Private Declare PtrSafe Function CloseClipboard Lib "USER32" () As Long
- Private Declare PtrSafe Function OpenClipboard Lib "USER32" (ByVal hwnd As LongPtr) As LongPtr
- Private Declare PtrSafe Function EmptyClipboard Lib "USER32" () As Long
- Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
- Private Declare PtrSafe Function SetClipboardData Lib "USER32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
- Private Declare PtrSafe Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As LongPtr)
  
