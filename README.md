# ⭐iZalo - Trình điều khiển Zalo Web cho Excel phiên bản 2026

 Gửi hoặc đặt lịch gửi tin nhắn, hình ảnh, bảng, biểu đồ, tập tin hoặc bộ nhớ tạm với Add-in Excel iZalo \
 Quản lý nhóm, tạo bình chọn, lấy kết quả bình chọn.

Ứng dụng an toàn và miễn phí 100%

<p align="center">
<img title="zaloExcel_LoadContact" src="https://github.com/SanbiVN/ZaloExcel/assets/58664571/93cd8193-6a79-4c92-98f6-02a19cf434ec" width="790">
</p>
Vấn đề bảo mật tài khoản Zalo không bị ảnh hưởng khi sử dụng ứng dụng này.

# TẢI XUỐNG
<!-- items that need to be updated release to release -->
[ptUserAddin]: https://github.com/SanbiVN/ZaloExcel/releases/download/izalo_3.1/iZalo_v3.1.zip

[![Lượt tải](https://img.shields.io/github/downloads/SanbiVN/ZaloExcel/total.svg)](https://github.com/SanbiVN/ZaloExcel/releases/) 
|  Tệp   | Tải xuống | Ghi chú |
|--------------|-----------|----------|
| iZalo | [Nhấn tải][ptUserAddin] | Tệp excel xlam |


## ✳HƯỚNG DẪN TÓM TẮT

   - **Tải tệp về** > **Giải nén vào thư mục phù hợp** > **Bỏ block tệp (nếu có)**
   - **Nhấn vào tệp iZalo.xlam để mở với Excel** > **Nhấn đăng nhập để tạo tệp Zalo cá nhân vào thư mục tự chọn**
   - Một tệp iZalo.xlsm sẽ tạo vào thư mục và mở lên. Đó là iZalo chính điều khiển.
   - Để mở lần sau, trên Ribbon chọn tên đăng nhập trước đó và nhấn đăng nhập.

## ✳HƯỚNG DẪN CÀI ĐẶT

Giải nén vào một thư mục được đặt tên phù hợp, sau khi giải nén, vào thông tin tệp ngoài thư mục bỏ unblock tệp trước khi cài đặt nếu có.

<img width="377" height="389" alt="image" src="https://github.com/user-attachments/assets/e8cf3b18-41ab-433f-a873-b32b76e079de" />
<img width="363" height="478" alt="image" src="https://github.com/user-attachments/assets/359bee94-f4b7-4fa2-bc48-23ab7723fd7b" />


(Đừng bỏ cuộc nếu bạn chưa biết đến các bước cài đặt căn bản cho tệp Excel, đó là: Tạo thư mục an toàn trong thiết lập của Excel để Excel nhận diện)

**Cách 1:** 
- Mở trực tiếp Add-in hoặc nhấn chuột vào tệp để mở, trong Excel cần **```Enabled Macro```** để chương trình hoạt động. 
- Nếu chương trình chưa cài đặt khởi động cùng Excel, khi nhấn **BẮT ĐẦU** chương trình sẽ hỏi có cài đặt khởi động vào Excel không?

**Cách 2:** Thực hiện cài đặt Add-in bằng tay: 
  - Nếu chưa có tab Deverloper hiển thị trên thanh Ribbon (Thanh công cụ): nhấn chuột phải vào thanh Ribbon, chọn **```Customize the Ribbon```**.
  - Trong thẻ Deverloper chọn **```Excel Add-ins```**, sau đó chọn nút **```Browse...```** vào thư mục chứa tệp Add-in, đánh dấu Add-in vừa thêm và chọn nút OK 
  - Nếu đã cài đặt vào Excel, nhưng mỗi khi mở ứng dụng không thấy trên thanh Ribbon, thì vào **```Task Manager```** cần End Task ứng dụng Excel chạy ngầm.

 Nếu ứng dụng bị chặn không cho chạy macro thì hãy vào Cài đặt Excel, vào Trust Center, vào tạo đường dẫn thư mục an toàn cho thư mục chứa add-in tải về.

Khi mở ứng dụng iZalo lên lần đầu, bạn sẽ thấy cảnh báo có nút nhấn Enable Macro Hoặc Enable Content, nút này nhấn để cho phép ứng dụng chạy Macro VBA
<p align="center">
<img title="Enable Macro" src="https://github.com/SanbiVN/ZaloExcel/assets/58664571/0176eae9-141f-46e6-aa11-2e82e8bfb1e9" width="260">
<img title="Enable Content" src="https://github.com/SanbiVN/ZaloExcel/assets/58664571/1ec53333-5fe5-4848-b4f1-c192c852f575" width="360">
</p>
Màn hình đầu tiên là hướng dẫn, các bạn hãy đọc từng bước hướng dẫn và làm theo, nhấn vùng trống để đóng hướng dẫn

<p align="center">
<img title="tutorials 1" src="https://github.com/SanbiVN/ZaloExcel/assets/58664571/49e63b11-831e-4d62-9417-ad17349dc14c" width="660">
</p>



## ✳HƯỚNG DẪN SỬ DỤNG

### ✨Cập nhật thành viên trong danh bạ hoặc nhập tay
<p align="center"><img title="Nút nhấn cập nhật" src="https://github.com/SanbiVN/ZaloExcel/assets/58664571/0ea35e45-0bb7-41e5-b1e7-2a782ca6cf04" width="460"></p>

1. Với hộp kiểm [Kèm ảnh đại diện] sẽ tải kèm ảnh đại diện về
2. Nút [TẢI DANH BẠ] sẽ tải danh bạ bạn bè về
3. Nút [TẢI DANH SÁCH THOẠI] sẽ tải danh sách hội thoại về (Danh sách thoại chỉ được bắt đầu lưu trữ từ khi đăng nhập lần đầu)
4. Nhập tay hoặc chép vào danh sách số điện thoại hoặc tên:

<p align="center"><img title="Danh bạ ZaloExcel" src="https://github.com/SanbiVN/ZaloExcel/assets/58664571/84bc277a-b943-49f4-b888-f907b53e3ddf" width="460"></p>

5. Tìm số Zalo và gửi kết bạn

<p align="center"><img title="Tìm số Zalo và gửi kết bạn" src="https://github.com/SanbiVN/ZaloExcel/assets/58664571/45854b95-c144-4192-aa36-39b44d71aa45" width="300"></p>
Sau khi nhập số điện thoại vào danh sách, tích chọn số cần thực hiện và nhấn nút.




Chọn kiểu dữ liệu để gửi (hình ảnh hướng dẫn):

<p align="center"><img title="zalo_pick" src="https://user-images.githubusercontent.com/58664571/160544552-41b74783-6fe4-44f8-aa0d-b8c28ffb0df1.gif" width="460"></p>

### ✨Hướng dẫn gửi tin

(Không nên gửi đồng loạt quá nhiều, Zalo sẽ tự động phát hiện spam và khóa tài khoản) 

#### ✨Cài đặt gửi chung

<p align="center"><img title="Cài đặt gửi chung" src="https://github.com/SanbiVN/ZaloExcel/assets/58664571/b992c17d-7277-4129-94e0-a23a0b0b808f" width="460"></p>
Sau khi đặt các cài đặt, tích chọn mục sẽ gửi, thì các mục này sẽ được gửi cho tất cả các thành viên được tích chọn gửi dưới danh sách.

#### ✨Nhập gửi từng thành viên
<p align="center"><img title="Nhập gửi từng thành viên" src="https://github.com/SanbiVN/ZaloExcel/assets/58664571/ad8d75a4-a8b2-4a04-a320-a80a0caa0fbf" width="460"></p>
Nhập tin văn bản cần gửi vào cột [Tin nhắn]
Nhập tệp, Ô, Đối tượng, Bộ nhớ tạm

Sau khi đã hoàn thành thiết lặp gửi, nhấn nút <img title="Nút gửi" src="https://github.com/SanbiVN/ZaloExcel/assets/58664571/948049bb-28fe-4992-9348-36e55e3a3d35" width="100"> để thực hiện gửi.

Các tùy chọn gửi:
<p align="center"><img title="Các tùy chọn gửi" src="https://github.com/SanbiVN/ZaloExcel/assets/58664571/bd852d4a-f8e6-4ae5-b4d4-306d5f9d2b16" width="460"></p>

1. Gửi đồng thời: chế độ gửi đồng thời của Zalo chỉ gửi được tin văn bản đến nhiều thành viên cùng lúc.
2. Gửi lại tin đã gửi: có thể tin nhắn đã gửi trước đó
3. Máy tính tự động kích hoạt ngủ đông sau khi gửi xong.

### ✨Tạo nhóm, gán thẻ và đổi tên gợi nhớ (*Bắt buộc: Đã là bạn bè) 
<p align="center"><img title="Tạo nhóm, gán thẻ" src="https://github.com/SanbiVN/ZaloExcel/assets/58664571/d5adfe20-692d-4a00-90f5-8e44f3612443" width="460"></p>

- Tạo nhóm hoặc Gán thẻ
   Có hai cách thực hiện điều này:
  + Tạo gán nhóm/thẻ chung
     Nhập tên nhóm/thẻ vào ô [Nhập thẻ/nhóm], sau đó tích chọn thành viên và nhấn nút [Gán thẻ] hoặc [Tạo nhóm]
    
  + Tạo gán nhóm/thẻ riêng lẻ
     Nhập tên nhóm/thẻ riêng lẻ khác nhau vào cột [Tên mới/Nhóm/Thẻ] trong danh sách, nhấn dấu <img title="Tạo nhóm, gán thẻ" src="https://github.com/SanbiVN/ZaloExcel/assets/58664571/9d0f3e6d-5c68-44cc-9703-fd225202acd9" width="160"> phía trên, menu hiện lên nhấn chọn [Tạo nhiều nhóm] hoặc [Tạo nhiều thẻ]
    

- Đổi tên gợi nhớ 
   Nhập tên mới vào cột [Tên mới/Nhóm/Thẻ] và nhấn nút [Đổi tên mới]

### ✨Đặt lịch gửi (Chưa hoàn thiện)
### ✨Quản lý nhóm (Chưa hoàn thiện)
### ✨Tải tin nhắn hội thoại
  Nhấn vào Icon Zalo, di chuyển đến trang ZaloData
  Nhập tên vào danh sách hoặc nhấn nút <img title="Dịch chuyển qua lại các trang tính" src="https://github.com/SanbiVN/ZaloExcel/assets/58664571/ea80c7d9-9935-4971-af1d-f5b9d9fec662" width="230"> để tải danh sách thoại
  Tích chọn tên dưới danh sách sẽ thực hiện thu thập tin nhắn <img title="Dịch chuyển qua lại các trang tính" src="https://github.com/SanbiVN/ZaloExcel/assets/58664571/8981ac87-8b53-4f4d-9357-47fa10739f71" width="130">

### ✨Tính năng di chuyển qua lại nhanh các trang tính
  Nhấn vào Icon Zalo sẽ hiện ra danh sách trang tính, rê chuột lên sẽ tự động Dịch chuyển
  Nhấn vào tên để Activate, nếu không nhấn rê chuột ra ngoài sẽ tự động trở lại Trang tính trước đó
  
<p align="center"><img title="Dịch chuyển qua lại các trang tính" src="https://github.com/SanbiVN/ZaloExcel/assets/58664571/1b760bfc-b5c7-453d-ada2-c94927e71394" width="760"></p>

## Rủi ro gửi tin Zalo tự động:
 1. Nếu gửi quá nhiều tin cho nhiều số điện thoại, tài khoản có thể bị Zalo khóa nếu bị phát hiện có hành vi Spam tin nhắn.
 2. Không gửi 1 tin nhắn cho quá nhiều người (có người lạ) và nhóm. Tài khoản sẽ bị Zalo khóa. 
 3. Tài khoản đã đăng nhập, máy tính của bạn người khác sử dụng gửi tin vào mục đích sai trái. (nên đăng xuất sau khi sử dụng)

## AN TOÀN VÀ BẢO MẬT
Ứng dụng an toàn và không chứa mã độc như trình quét đã phát hiện\
(Ứng dụng miễn phí không có chữ ký số, nên trình quét không duyệt là ứng dụng an toàn)\
Vấn đề bảo mật tài khoản Zalo không bị ảnh hưởng khi sử dụng ứng dụng này.\
  
## THAM GIA NHÓM HỖ TRỢ ZALOEXCEL
Quét mã QR tham gia nhóm hỗ trợ ZaloExcel:
<p align="center"><img title="THAM GIA NHÓM HỖ TRỢ ZALOEXCEL" src="https://github.com/SanbiVN/ZaloExcel/assets/58664571/f8a06068-0452-4757-9981-a75c2e38518c" width="260"></p>
