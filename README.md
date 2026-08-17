# Clinic Lead Collector — desktop delivery

## Gửi cho đối tác (không cần cài Python)

1. Gửi file `Clinic-Lead-Collector-windows.zip` trong thư mục `dist`.
2. Đối tác giải nén ZIP vào một thư mục có quyền ghi, ví dụ `Documents\Clinic Lead Collector`.
3. Đối tác mở `Clinic Lead Collector.exe`. Lần đầu Windows có thể hỏi quyền tường lửa cho máy chủ cục bộ; chọn Allow.
4. Google Chrome phải được cài sẵn trên Windows. Chrome được mở hiển thị để người dùng tự xác minh CAPTCHA khi Google yêu cầu.

Checkpoint Excel được lưu dưới `%LOCALAPPDATA%\ClinicLeadCollector\checkpoints` và có thể tải ngay trong app. App chạy hoàn toàn local; Internet chỉ được dùng cho Google Maps, Google AI Mode, Gemini và kiểm tra bản cập nhật.

## Cập nhật từ xa qua GitHub

Repository GitHub phải để **Public** để máy đối tác tải cập nhật mà không cần token. Khi app được mở, nếu có GitHub Release mới, thanh bên sẽ hiện nút **Cập nhật app và khởi động lại**. Nút này tải bản mới, thay file, rồi mở lại app. Gemini key, Excel và dữ liệu lead không được gửi lên GitHub.

Mỗi lần đẩy thay đổi lên nhánh `main`, GitHub Actions tự build ZIP Windows và tạo GitHub Release mới. Đối tác chỉ cần mở app và bấm nút cập nhật.

## Build thủ công

Máy phát triển cần Python 3.11 (64-bit), Google Chrome và Internet. Trong PowerShell ở thư mục dự án:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\build_windows_app.ps1 -Version "local"
```

Gửi file `dist\Clinic-Lead-Collector-windows.zip` cho đối tác.
