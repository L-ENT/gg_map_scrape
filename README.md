# Clinic Lead Collector — Windows và macOS Apple Silicon

## Gửi bản Windows cho đối tác

1. Gửi file `Clinic-Lead-Collector-windows.zip` trong thư mục `dist`.
2. Đối tác giải nén ZIP vào một thư mục có quyền ghi, ví dụ `Documents\Clinic Lead Collector`.
3. Đối tác mở `Clinic Lead Collector.exe`. Lần đầu Windows có thể hỏi quyền tường lửa cho máy chủ cục bộ; chọn Allow.
4. Google Chrome phải được cài sẵn. Chrome được mở hiển thị để người dùng tự xác minh CAPTCHA khi Google yêu cầu.

Checkpoint Excel được lưu dưới `%LOCALAPPDATA%\ClinicLeadCollector\checkpoints` và có thể tải ngay trong app. App chạy hoàn toàn local; Internet chỉ được dùng cho Google Maps, Google AI Mode, Gemini và kiểm tra bản cập nhật.

## Gửi bản Mac M1/M2/M3/M4 cho đối tác

1. Tải `Clinic-Lead-Collector-macos-arm64.zip` từ GitHub Releases.
2. Giải nén và kéo `Clinic Lead Collector.app` vào thư mục `Applications`.
3. Vì bản hiện tại chưa có Apple Developer ID, lần đầu hãy nhấp chuột phải vào app, chọn **Open**. Nếu macOS vẫn chặn, vào **System Settings → Privacy & Security → Open Anyway**.
4. Google Chrome phải được cài sẵn.

Checkpoint trên Mac được lưu dưới `~/Library/Application Support/ClinicLeadCollector/checkpoints`.

## Cập nhật từ xa qua GitHub

Repository GitHub phải để **Public** để máy đối tác tải cập nhật mà không cần token. Khi app được mở, nếu có GitHub Release mới, thanh bên sẽ hiện nút **Cập nhật app và khởi động lại**. Nút này tải bản mới, thay file, rồi mở lại app. Gemini key, Excel và dữ liệu lead không được gửi lên GitHub.

Mỗi lần đẩy thay đổi lên nhánh `main`, GitHub Actions chạy test trước. Chỉ khi test đạt, workflow mới build cả ZIP Windows và ZIP macOS ARM64 rồi tạo một GitHub Release. App tự chọn đúng gói cập nhật theo hệ điều hành; đối tác chỉ cần mở app và bấm nút cập nhật.

## Build Windows thủ công

Máy phát triển cần Python 3.11 (64-bit), Google Chrome và Internet. Trong PowerShell ở thư mục dự án:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\build_windows_app.ps1 -Version "local"
```

Gửi file `dist\Clinic-Lead-Collector-windows.zip` cho đối tác.

## Build macOS ARM64 thủ công

Trên máy Mac Apple Silicon có Python 3.11 và Xcode Command Line Tools:

```bash
bash ./build_macos_app.sh local
```

File kết quả là `dist/Clinic-Lead-Collector-macos-arm64.zip`.

## Những file cần đưa lên GitHub

Commit toàn bộ mã nguồn, test, script build và workflow. Không commit `dist/`, `build/`, `venv/`, `.packaging-venv/` hoặc `.packaging-venv-macos/`; GitHub Actions tự tạo các file đóng gói này.
