# Quy trình triển khai (Deployment Workflow)

## Quy trình phát hành chuẩn

1. Sửa mã trên VS Code.
2. Kiểm thử mã trên Apps Script (nếu có thay đổi cần kiểm thử trước khi phát hành).
3. Chạy `git add` và `git commit`.
4. Chạy `git push` để lưu mã nguồn chính thức.
5. Chạy `clasp push` để đồng bộ mã lên Apps Script.
6. Chạy `clasp version` để tạo Version mới.
7. Trong Apps Script, chọn **Deploy → Manage deployments**.
8. Chọn Deployment chính thức, chọn **Edit**, chọn Version mới rồi **Deploy**.
9. Kiểm thử Web App bằng URL chính thức.

## Trường hợp ngoại lệ: chỉnh sửa trực tiếp trên Apps Script

Nếu có chỉnh sửa trực tiếp trên Google Apps Script Editor, thực hiện theo thứ tự:

1. Hoàn thành chỉnh sửa và kiểm thử trên Apps Script.
2. Chạy `clasp pull` để đồng bộ mã mới nhất về VS Code.
3. Kiểm tra thay đổi bằng `git status` và `git diff`.
4. Chạy `git add` và `git commit`.
5. Chạy `git push` để cập nhật GitHub.

Lưu ý:

- Không chạy `clasp push` ngay sau khi vừa sửa trực tiếp trên Apps Script, vì có thể ghi đè thay đổi vừa thực hiện.
- Chỉ chạy `clasp push` khi mã trong VS Code đã được xác nhận là phiên bản mới nhất.

## Quy ước thao tác

| Trường hợp | Thao tác |
|------------|----------|
| Chỉnh sửa mã | VS Code |
| Lưu mã nguồn chính thức | `git push` |
| Đồng bộ Apps Script | `clasp push` |
| Tạo Version | `clasp version` |
| Cập nhật Web App | Edit Deployment |
| New Deployment | Chỉ sử dụng khi cần tạo một Web App mới với URL mới |

## Lưu ý

- Không tạo New Deployment cho mỗi lần cập nhật.
- Luôn cập nhật Deployment hiện có bằng Edit Deployment.
- `clasp version` chỉ tạo Version mới, không tự động cập nhật Deployment.
- Giữ nguyên URL Web App chính thức.
- Deployment Archived chỉ là lịch sử triển khai, không cần xóa nếu không có lý do đặc biệt.
- Sau khi tạo hoặc cập nhật Web App Deployment, phải mở trực tiếp URL `/exec` để kiểm tra Deployment hoạt động đúng trước khi kiểm thử nghiệp vụ.
- Khi tạo Web App Deployment mới, phải sao chép đúng Web app URL, không sử dụng Library URL.
- Sau khi đổi Web App URL, phải cập nhật `CONFIRM_WEB_APP_URL` trong mã nguồn và gửi email kiểm thử mới.
- Email cũ vẫn chứa URL cũ và không tự động được cập nhật.
