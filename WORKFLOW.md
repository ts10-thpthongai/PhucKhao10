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
