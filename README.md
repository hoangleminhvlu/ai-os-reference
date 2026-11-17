# ai-reference

Tài liệu tham chiếu của iThunder AI OS. Repo này chỉ theo dõi thư mục `aios-docs/`; các tài liệu khác trong thư mục (Tờ trình, TARA, NISCI, Tech, PDF giới thiệu) không đưa lên repo.

## aios-docs

| Tệp | Nội dung |
|---|---|
| `iThunder_AI_OS_TL1_Huong_dan_su_dung_AI_Workspace.{docx,pdf}` | TL1 – Hướng dẫn sử dụng AI Workspace, bản 0.10 (17/11/2026) |
| `iThunder_AI_OS_TL2_Cai_dat_va_Van_hanh_AI_Platform.{docx,pdf}` | TL2 – Cài đặt và Vận hành AI Platform, bản 0.10 (17/11/2026). Phụ lục G là nội bộ, gỡ trước khi gửi đối tác |
| `archive_v0.9/` | Bản 0.9 (04/11/2026) |
| `screenshots/` | Ảnh màn hình dùng trong TL1, TL2 (dữ liệu mẫu, AI Gateway giả lập) |
| `screenshots/_tools/` | Công cụ chụp lại: `mock_docs.py` (AI Gateway giả lập, cổng 8098), `vite.docs.config.mjs` (Vite cổng 5175), `shots.mjs` (Playwright), `add_shots.py` + `docedit.py` (chèn ảnh vào DOCX). Đường dẫn thư mục tạm trong các script cần sửa lại trước khi chạy |
| `nginx_de_xuat/` | Đề xuất thay đổi cấu hình nginx cho AI Gateway |
| `mau_tich_hop/` | Mã mẫu MCP Server cho nhà cung cấp phần mềm (Phụ lục E của TL2) |

Mục lục trong DOCX là trường Word đã điền sẵn kết quả (bookmark, liên kết, PAGEREF) bằng `screenshots/_tools/build_toc.py`, số trang lấy từ PDF do LibreOffice xuất. Nếu sửa tài liệu trong Word, bấm Update Field rồi lưu lại PDF.
