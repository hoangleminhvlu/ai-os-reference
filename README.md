# ai-reference

Tài liệu tham chiếu của iThunder AI OS. Repo này chỉ theo dõi thư mục `aios-docs/`; các tài liệu khác trong thư mục (Tờ trình, TARA, NISCI, Tech, PDF giới thiệu) không đưa lên repo.

## aios-docs

| Tệp | Nội dung |
|---|---|
| `iThunder_AI_OS_TL1_Huong_dan_su_dung_AI_Workspace.{docx,pdf}` | TL1 – Hướng dẫn sử dụng AI Workspace, bản 0.13 (18/11/2026): mô hình On-Premises dùng hiểu biết sẵn có, chỉ On-Clouds trả lời theo Kho (1.2, 4.3); nút Sao chép / Lưu DOCX thống nhất (4.7, 8.3); tab AI On-Clouds nhập khóa API của Tổ chức (C.5). Mục 2.8 về ứng dụng Windows 11 và Android |
| `iThunder_AI_OS_TL2_Cai_dat_va_Van_hanh_AI_Platform.{docx,pdf}` | TL2 – Cài đặt và Vận hành AI Platform, bản 0.13 (18/11/2026): mục 5.6 về khóa API AI On-Clouds của Tổ chức (`cloud_ai/`, mã hóa, `/api/cloud-ai`), quy tắc trả lời theo nơi chạy mô hình (5.7, 8.2), kiểm thử (11.1). Phụ lục G là nội bộ, gỡ trước khi gửi đối tác |
| `archive_v0.12/` | Bản 0.12 của TL1, TL2 (17/11/2026) |
| `archive_v0.11/` | Bản 0.11 của TL1, TL2 (17/11/2026) |
| `archive_v0.10/` | Bản 0.10 của TL1, TL2 (17/11/2026) |
| `archive_v0.9/` | Bản 0.9 (04/11/2026) |
| `screenshots/` | Ảnh màn hình dùng trong TL1, TL2 (dữ liệu mẫu, AI Gateway giả lập) |
| `screenshots/_tools/` | Công cụ chụp lại: `mock_docs.py` (AI Gateway giả lập, cổng 8098), `vite.docs.config.mjs` (Vite cổng 5175), `shots.mjs` (Playwright), `add_shots.py` + `docedit.py` (chèn ảnh vào DOCX), `update_v011.py`, `update_tl2_v011.py` (dựng TL1, TL2 bản 0.11 từ bản 0.10 chưa điền mục lục), `update_v012.py` (dựng bản 0.12 từ bản 0.11), `update_v013.py` (dựng bản 0.13 từ bản 0.12 đã điền mục lục: tự gỡ mục lục cũ, thay ảnh hội thoại nhóm, thêm ảnh `tl1_11_cloud_ai.png`; ảnh tab AI On-Clouds chụp trên AI Gateway giả lập có router `/api/cloud-ai` thật và kho khóa trong bộ nhớ). Ảnh ứng dụng Windows, Android chụp với `MOCK_NETWORK=onpremise`. Đường dẫn thư mục tạm trong các script cần sửa lại trước khi chạy |
| `nginx_de_xuat/` | Đề xuất thay đổi cấu hình nginx cho AI Gateway |
| `mau_tich_hop/` | Mã mẫu MCP Server cho nhà cung cấp phần mềm (Phụ lục E của TL2) |

Mục lục trong DOCX là trường Word đã điền sẵn kết quả (bookmark, liên kết, PAGEREF) bằng `screenshots/_tools/build_toc.py` (tham số `TL1` hoặc `TL2` để dựng một tài liệu; đầu vào là bản chưa điền mục lục đặt ở `_tools/toc_work/<tên>.src.docx`), số trang lấy từ PDF do LibreOffice xuất. Nếu sửa tài liệu trong Word, bấm Update Field rồi lưu lại PDF.
