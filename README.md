# ai-os-reference

Tài liệu tham chiếu công khai của iThunder AI OS. Kho này chỉ theo dõi thư mục `aios-docs/`; các tài liệu khác trong thư mục làm việc (Tờ trình, TARA, NISCI, Tech, PDF giới thiệu, thư mục nội bộ `noi-bo/`) không đưa lên kho.

## aios-docs

| Tệp | Nội dung |
|---|---|
| `iThunder_AI_OS_TL1_Huong_dan_su_dung_AI_Workspace.{docx,pdf}` | TL1 – Hướng dẫn sử dụng AI Workspace, bản 0.13 (18/11/2026): mô hình On-Premises dùng hiểu biết sẵn có, chỉ On-Clouds trả lời theo Kho (1.2, 4.3); nút Sao chép / Lưu DOCX thống nhất (4.7, 8.3); tab AI On-Clouds nhập khóa API của Tổ chức (C.5). Mục 2.8 về ứng dụng Windows 11 và Android |
| `iThunder_AI_OS_TL2_Cai_dat_va_Van_hanh_AI_Platform.{docx,pdf}` | TL2 – Cài đặt và Vận hành AI Platform, bản 0.14 (20/11/2026): bản công khai, nội dung kỹ thuật như bản 0.13 (mục 5.6 về khóa API AI On-Clouds của Tổ chức, quy tắc trả lời theo nơi chạy mô hình ở 5.7 và 8.2, kiểm thử ở 11.1). Phụ lục G là Tài liệu tham chiếu |
| `screenshots/` | Ảnh màn hình dùng trong TL1, TL2 (dữ liệu mẫu, AI Gateway giả lập) |
| `screenshots/_tools/` | Công cụ chụp lại và dựng tài liệu: `mock_docs.py` (AI Gateway giả lập, cổng 8098), `vite.docs.config.mjs` (Vite cổng 5175), `shots.mjs` (Playwright), `add_shots.py` + `docedit.py` (sửa DOCX, chèn ảnh), `doc_versioning.py` (hàm dùng chung để dựng bản mới từ bản trước: gỡ mục lục đã điền, thêm mục "Những điểm mới", thay ảnh), `update_v014.py` (dựng TL2 bản 0.14 từ bản 0.13), `build_toc.py` (điền mục lục, xuất PDF). Ảnh ứng dụng Windows, Android chụp với `MOCK_NETWORK=onpremise`. Đường dẫn thư mục tạm trong các script cần sửa lại trước khi chạy |
| `nginx_de_xuat/` | Đề xuất thay đổi cấu hình nginx cho AI Gateway (`<IP_AI_GATEWAY>` thay bằng IP nội bộ thật khi áp dụng) |
| `mau_tich_hop/` | Mã mẫu MCP Server cho nhà cung cấp phần mềm (Phụ lục E của TL2) |

Mục lục trong DOCX là trường Word đã điền sẵn kết quả (bookmark, liên kết, PAGEREF) bằng `screenshots/_tools/build_toc.py` (tham số `TL1` hoặc `TL2` để dựng một tài liệu; đầu vào là bản chưa điền mục lục đặt ở `_tools/toc_work/<tên>.src.docx`), số trang lấy từ PDF do LibreOffice xuất. Nếu sửa tài liệu trong Word, bấm Update Field rồi lưu lại PDF.

Kho này là công khai: không đưa vào khóa, mật khẩu, địa chỉ IP nội bộ, dữ liệu người dùng thật, cũng như phụ lục và bản lưu trữ dành cho nội bộ.
