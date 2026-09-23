# RASOAT Survey Platform 3.0

Frontend: https://rasoat.pages.dev/
API: 

## Cài đặt
1. Mở Google Apps Script.
2. Copy .
3. Script Properties:  = ID Google Sheet.
4. Chạy  một lần.
5. Deploy Web App: Execute as Me, Who has access Anyone.
6. Frontend đã trỏ sẵn tới API ở trên.

Admin mặc định:  / . Đổi ngay sau khi đăng nhập.

## Các sheet
DSCNV, AdminUsers, Surveys, SurveySections, Questions, QuestionOptions, QuestionRules, Responses, Answers, AuditLog.

## Chức năng
- Survey Manager: tạo, sửa, clone, draft/publish/close/archive, version, thời gian.
- Form Builder: text, textarea, number, date, datetime, email, phone, select, radio, multiselect, checkbox, yes/no, rating, Likert.
- Conditional Logic: show/hide câu hỏi theo câu trả lời.
- Question Bank.
- Dashboard + Responses + chi tiết.
- CSV export.
- Admin password + session + audit log.
- Spreadsheet ID chỉ nằm trong Script Properties.

URL khảo sát: 

Sau này tạo khảo sát mới trong Admin, không cần sửa frontend.
