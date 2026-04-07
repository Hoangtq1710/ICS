import openpyxl
from datetime import datetime

def create_ics_from_vertical_excel(excel_path, calendar_name, file_name):
    # 1. Load workbook
    try:
        wb = openpyxl.load_workbook(excel_path)
        sheet = wb.active
    except Exception as e:
        print(f"Không thể mở file Excel: {e}")
        return

    # Khởi tạo nội dung ICS với Header và Múi giờ Việt Nam
    ics_content = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//HoangTQ//WorldCup 2026//EN",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        f"X-WR-CALNAME:{calendar_name}",
        "X-WR-TIMEZONE:Asia/Ho_Chi_Minh",
        "BEGIN:VTIMEZONE",
        "TZID:Asia/Ho_Chi_Minh",
        "X-LIC-LOCATION:Asia/Ho_Chi_Minh",
        "BEGIN:STANDARD",
        "DTSTART:19700101T000000",
        "TZOFFSETFROM:+0700",
        "TZOFFSETTO:+0700",
        "TZNAME:GMT+7",
        "END:STANDARD",
        "END:VTIMEZONE"
    ]

    # Đọc tất cả các cell vào một list
    all_rows = [cell[0].value for cell in sheet.iter_rows(min_col=1, max_col=1) if cell[0].value is not None]
    
    event_count = 1
    current_timestamp = datetime.now().strftime("%Y%m%dT%H%M%SZ") # Dùng cho DTSTAMP

    for i in range(0, len(all_rows), 4):
        try:
            if i + 2 >= len(all_rows):
                break
            
            raw_time = str(all_rows[i]).strip()      # Dòng 1: 03.04. 15:00
            team1 = str(all_rows[i+1]).strip()       # Dòng 2: Gen.G
            team2 = str(all_rows[i+2]).strip()       # Dòng 3: KT Rolster
            summary = f"{team1} vs {team2}"

            # 2. Parse thời gian (Giả định năm 2026)
            event_time = datetime.strptime(raw_time, "%d.%m. %H:%M")
            event_time = event_time.replace(year=2026)
            
            # Định dạng thời gian cho ICS (Local time, không có chữ Z ở cuối)
            dt_start_local = event_time.strftime("%Y%m%dT%H%M%S")
            # Giả sử mỗi trận đấu kéo dài 2 tiếng (120 phút)
            dt_end_local = (event_time.replace(hour=event_time.hour + 2)).strftime("%Y%m%dT%H%M%S") if event_time.hour < 22 else event_time.strftime("%Y%m%dT235959")
            
            # 3. Tính toán UID
            uid_str = f"{event_count:03d}_{event_time.strftime('%Y%m%d_%H%M')}@wc.com"

            # 4. Tạo block VEVENT theo chuẩn chuyên nghiệp
            ics_content.append("BEGIN:VEVENT")
            ics_content.append(f"UID:{uid_str}")
            ics_content.append(f"DTSTAMP:{current_timestamp}")
            ics_content.append(f"DTSTART;TZID=Asia/Ho_Chi_Minh:{dt_start_local}")
            ics_content.append(f"DTEND;TZID=Asia/Ho_Chi_Minh:{dt_end_local}")
            ics_content.append(f"SUMMARY:{summary}")
            ics_content.append(f"DESCRIPTION:Next match: {summary}")
            ics_content.append("STATUS:CONFIRMED")
            ics_content.append("TRANSP:OPAQUE")

            # Khối Alarm nhắc trước 30 phút (PT30M)
            ics_content.append("BEGIN:VALARM")
            ics_content.append("ACTION:DISPLAY")
            ics_content.append(f"DESCRIPTION:Reminder: {summary}")
            ics_content.append(f"SUMMARY:Next match: {summary}")
            ics_content.append("TRIGGER:-PT30M")
            ics_content.append("X-APPLE-DEFAULT-ALARM:TRUE")
            ics_content.append(f"X-WR-ALARMUID:ALARM_{uid_str}")
            ics_content.append("END:VALARM")
            
            ics_content.append("END:VEVENT")
            
            event_count += 1
            
        except Exception as e:
            print(f"Lỗi tại vị trí dòng {i+1}: {e}")

    ics_content.append("END:VCALENDAR")

    # 5. Xuất file .ics
    output_filename = f"{file_name}.ics"
    with open(output_filename, "w", encoding="utf-8") as f:
        f.write("\n".join(ics_content))
    
    print(f"--- Hoàn thành! Đã tạo {event_count-1} sự kiện vào file: {output_filename} ---")

# Cấu hình đường dẫn
file_name = "wc-schedule-2026"
# Lưu ý: Dùng r"" cho đường dẫn Windows để tránh lỗi escape character
excel_path = r"C:\Python\excel2ics\wc-schedule.xlsx"
calendar_name = "WorldCup 2026"

# --- CHẠY TOOL ---
create_ics_from_vertical_excel(
    excel_path=excel_path, 
    calendar_name=calendar_name,
    file_name=file_name
)