import os
import zipfile
import shutil
import tempfile
import pandas as pd
from lxml import etree
from copy import deepcopy
from pathlib import Path

# =========================
# FILE PATHS
# =========================
BASE_DIR = Path(r"C:\Users\prajw\OneDrive\Documents\Projects\scr_automation")

EXCEL_PATH = str(BASE_DIR / "input" / "Automation data.xlsx")
TEMPLATE_PATH = str(BASE_DIR / "templates" / "SCR_TEMPLATE_UPDATED.docx")
DUTIES_PATH = str(BASE_DIR / "reference" / "Duties Per Unit.xlsx")

OUTPUT_DIR = BASE_DIR / "output"
SUPPORT_OUTPUT_DIR = BASE_DIR / "output" / "support_session_scrs"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
SUPPORT_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# =========================
# XML NAMESPACE
# =========================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"

NS = {"w": W_NS}


def qn(tag: str) -> str:
    prefix, name = tag.split(":")
    return f"{{{NS[prefix]}}}{name}"


def get_text(el) -> str:
    if el is None:
        return ""
    return "".join(el.xpath(".//w:t/text()", namespaces=NS)).strip()


def normalize_text(text: str) -> str:
    return " ".join(str(text).replace("\xa0", " ").split()).strip().lower()


def clear_paragraph(paragraph):
    for child in list(paragraph):
        paragraph.remove(child)


def make_run(text, font_name="Calibri", font_size=10, bold=False):
    r = etree.Element(qn("w:r"))
    rPr = etree.SubElement(r, qn("w:rPr"))

    rFonts = etree.SubElement(rPr, qn("w:rFonts"))
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:eastAsia"), font_name)
    rFonts.set(qn("w:cs"), font_name)

    sz = etree.SubElement(rPr, qn("w:sz"))
    sz.set(qn("w:val"), str(font_size * 2))

    szCs = etree.SubElement(rPr, qn("w:szCs"))
    szCs.set(qn("w:val"), str(font_size * 2))

    if bold:
        etree.SubElement(rPr, qn("w:b"))
        etree.SubElement(rPr, qn("w:bCs"))

    t = etree.SubElement(r, qn("w:t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = str(text)

    return r


def set_sdt_text(sdt, value, font_name="Calibri", font_size=10):
    """
    Fill a content control (dropdown/plain text/date-like control) with visible text.
    """
    sdt_content = sdt.find("w:sdtContent", namespaces=NS)
    if sdt_content is None:
        return False

    paragraphs = sdt_content.xpath(".//w:p", namespaces=NS)
    if not paragraphs:
        p = etree.SubElement(sdt_content, qn("w:p"))
    else:
        p = paragraphs[0]
        clear_paragraph(p)

    pPr = p.find("w:pPr", namespaces=NS)
    if pPr is None:
        pPr = etree.Element(qn("w:pPr"))
        p.insert(0, pPr)

    p.append(make_run(value, font_name=font_name, font_size=font_size, bold=False))
    return True


def set_text_or_first_sdt_in_cell(tc, value, font_name="Calibri", font_size=10):
    """
    If the cell contains an SDT/dropdown, fill that.
    Otherwise write directly into the cell.
    """
    sdts = tc.xpath(".//w:sdt", namespaces=NS)
    if sdts:
        return set_sdt_text(sdts[0], value, font_name=font_name, font_size=font_size)

    return set_any_text_in_cell(tc, value, font_name=font_name, font_size=font_size)


def set_cell_text(tc, value, font_name="Calibri", font_size=10):
    paragraphs = tc.xpath("./w:p", namespaces=NS)
    if not paragraphs:
        p = etree.SubElement(tc, qn("w:p"))
    else:
        p = paragraphs[0]
        clear_paragraph(p)

    pPr = p.find("w:pPr", namespaces=NS)
    if pPr is None:
        pPr = etree.Element(qn("w:pPr"))
        p.insert(0, pPr)

    p.append(make_run(value, font_name=font_name, font_size=font_size, bold=False))

    extra_paras = tc.xpath("./w:p[position()>1]", namespaces=NS)
    for ep in extra_paras:
        tc.remove(ep)


def set_date_sdt_text(sdt, value, font_name="Calibri", font_size=10):
    sdt_content = sdt.find("w:sdtContent", namespaces=NS)
    if sdt_content is None:
        return False

    tc = sdt_content.find("w:tc", namespaces=NS)
    if tc is None:
        return False

    paragraphs = tc.xpath("./w:p", namespaces=NS)
    if not paragraphs:
        p = etree.SubElement(tc, qn("w:p"))
    else:
        p = paragraphs[0]
        clear_paragraph(p)

    pPr = p.find("w:pPr", namespaces=NS)
    if pPr is None:
        pPr = etree.Element(qn("w:pPr"))
        p.insert(0, pPr)

    p.append(make_run(value, font_name=font_name, font_size=font_size, bold=False))

    date_node = sdt.find(".//w:date", namespaces=NS)
    if date_node is not None:
        full_date = date_node.find("w:fullDate", namespaces=NS)
        dt = pd.to_datetime(value, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            iso_val = dt.strftime("%Y-%m-%dT00:00:00Z")
            if full_date is None:
                full_date = etree.SubElement(date_node, qn("w:fullDate"))
            full_date.set(qn("w:val"), iso_val)

    return True


def set_any_text_in_cell(tc, value, font_name="Calibri", font_size=10):
    """
    Force-write text into a cell.
    Tries content controls first, then normal paragraphs.
    """
    sdts = tc.xpath(".//w:sdt", namespaces=NS)
    if sdts:
        for sdt in sdts:
            sdt_content = sdt.find("w:sdtContent", namespaces=NS)
            if sdt_content is None:
                continue

            paragraphs = sdt_content.xpath(".//w:p", namespaces=NS)
            if paragraphs:
                p = paragraphs[0]
                clear_paragraph(p)

                pPr = p.find("w:pPr", namespaces=NS)
                if pPr is None:
                    pPr = etree.Element(qn("w:pPr"))
                    p.insert(0, pPr)

                p.append(make_run(value, font_name=font_name, font_size=font_size, bold=False))
                return True

    paragraphs = tc.xpath("./w:p", namespaces=NS)
    if not paragraphs:
        p = etree.SubElement(tc, qn("w:p"))
    else:
        p = paragraphs[0]
        clear_paragraph(p)

    pPr = p.find("w:pPr", namespaces=NS)
    if pPr is None:
        pPr = etree.Element(qn("w:pPr"))
        p.insert(0, pPr)

    p.append(make_run(value, font_name=font_name, font_size=font_size, bold=False))
    return True


def tick_checkbox_sdt(sdt):
    """
    Tick a Word checkbox content control by updating:
    1) w14:checked = 1
    2) visible symbol in w:sdtContent to ☑
    """
    checked = sdt.find(".//{%s}checked" % W14_NS)
    if checked is not None:
        checked.set("{%s}val" % W14_NS, "1")

    texts = sdt.xpath(".//w:sdtContent//w:t", namespaces=NS)
    if texts:
        texts[0].text = "☑"

    return True


def tick_checkbox_in_paragraph_before_label(tbl, label_text):
    """
    Find the paragraph containing the target label and tick the checkbox
    that appears in the same paragraph before the label text.
    """
    target = normalize_text(label_text)

    paragraphs = tbl.xpath(".//w:p", namespaces=NS)

    for p in paragraphs:
        p_text = normalize_text(get_text(p))
        if target not in p_text:
            continue

        sdts = p.xpath(".//w:sdt[w:sdtPr/w14:checkbox]", namespaces={**NS, "w14": W14_NS})
        if sdts:
            return tick_checkbox_sdt(sdts[0])

    return False


def tick_required_scr_checkboxes(tables, data):
    """
    Tick only if BOTH:
    - End Time exists
    - Location exists
    """
    if not data.get("finish_time") or not data.get("location"):
        print("Skipping checkbox tick (missing End Time or Location).")
        return

    target_table = find_table_by_labels(
        tables,
        ["Activity Type", "Learning and/or Assessment", "Trade Training Centre"]
    )

    if target_table is None:
        print("Checkbox table not found.")
        return

    ok1 = tick_checkbox_in_paragraph_before_label(
        target_table, "Learning and/or Assessment"
    )
    ok2 = tick_checkbox_in_paragraph_before_label(
        target_table, "Trade Training Centre"
    )

    print("Learning and/or Assessment ticked:", ok1)
    print("Trade Training Centre ticked:", ok2)


def get_table_rows(tbl):
    return tbl.xpath("./w:tr", namespaces=NS)


def get_row_cells(row):
    return row.xpath("./w:tc", namespaces=NS)


def find_table_by_labels(tables, required_labels):
    required = [normalize_text(x) for x in required_labels]
    for tbl in tables:
        tbl_text = normalize_text(get_text(tbl))
        if all(label in tbl_text for label in required):
            return tbl
    return None


def find_row_by_label(tbl, label_text):
    target = normalize_text(label_text)
    rows = get_table_rows(tbl)

    for row in rows:
        row_text = normalize_text(get_text(row))
        if target in row_text:
            return row
    return None


def set_value_in_row_next_cell(tbl, label_text, value):
    row = find_row_by_label(tbl, label_text)
    if row is None:
        print(f"Row not found for label: {label_text}")
        return False

    cells = get_row_cells(row)
    if len(cells) < 2:
        print(f"Not enough cells for label: {label_text}")
        return False

    set_cell_text(cells[1], value)
    return True


def set_date_in_row(tbl, label_text, value):
    row = find_row_by_label(tbl, label_text)
    if row is None:
        print(f"Date row not found for label: {label_text}")
        return False

    sdt = row.find(".//w:sdt", namespaces=NS)
    if sdt is not None:
        return set_date_sdt_text(sdt, value)

    cells = get_row_cells(row)
    if len(cells) >= 2:
        set_cell_text(cells[1], value)
        return True

    return False


def format_time_value(value):
    if pd.isna(value) or str(value).strip() == "":
        return ""

    if isinstance(value, pd.Timestamp):
        return value.strftime("%I:%M %p").lstrip("0")

    text = str(value).strip()

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        dt = pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%I:%M %p").lstrip("0")

    dt = pd.to_datetime(value, errors="coerce")
    if pd.notna(dt):
        return dt.strftime("%I:%M %p").lstrip("0")

    for fmt in ("%H:%M:%S", "%H:%M", "%I:%M %p", "%I:%M%p"):
        try:
            dt = pd.to_datetime(text, format=fmt, errors="raise")
            return dt.strftime("%I:%M %p").lstrip("0")
        except Exception:
            pass

    return text


def safe_str(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def format_output_date(value):
    dt = parse_contact_date(value)
    if pd.isna(dt):
        return ""
    return dt.strftime("%d.%m.%Y")


def parse_contact_date(value):
    """
    Parse Date of Contact safely for Australian dd/mm/yyyy style.
    Returns pandas Timestamp or NaT.
    """
    if pd.isna(value) or str(value).strip() == "":
        return pd.NaT

    if isinstance(value, pd.Timestamp):
        return value

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        dt = pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
        if pd.notna(dt):
            return dt
        return pd.NaT

    text = str(value).strip()

    # Try explicit day-first formats first
    possible_formats = [
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%d.%m.%Y",
        "%d/%m/%y",
        "%d-%m-%y",
        "%Y-%m-%d",
    ]

    for fmt in possible_formats:
        try:
            return pd.to_datetime(text, format=fmt, errors="raise")
        except Exception:
            pass

    # Final fallback: still treat as day-first
    return pd.to_datetime(text, dayfirst=True, errors="coerce")


def calculate_total_hours(start_value, finish_value):
    """
    Returns duration as decimal hours string, e.g. 2.5, 1.0, 3.25
    """
    if pd.isna(start_value) or pd.isna(finish_value):
        return ""

    def parse_excel_time(val):
        if isinstance(val, pd.Timestamp):
            return val

        if isinstance(val, (int, float)) and not isinstance(val, bool):
            dt = pd.to_datetime(val, unit="D", origin="1899-12-30", errors="coerce")
            if pd.notna(dt):
                return dt

        dt = pd.to_datetime(val, errors="coerce")
        if pd.notna(dt):
            return dt

        text = str(val).strip()
        for fmt in ("%H:%M:%S", "%H:%M", "%I:%M %p", "%I:%M%p"):
            try:
                return pd.to_datetime(text, format=fmt, errors="raise")
            except Exception:
                pass

        return pd.NaT

    start_dt = parse_excel_time(start_value)
    finish_dt = parse_excel_time(finish_value)

    if pd.isna(start_dt) or pd.isna(finish_dt):
        return ""

    if finish_dt < start_dt:
        finish_dt = finish_dt + pd.Timedelta(days=1)

    total_seconds = (finish_dt - start_dt).total_seconds()
    total_hours = total_seconds / 3600

    return f"{total_hours:.2f}".rstrip("0").rstrip(".")


def build_data_from_row(row):
    student_surname = safe_str(row["Student Surname"])
    student_given_name = safe_str(row["Student Given Name"])
    employer_company = safe_str(row["Employer Company"])
    location = safe_str(row["Location"])
    supervisor_surname = safe_str(row["Supervisor Surname"])
    supervisor_given_name = safe_str(row["Supervisor Given Name"])
    trainer_surname = safe_str(row["Trainer Surname"])
    trainer_given_name = safe_str(row["Trainer Given Name"])

    supervisor_name = f"{supervisor_given_name} {supervisor_surname}".strip()
    student_name = f"{student_given_name} {student_surname}".strip()

    trainer_name = f"{trainer_given_name} {trainer_surname}".strip()
    trainer_name_display = f"{trainer_given_name} {trainer_surname}".strip()

    contact_dt = parse_contact_date(row["Date of Contact"])

    date_of_contact_doc = ""
    date_of_contact_file = ""

    if pd.notna(contact_dt):
        date_of_contact_doc = contact_dt.strftime("%d/%m/%Y")
        date_of_contact_file = contact_dt.strftime("%d.%m.%Y")

    start_time_raw = row["Start Time"] if "Start Time" in row.index else ""
    finish_time_raw = (
        row["End Time"] if "End Time" in row.index
        else row["Finish Time"] if "Finish Time" in row.index
        else ""
    )

    start_time = format_time_value(start_time_raw)
    finish_time = format_time_value(finish_time_raw)
    total_hours = calculate_total_hours(start_time_raw, finish_time_raw)

    commenced_value = safe_str(row["Commenced"])
    commenced_codes = [c.strip() for c in commenced_value.split(",") if c.strip()]

    data = {
        "student_surname": student_surname,
        "student_given_name": student_given_name,
        "student_name": student_name,
        "employer_company": employer_company,
        "location": location,
        "supervisor_surname": supervisor_surname,
        "supervisor_given_name": supervisor_given_name,
        "supervisor_name": supervisor_name,
        "trainer_surname": trainer_surname,
        "trainer_given_name": trainer_given_name,
        "trainer_name": trainer_name,
        "trainer_name_display": trainer_name_display,
        "date_of_contact": date_of_contact_doc,
        "date_of_contact_file": date_of_contact_file,
        "start_time": start_time,
        "finish_time": finish_time,
        "total_hours": total_hours,
        "commenced_codes": commenced_codes,
    }

    return data


def sanitize_filename(name):
    invalid_chars = r'<>:"/\|?*'
    for ch in invalid_chars:
        name = name.replace(ch, "")
    return " ".join(name.split()).strip()


def build_output_filename(data):
    filename = (
        f'{data["student_name"]}_'
        f'{data["date_of_contact_file"]}_'
        f'CPC30220 - Victoria V3.1-1-2 - '
        f'{data["trainer_name_display"]}.docx'
    )
    return sanitize_filename(filename)


def fill_first_top_section(tables, data):
    tbl = tables[0]
    rows = get_table_rows(tbl)

    set_cell_text(get_row_cells(rows[0])[1], data["student_surname"])
    set_cell_text(get_row_cells(rows[0])[3], data["student_given_name"])
    set_cell_text(get_row_cells(rows[2])[1], data["employer_company"])
    set_cell_text(get_row_cells(rows[3])[1], data["supervisor_surname"])
    set_cell_text(get_row_cells(rows[3])[3], data["supervisor_given_name"])
    set_cell_text(get_row_cells(rows[4])[1], data["trainer_surname"])
    set_cell_text(get_row_cells(rows[4])[3], data["trainer_given_name"])
    set_cell_text(get_row_cells(rows[5])[1], data["location"])

    row6 = rows[6]
    sdt = row6.find("./w:sdt", namespaces=NS)
    if sdt is None:
        raise ValueError("Date content control not found in row 6.")
    set_date_sdt_text(sdt, data["date_of_contact"])


def fill_employer_endorsement_section(tables, data):
    tbl = find_table_by_labels(
        tables,
        ["Employer / Workplace Supervisor Name", "Employer / Workplace Supervisor Signature", "Date"]
    )
    if tbl is None:
        print("Employer endorsement table not found.")
        return

    set_value_in_row_next_cell(tbl, "Employer / Workplace Supervisor Name", data["supervisor_name"])
    set_date_in_row(tbl, "Date", data["date_of_contact"])


def fill_student_endorsement_section(tables, data):
    tbl = find_table_by_labels(
        tables,
        ["Student Name", "Student Signature", "Date"]
    )
    if tbl is None:
        print("Student endorsement table not found.")
        return

    set_value_in_row_next_cell(tbl, "Student Name", data["student_name"])
    set_date_in_row(tbl, "Date", data["date_of_contact"])


def fill_name_signature_date_section_for_trainer(tables, data):
    for tbl in tables:
        tbl_text = normalize_text(get_text(tbl))

        if all(x in tbl_text for x in ["name", "signature", "date"]):
            if "student name" in tbl_text or "employer / workplace supervisor name" in tbl_text:
                continue

            set_value_in_row_next_cell(tbl, "Name", data["trainer_name_display"])
            set_date_in_row(tbl, "Date", data["date_of_contact"])
            return

    print("Trainer name/signature/date table not found.")


def fill_site_visit_section(tables, data):
    tbl = find_table_by_labels(
        tables,
        ["Site Visit Date", "Employer Company", "Employer / Workplace Supervisor Name", "Start Time", "Finish Time", "Location"]
    )
    if tbl is None:
        print("Site visit table not found.")
        return

    set_value_in_row_next_cell(tbl, "Trainer Name", data["trainer_name"])
    set_value_in_row_next_cell(tbl, "Employer Company", data["employer_company"])
    set_value_in_row_next_cell(tbl, "Employer / Workplace Supervisor Name", data["supervisor_name"])
    set_value_in_row_next_cell(tbl, "Location", data["location"])
    set_date_in_row(tbl, "Site Visit Date", data["date_of_contact"])

    rows = get_table_rows(tbl)
    for row in rows:
        row_text = normalize_text(get_text(row))
        if "start time" in row_text and "finish time" in row_text:
            cells = get_row_cells(row)
            if len(cells) >= 4:
                set_cell_text(cells[1], data["start_time"])
                if data["finish_time"]:
                    set_cell_text(cells[3], data["finish_time"])
            break


def fill_start_end_total_hours_section(tables, data):
    tbl = find_table_by_labels(
        tables,
        ["Start Time:", "End Time:", "Total Hours"]
    )
    if tbl is None:
        print("Start/End/Total Hours table not found.")
        return

    rows = get_table_rows(tbl)
    if len(rows) < 2:
        print("Start/End/Total Hours value row not found.")
        return

    value_row = rows[1]
    cells = get_row_cells(value_row)

    if len(cells) >= 3:
        set_any_text_in_cell(cells[0], data["start_time"])
        set_any_text_in_cell(cells[1], data["finish_time"])
        set_any_text_in_cell(cells[2], data["total_hours"])
        print("Start Time, End Time and Total Hours filled.")
    else:
        print("Unexpected Start/End/Total Hours row structure.")


def fill_attendance_section_strict(tables, data):
    """
    Fills only:
    Student Name | Student Signature | Time In | Time out | Attendance Type
    Leaves Trainer Endorsement Date and Time untouched.
    """
    tbl = find_table_by_labels(
        tables,
        ["Trainer Signature", "Trainer Endorsement Date and Time", "Student Name", "Time In", "Time out", "Attendance Type"]
    )
    if tbl is None:
        print("Attendance section not found.")
        return

    rows = get_table_rows(tbl)

    for idx, row in enumerate(rows):
        row_text = normalize_text(get_text(row))

        if all(x in row_text for x in ["student name", "student signature", "time in", "time out", "attendance type"]):
            if idx + 1 >= len(rows):
                print("Attendance value row not found.")
                return

            value_row = rows[idx + 1]
            cells = get_row_cells(value_row)

            print("Attendance value row cell count:", len(cells))

            if len(cells) >= 5:
                set_any_text_in_cell(cells[0], data["student_name"])
                set_any_text_in_cell(cells[2], data["start_time"])
                if data["finish_time"]:
                    set_any_text_in_cell(cells[3], data["finish_time"])
                print("Attendance section filled.")
                return

            print("Unexpected attendance row structure.")
            return

    print("Attendance header row not found.")


def clear_trainer_endorsement_datetime(tables):
    """
    Leaves Trainer Endorsement Date and Time blank.
    Clears any accidentally populated text in the adjacent value cell.
    """
    tbl = find_table_by_labels(
        tables,
        ["Trainer Signature", "Trainer Endorsement Date and Time", "Student Name", "Time In", "Time out", "Attendance Type"]
    )
    if tbl is None:
        return

    rows = get_table_rows(tbl)
    for row in rows:
        row_text = normalize_text(get_text(row))
        if "trainer endorsement date and time" in row_text:
            cells = get_row_cells(row)
            if len(cells) >= 4:
                set_any_text_in_cell(cells[3], "")
            elif len(cells) >= 2:
                set_any_text_in_cell(cells[1], "")
            break


def fill_subjects_section(tables, data, duties_df):
    codes = data["commenced_codes"]

    if not codes:
        print("Commenced column is empty.")
        return

    target_tbl = None
    for t in tables:
        txt = normalize_text(get_text(t))
        if (
            "subject code" in txt
            and "subject name" in txt
            and "continued" in txt
            and "completed" in txt
        ):
            target_tbl = t
            break

    if target_tbl is None:
        print("Main subjects table not found.")
        return

    rows = get_table_rows(target_tbl)
    print(f"Processing main subject table, total rows: {len(rows)}")

    write_row = 1

    for code in codes:
        if write_row >= len(rows):
            break

        match = duties_df[
            duties_df["Serial No"].astype(str).str.strip() == code
        ]

        unit_name = ""
        if not match.empty:
            unit_name = str(match.iloc[0]["Unit Name"]).strip()

        if unit_name.upper().startswith(code.upper()):
            unit_name = unit_name[len(code):].lstrip(" -")

        merged_text = f"{code} - {unit_name}".strip()

        row = rows[write_row]
        cells = get_row_cells(row)

        print(f"Writing main row {write_row}, cell count = {len(cells)}, text = {merged_text}")

        if len(cells) >= 4:
            set_cell_text(cells[0], merged_text)

            continued_sdts = cells[1].xpath(
                ".//w:sdt[w:sdtPr/w14:checkbox]",
                namespaces={**NS, "w14": W14_NS}
            )
            if continued_sdts:
                tick_checkbox_sdt(continued_sdts[0])

        write_row += 1

    print("Main subjects table filled successfully.")


def fill_active_subjects_in_site_visit_table(tables, data, duties_df):
    codes = data["commenced_codes"]

    if not codes:
        print("Commenced column is empty.")
        return

    target_tbl = None
    for i, tbl in enumerate(tables):
        txt = normalize_text(get_text(tbl))
        print("Checking table:", txt[:200])

        if "actively engaged in training" in txt and "subject" in txt:
            target_tbl = tbl
            print("✅ Active subjects table found at index:", i)
            break

    if target_tbl is None:
        print("❌ Active subjects table not found.")
        return

    rows = get_table_rows(target_tbl)

    start_row_idx = None
    for idx, row in enumerate(rows):
        row_text = normalize_text(get_text(row))

        if "actively engaged in training" in row_text:
            start_row_idx = idx + 1
            break

    if start_row_idx is None:
        print("❌ Heading row not found.")
        return

    write_row = start_row_idx

    for code in codes:
        if write_row >= len(rows):
            break

        match = duties_df[
            duties_df["Serial No"].astype(str).str.strip() == code
        ]

        unit_name = ""
        if not match.empty:
            unit_name = str(match.iloc[0]["Unit Name"]).strip()

        if unit_name.upper().startswith(code.upper()):
            unit_name = unit_name[len(code):].lstrip(" -")

        row = rows[write_row]
        cells = get_row_cells(row)

        if not cells:
            write_row += 1
            continue

        merged_text = f"{code} - {unit_name}".strip()
        set_cell_text(cells[0], merged_text)

        print(f"✅ Filled row {write_row}: {merged_text}")
        write_row += 1

    print("✅ Active subjects section filled successfully.")


def fill_routine_work_duties_section(tables, data, duties_df):
    """
    Fill Routine Work Duties section using fixed row positions:
    - first rows get duty text
    - remaining rows get this section's own dropdown
    - source dropdown is the LAST visible 'Choose an item' row before 'Terms and Conditions'
    """
    target_tbl = None
    for tbl in tables:
        txt = normalize_text(get_text(tbl))
        if (
            "site visit attendance and engagement record" in txt
            and "routine work duties engagement" in txt
            and "terms and conditions" in txt
        ):
            target_tbl = tbl
            break

    if target_tbl is None:
        print("Routine Work Duties table not found.")
        return

    rows = get_table_rows(target_tbl)

    heading_idx = None
    terms_idx = None

    for i, row in enumerate(rows):
        row_text = normalize_text(get_text(row))

        if (
            heading_idx is None
            and "routine work duties engagement" in row_text
            and "minimum of 1 must be chosen" in row_text
        ):
            heading_idx = i
            continue

        if heading_idx is not None and "terms and conditions" in row_text:
            terms_idx = i
            break

    if heading_idx is None:
        print("Routine Work Duties heading not found.")
        return

    if terms_idx is None:
        print("Terms and Conditions row not found.")
        return

    source_row_idx = None
    source_sdt = None

    for i in range(terms_idx - 1, heading_idx, -1):
        row = rows[i]
        row_text = normalize_text(get_text(row))
        sdts = row.xpath(".//w:sdt", namespaces=NS)

        if sdts and "choose an item" in row_text:
            source_row_idx = i
            source_sdt = sdts[0]
            break

    if source_row_idx is None or source_sdt is None:
        print("Routine Work Duties source dropdown not found.")
        return

    candidate_row_indices = list(range(heading_idx + 1, source_row_idx))

    print("Routine Work heading row:", heading_idx)
    print("Routine Work source row:", source_row_idx)
    print("Routine Work terms row:", terms_idx)
    print("Routine Work candidate rows:", candidate_row_indices)
    print("Commenced codes:", data.get("commenced_codes"))

    duty_texts = []
    commenced_codes = data.get("commenced_codes", [])

    duties_df_copy = duties_df.copy()
    duties_df_copy["Serial No_clean"] = duties_df_copy["Serial No"].astype(str).str.strip().str.upper()

    for code in commenced_codes:
        code_clean = str(code).strip().upper()
        match = duties_df_copy[duties_df_copy["Serial No_clean"] == code_clean]

        if not match.empty and "Duties" in match.columns:
            duty_text = str(match.iloc[0]["Duties"]).strip()
            if duty_text:
                duty_texts.append(duty_text)

    print("Duty texts found:", duty_texts)

    rows_for_text = candidate_row_indices[:len(duty_texts)]
    for row_idx, duty_text in zip(rows_for_text, duty_texts):
        row = rows[row_idx]
        cells = get_row_cells(row)
        if not cells:
            continue

        target_cell = cells[0]

        for child in list(target_cell):
            if child.tag != qn("w:tcPr"):
                target_cell.remove(child)

        set_cell_text(target_cell, duty_text)
        print(f"Filled routine duty row {row_idx}: {duty_text}")

    rows_for_dropdown = candidate_row_indices[len(duty_texts):]
    for row_idx in rows_for_dropdown:
        row = rows[row_idx]
        cells = get_row_cells(row)
        if not cells:
            continue

        target_cell = cells[0]

        for child in list(target_cell):
            if child.tag != qn("w:tcPr"):
                target_cell.remove(child)

        cloned_sdt = deepcopy(source_sdt)
        target_cell.append(cloned_sdt)
        print(f"Added dropdown to row {row_idx}")

    print("Routine Work Duties section filled successfully.")


def remove_empty_paragraph_runs(root):
    """
    Remove only truly empty runs.
    Do not remove paragraph properties or structural paragraphs.
    """
    paragraphs = root.xpath(".//w:p", namespaces=NS)

    for p in paragraphs:
        runs = p.xpath("./w:r", namespaces=NS)
        for r in runs:
            texts = r.xpath(".//w:t/text()", namespaces=NS)
            drawings = r.xpath(".//w:drawing", namespaces=NS)
            brs = r.xpath(".//w:br", namespaces=NS)
            tabs = r.xpath(".//w:tab", namespaces=NS)

            has_visible_content = any(t.strip() for t in texts) or drawings or brs or tabs
            if not has_visible_content:
                try:
                    p.remove(r)
                except Exception:
                    pass


def clone_subject_dropdowns_into_empty_rows(tables):
    """
    Clone the existing dropdown into visually empty rows of the top subjects table
    without breaking Word table structure.
    """
    target_tbl = None
    for tbl in tables:
        txt = normalize_text(get_text(tbl))
        if (
            "subject code" in txt
            and "subject name" in txt
            and "continued" in txt
            and "completed" in txt
            and "trade training centre" in txt
        ):
            target_tbl = tbl
            break

    if target_tbl is None:
        print("Subjects table not found.")
        return

    rows = get_table_rows(target_tbl)

    source_sdt = None
    source_row_index = None

    for r_idx, row in enumerate(rows):
        row_text = normalize_text(get_text(row))
        sdts = row.xpath(".//w:sdt", namespaces=NS)

        if sdts and "choose an item" in row_text:
            source_sdt = sdts[0]
            source_row_index = r_idx
            break

    if source_sdt is None:
        print("Dropdown source not found.")
        return

    print(f"Dropdown source found in row {source_row_index}")

    for r_idx in range(1, source_row_index):
        row = rows[r_idx]
        cells = get_row_cells(row)

        if not cells:
            continue

        target_cell = cells[0]

        if target_cell.xpath(".//w:sdt", namespaces=NS):
            continue

        texts = target_cell.xpath(".//w:t/text()", namespaces=NS)
        visible_text = "".join(t.strip() for t in texts)

        if visible_text != "":
            continue

        for child in list(target_cell):
            if child.tag != qn("w:tcPr"):
                target_cell.remove(child)

        cloned_sdt = deepcopy(source_sdt)
        target_cell.append(cloned_sdt)

        print(f"Dropdown added to row {r_idx}")

    print("Dropdown cloning completed.")


def clone_site_visit_dropdowns_into_empty_rows(tables):
    """
    Copy the exact template dropdown row structure for empty rows in the
    Site Visit subject section.

    This fixes spacing/formatting issues by cloning the whole row, not just the SDT.
    """
    target_tbl = None

    for tbl in tables:
        txt = normalize_text(get_text(tbl))
        if (
            "site visit attendance and engagement record" in txt
            and "subject(s) actively engaged in training" in txt
            and "routine work duties engagement" in txt
        ):
            target_tbl = tbl
            break

    if target_tbl is None:
        print("Site Visit table not found.")
        return

    rows = get_table_rows(target_tbl)

    subject_heading_idx = None
    routine_heading_idx = None

    for i, row in enumerate(rows):
        row_text = normalize_text(get_text(row))

        if (
            subject_heading_idx is None
            and "subject(s) actively engaged in training" in row_text
        ):
            subject_heading_idx = i
            continue

        if (
            subject_heading_idx is not None
            and "routine work duties engagement" in row_text
        ):
            routine_heading_idx = i
            break

    if subject_heading_idx is None:
        print("Subject section heading not found.")
        return

    if routine_heading_idx is None:
        print("Routine Work Duties heading not found.")
        return

    # ---------------------------------------------------------
    # Find the correct source row inside the subject section
    # ---------------------------------------------------------
    source_row_idx = None
    source_row = None

    for i in range(subject_heading_idx + 1, routine_heading_idx):
        row = rows[i]
        row_text = normalize_text(get_text(row))

        if "choose an item" in row_text:
            sdts = row.xpath(".//w:sdt", namespaces=NS)
            if sdts:
                source_row_idx = i
                source_row = row
                break

    if source_row is None:
        print("Subject dropdown source row not found.")
        return

    print(f"Subject heading row: {subject_heading_idx}")
    print(f"Routine heading row: {routine_heading_idx}")
    print(f"Source dropdown row: {source_row_idx}")

    # ---------------------------------------------------------
    # Replace every empty/placeholder row with a clone of the
    # exact source row so formatting matches the template
    # ---------------------------------------------------------
    rows_to_replace = []

    for i in range(subject_heading_idx + 1, routine_heading_idx):
        if i == source_row_idx:
            continue

        row = rows[i]
        row_text = normalize_text(get_text(row))

        # Keep rows that already have real subject text
        if row_text != "" and "choose an item" not in row_text:
            continue

        rows_to_replace.append(row)

    replaced_count = 0

    for row in rows_to_replace:
        parent = row.getparent()
        insert_at = parent.index(row)
        new_row = deepcopy(source_row)
        parent.remove(row)
        parent.insert(insert_at, new_row)
        replaced_count += 1

    print(f"✅ Site Visit dropdown rows replaced with template clones: {replaced_count}")


def fill_scr_template(is_support_session=False):
    df = pd.read_excel(EXCEL_PATH)
    duties_df = pd.read_excel(DUTIES_PATH)

    duties_df["Unit Name"] = duties_df.apply(
        lambda x: str(x["Unit Name"]).split(str(x["Serial No"]))[-1].lstrip(" -"),
        axis=1
    )

    if df.empty:
        raise ValueError("Automation data.xlsx is empty.")

    output_dir = SUPPORT_OUTPUT_DIR if is_support_session else OUTPUT_DIR
    output_dir.mkdir(parents=True, exist_ok=True)

    for idx, row in df.iterrows():
        print("=" * 80)
        print(f"Processing record {idx + 1} of {len(df)}")

        data = build_data_from_row(row)
        output_filename = build_output_filename(data)

        trainer_folder_name = sanitize_filename(data["trainer_name_display"]) or "Unknown Trainer"
        trainer_output_dir = output_dir / trainer_folder_name
        trainer_output_dir.mkdir(parents=True, exist_ok=True)

        record_output_path = trainer_output_dir / output_filename
        temp_dir = tempfile.mkdtemp()

        try:
            with zipfile.ZipFile(TEMPLATE_PATH, "r") as z:
                z.extractall(temp_dir)

            document_xml_path = os.path.join(temp_dir, "word", "document.xml")
            parser = etree.XMLParser(remove_blank_text=False)
            tree = etree.parse(document_xml_path, parser)
            root = tree.getroot()

            tables = root.xpath(".//w:tbl", namespaces=NS)

            if not tables:
                raise ValueError("No tables found in template.")

            fill_first_top_section(tables, data)
            fill_employer_endorsement_section(tables, data)
            fill_student_endorsement_section(tables, data)
            fill_name_signature_date_section_for_trainer(tables, data)
            fill_site_visit_section(tables, data)
            fill_start_end_total_hours_section(tables, data)
            fill_attendance_section_strict(tables, data)
            clear_trainer_endorsement_datetime(tables)
            fill_subjects_section(tables, data, duties_df)
            clone_subject_dropdowns_into_empty_rows(tables)
            tick_required_scr_checkboxes(tables, data)
            fill_active_subjects_in_site_visit_table(tables, data, duties_df)
            clone_site_visit_dropdowns_into_empty_rows(tables)
            fill_routine_work_duties_section(tables, data, duties_df)
            remove_empty_paragraph_runs(root)

            tree.write(
                document_xml_path,
                xml_declaration=True,
                encoding="UTF-8",
                standalone="yes"
            )

            with zipfile.ZipFile(str(record_output_path), "w", zipfile.ZIP_DEFLATED) as out_zip:
                for foldername, _, filenames in os.walk(temp_dir):
                    for filename in filenames:
                        file_path = os.path.join(foldername, filename)
                        arcname = os.path.relpath(file_path, temp_dir)
                        out_zip.write(file_path, arcname)

            print(f"Saved: {record_output_path}")

        except Exception as e:
            print(f"Error while processing row {idx + 1}: {e}")

        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)


def run_scr_automation(excel_path=None, is_support_session=False):
    global EXCEL_PATH

    original_excel_path = EXCEL_PATH

    if excel_path is not None:
        EXCEL_PATH = str(excel_path)

    try:
        fill_scr_template(is_support_session=is_support_session)
    finally:
        EXCEL_PATH = original_excel_path


if __name__ == "__main__":
    run_scr_automation()

