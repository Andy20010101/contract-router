import json
import sys
import types
from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
try:
    import tkinter  # noqa: F401
except ModuleNotFoundError:
    tkinter_stub = types.ModuleType("tkinter")
    tkinter_stub.Tk = object
    tkinter_stub.TclError = RuntimeError
    tkinter_stub.ttk = types.SimpleNamespace()
    tkinter_stub.filedialog = types.SimpleNamespace()
    tkinter_stub.messagebox = types.SimpleNamespace()
    tkinter_stub.scrolledtext = types.SimpleNamespace()
    tkinter_stub.simpledialog = types.SimpleNamespace()
    sys.modules["tkinter"] = tkinter_stub
import file_classifier as fc  # noqa: E402


def patch_app_paths(monkeypatch, tmp_path: Path):
    monkeypatch.setattr(fc, "SCRIPT_DIR", tmp_path)
    monkeypatch.setattr(fc, "DATA_DIR", tmp_path / "data")
    monkeypatch.setattr(fc, "DEFAULT_COMPANY_LEDGER", tmp_path / "data" / "公司系统出运统计发票号清单.xlsx")
    monkeypatch.setattr(fc, "DEFAULT_INVOICE_WORKBOOK", tmp_path / "04月发票号清单.xlsx")
    monkeypatch.setattr(fc, "APP_SETTINGS_PATH", tmp_path / "app_settings.json")


def save_workbook(path: Path, sheets: dict[str, list[list[object]]]):
    wb = openpyxl.Workbook()
    first = True
    for title, rows in sheets.items():
        ws = wb.active if first else wb.create_sheet(title)
        ws.title = title
        first = False
        for row in rows:
            ws.append(row)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    wb.close()


def cell_value_by_header(path: Path, sheet_name: str, invoice_no: str, header: str):
    wb = openpyxl.load_workbook(path)
    try:
        ws = wb[sheet_name]
        headers = [cell.value for cell in ws[1]]
        header_col = headers.index(header) + 1
        invoice_col = headers.index("发票号") + 1 if "发票号" in headers else 1
        for row_idx in range(2, ws.max_row + 1):
            if fc.invoice_equivalent(str(ws.cell(row=row_idx, column=invoice_col).value), invoice_no):
                return ws.cell(row=row_idx, column=header_col).value
        return None
    finally:
        wb.close()


def test_find_invoice_workbook_uses_fixed_ledger_and_never_scans_batch(monkeypatch, tmp_path):
    patch_app_paths(monkeypatch, tmp_path)
    settings = fc.default_app_settings()
    ledger = fc.DEFAULT_COMPANY_LEDGER
    ledger.parent.mkdir(parents=True)
    ledger.touch()
    finance_batch = tmp_path / "财务批次表.xlsx"
    finance_batch.touch()

    assert fc.find_invoice_workbook(settings) == ledger

    ledger.unlink()
    assert fc.find_invoice_workbook(settings) is None

    legacy = fc.DEFAULT_INVOICE_WORKBOOK
    legacy.touch()
    assert fc.find_invoice_workbook(settings) == legacy


def test_invoice_matcher_purchase_contract_fallback():
    matcher = fc.InvoiceMatcher(["LY25112915"])

    assert matcher.match("LY25112915 采购合同俊源.jpg") == "LY25112915"
    assert matcher.match("LY99999999 采购合同俊源.jpg") == ""
    assert matcher.match("LY99999999 运费发票.jpg") == ""
    assert matcher.match("采购合同俊源 LY99999999.jpg") == ""


def test_invoice_matcher_purchase_contract_matches_unique_ledger_digits():
    matcher = fc.InvoiceMatcher(["LY25117260", "LY25117261"])

    assert matcher.match("采购合同 优尚 25117260.jpg") == "LY25117260"
    assert matcher.match("合同-优尚-25117260.jpg") == "LY25117260"
    assert matcher.match("采购合同 25117261 莱尔特.jpg") == "LY25117261"
    assert matcher.match("25117260 采购合同 优尚.jpg") == "LY25117260"
    assert matcher.match("采购合同 99999999.jpg") == ""
    assert matcher.match("采购合同 25117260 25117261.jpg") == ""


def test_invoice_matcher_pairs_unique_digit_prefix_to_full_invoice():
    matcher = fc.InvoiceMatcher(["LY25112915", "LY25112916"])

    assert matcher.match("25112915 提单.pdf") == "LY25112915"
    assert matcher.match("251129150 提单.pdf") == ""
    assert fc.strip_invoice_prefix("25112915 提单.pdf", "LY25112915") == "提单.pdf"


def test_invoice_matcher_does_not_pair_ambiguous_digit_prefix():
    matcher = fc.InvoiceMatcher(["LY25112915", "AB25112915"])

    assert matcher.match("25112915 提单.pdf") == ""
    assert matcher.match("LY25112915 提单.pdf") == "LY25112915"


def test_bill_of_lading_matches_scanned_ocr_text_without_filename_keyword(monkeypatch, tmp_path):
    image = tmp_path / "25112915.jpg"
    image.write_bytes(b"image")

    monkeypatch.setattr(
        fc,
        "ocr_text",
        lambda filepath: "BILL OF LADING\nSHIPPER\nCONSIGNEE\nABCD1234567\n12 CARTONS",
    )

    classifier = fc.MaterialClassifier(fc.default_material_config())
    result = classifier.classify_file(str(image), "LY25112915")

    assert result["status"] == "matched"
    assert result["material_id"] == "bill_of_lading"


def complete_material_texts():
    return {
        "random_a.pdf": (
            "PROFORMA INVOICE\n"
            "Invoice No. 25117260\n"
            "JIAXING LAYO\n"
            "Transport details\n"
            "Terms of payment\n"
            "Description of goods\n"
            "POLAND\n"
            "USD 100.00"
        ),
        "scan001.pdf": (
            "报关委托书\n"
            "代理报关委托\n"
            "报关单编号 123456789012345\n"
            "一般贸易\n"
            "177FCWCWS64050\n"
            "7"
        ),
        "20260212.pdf": (
            "电子发票\n"
            "发票号码 12345678\n"
            "国内货物运输代理服务\n"
            "报关\n"
            "嘉兴新锴国际货运代理\n"
            "177FCWCWS640507"
        ),
        "docx_export.pdf": (
            "OCEAN BILL OF LADING\n"
            "SHIPPER\n"
            "CONSIGNEE\n"
            "NOTIFY PARTY\n"
            "FREIGHT COLLECT\n"
            "SHIPPED ON BOARD\n"
            "SHIPPER'S LOAD COUNT\n"
            "B/L No. ESHGDN2600603\n"
            "Container MSBU5039559 Seal FX45652534\n"
            "12 CARTONS\n"
            "1234.56KGS\n"
            "MAY 20, 2026"
        ),
        "wechat_image.png": (
            "CONTAINER LOAD PLAN\n"
            "Container No. MSBU5039559\n"
            "Seal No. FX45652534\n"
            "Port of Loading SHANGHAI\n"
            "Port of Discharge GDANSK\n"
            "Place of Delivery GDANSK\n"
            "Packing Date 2026-05-19\n"
            "Total Packages 12\n"
            "Total Cargo Wt 1234.56KGS\n"
            "Total Meas 10\n"
            "40HC\n"
            "Bill of Lading No. 177FCWCWS64050\n"
            "7"
        ),
        "invoice_unknown.pdf": (
            "电⼦发票\n"
            "发票号码 87654321\n"
            "国际货运代理费\n"
            "港口码头费\n"
            "汇利达欧海国际货运代理\n"
            "B/L No. ESHGDN2600603"
        ),
    }


def test_random_filenames_complete_package_is_matched_and_strong_closed_loop(monkeypatch, tmp_path):
    texts = complete_material_texts()
    folder = tmp_path / "output" / "LY25117260"
    folder.mkdir(parents=True)
    for filename in texts:
        (folder / filename).write_bytes(b"dummy")

    monkeypatch.setattr(fc, "pdf_text", lambda filepath, max_pages=2: texts.get(Path(filepath).name, ""))
    monkeypatch.setattr(fc, "ocr_text", lambda filepath, max_pdf_pages=1: texts.get(Path(filepath).name, ""))

    processor = fc.TaxRefundProcessor(
        str(tmp_path / "output"),
        config_path=tmp_path / "materials_config.json",
        ledger_records={
            "LY25117260": {
                "发票号": "LY25117260",
                "金额": 100,
                "目的国": "波兰",
                "出运日期": "2026-05-20",
            }
        },
    )

    record = processor.evaluate_invoice(
        "LY25117260",
        str(folder),
        allow_prompt=False,
        allow_rename=False,
        save_report=False,
    )

    assert record["状态"] == "可改名"
    assert record["闭环状态"] == "强闭环通过"
    assert record["缺少材料"] == ""
    assert record["需人工确认文件"] == ""
    assert record["闭环冲突"] == ""
    for material_name in ["外销", "报关委托书", "报关费发票", "提单", "装箱单", "运费发票"]:
        assert material_name in record["已识别必需材料"]


def test_end_to_end_incoming_random_files_move_classify_and_close_loop(monkeypatch, tmp_path):
    ledger_path = tmp_path / "4个月.xlsx"
    save_workbook(
        ledger_path,
        {
            "2月": [
                ["发票号", "出运日期", "目的国(地区)", "金额"],
                ["LY25117260", "2026-05-20", "波兰", 100],
                ["LY25117261", "2026-05-21", "德国", 200],
            ],
        },
    )
    ledger_records = fc.load_invoice_records(str(ledger_path))
    matcher = fc.InvoiceMatcher(list(ledger_records.keys()))

    watch_dir = tmp_path / "watch"
    incoming_invoice_folder = watch_dir / "25117260 raw"
    incoming_invoice_folder.mkdir(parents=True)
    output_dir = tmp_path / "output"
    processor = fc.TaxRefundProcessor(
        str(output_dir),
        config_path=tmp_path / "materials_config.json",
        ledger_records=ledger_records,
    )

    texts = complete_material_texts()
    source_files = []
    for filename in texts:
        path = incoming_invoice_folder / filename
        path.write_bytes(b"dummy")
        source_files.append(path)
    contract = watch_dir / "合同-优尚-25117260.jpg"
    contract.write_bytes(b"contract")
    source_files.append(contract)

    moved_names = []
    for source in source_files:
        success, _, _, dest_path, invoice_no, status = fc.move_file_to_output(
            str(source),
            str(output_dir),
            matcher,
            invoice_dir_resolver=processor.resolve_target_folder,
            source_root=str(watch_dir),
        )
        assert success is True
        assert status == "成功"
        assert invoice_no == "LY25117260"
        moved_names.append(Path(dest_path).name)

    invoice_folder = output_dir / "LY25117260"
    assert sorted(moved_names) == sorted([*texts.keys(), "合同-优尚-25117260.jpg"])
    assert not (output_dir / fc.REVIEW_SUBFOLDER).exists()

    monkeypatch.setattr(fc, "pdf_text", lambda filepath, max_pages=2: texts.get(Path(filepath).name, ""))
    monkeypatch.setattr(fc, "ocr_text", lambda filepath, max_pdf_pages=1: texts.get(Path(filepath).name, ""))

    record = processor.evaluate_invoice(
        "LY25117260",
        str(invoice_folder),
        allow_prompt=False,
        allow_rename=False,
        save_report=False,
    )

    assert ledger_records["LY25117260"]["目的国"] == "波兰"
    assert record["状态"] == "可改名"
    assert record["闭环状态"] == "强闭环通过"
    assert record["报关号"] == "123456789012345"
    assert record["缺少材料"] == ""
    assert record["未识别文件"] == ""
    assert record["需人工确认文件"] == ""
    assert record["闭环冲突"] == ""
    assert "采购合同" not in record["缺少材料"]


def test_purchase_contract_filename_declaration_number_is_fallback(monkeypatch, tmp_path):
    texts = complete_material_texts()
    texts["scan001.pdf"] = (
        "报关委托书\n"
        "代理报关委托\n"
        "一般贸易\n"
        "177FCWCWS64050\n"
        "7"
    )
    folder = tmp_path / "output" / "LY25117260 工厂"
    folder.mkdir(parents=True)
    for filename in texts:
        (folder / filename).write_bytes(b"dummy")
    contract = folder / "123456789012345 25117260 采购合同 优尚.jpg"
    contract.write_bytes(b"contract")

    monkeypatch.setattr(fc, "pdf_text", lambda filepath, max_pages=2: texts.get(Path(filepath).name, ""))
    monkeypatch.setattr(fc, "ocr_text", lambda filepath, max_pdf_pages=1: texts.get(Path(filepath).name, ""))

    processor = fc.TaxRefundProcessor(
        str(tmp_path / "output"),
        config_path=tmp_path / "materials_config.json",
        ledger_records={
            "LY25117260": {
                "发票号": "LY25117260",
                "金额": 100,
                "目的国": "波兰",
                "出运日期": "2026-05-20",
            }
        },
    )

    record = processor.evaluate_invoice(
        "LY25117260",
        str(folder),
        allow_prompt=False,
        allow_rename=False,
        save_report=False,
    )

    assert record["报关号"] == "123456789012345"
    assert record["状态"] == "可改名"
    assert record["闭环状态"] == "强闭环通过"
    assert record["缺少材料"] == ""
    assert record["需人工确认文件"] == ""
    assert record["闭环冲突"] == ""


def test_text_normalization_joins_wrapped_identifiers_and_keeps_packing_bl_out_of_ocean_bl():
    assert fc.keyword_match("电子发票", "电⼦发票")

    fields = fc.extract_document_fields(
        "packing_list",
        "CONTAINER LOAD PLAN\nBill of Lading No. 177FCWCWS64050\n7",
        "random.png",
    )

    assert "177FCWCWS640507" in fields["booking_refs"]
    assert fields["bl_numbers"] == []


def test_closed_loop_field_correspondence_and_ledger_missing_unverified():
    texts = complete_material_texts()

    def matched_doc(material_id, filename):
        return {
            "status": "matched",
            "material_id": material_id,
            "fields": fc.extract_document_fields(material_id, texts[filename], filename),
        }

    classifications = [
        matched_doc("export_sales", "random_a.pdf"),
        matched_doc("customs_power_of_attorney", "scan001.pdf"),
        matched_doc("customs_fee_invoice", "20260212.pdf"),
        matched_doc("bill_of_lading", "docx_export.pdf"),
        matched_doc("packing_list", "wechat_image.png"),
        matched_doc("freight_invoice", "invoice_unknown.pdf"),
    ]

    packing_fields = classifications[4]["fields"]
    assert "177FCWCWS640507" in packing_fields["booking_refs"]
    assert packing_fields["bl_numbers"] == []

    closed_loop = fc.evaluate_closed_loop(
        "LY25117260",
        {"发票号": "LY25117260", "金额": 100, "目的国": "波兰", "出运日期": "2026-05-20"},
        classifications,
        "123456789012345",
    )

    assert closed_loop["status"] == "strong_pass"
    assert closed_loop["conflicts"] == []
    assert any("提单号连接提单与运费发票" in item for item in closed_loop["evidence"])
    assert any("箱号连接提单与装箱单" in item for item in closed_loop["evidence"])
    assert any("封号连接提单与装箱单" in item for item in closed_loop["evidence"])

    no_ledger_loop = fc.evaluate_closed_loop(
        "LY25117260",
        {},
        classifications,
        "123456789012345",
    )

    assert no_ledger_loop["status"] == "strong_pass"
    assert no_ledger_loop["conflicts"] == []
    assert "台账金额未核验" in no_ledger_loop["unverified"]


def test_resolve_target_folder_prefers_digits_folder_and_merges_full_invoice_folder(tmp_path):
    output_dir = tmp_path / "output"
    digits_folder = output_dir / "25112915"
    full_invoice_folder = output_dir / "LY25112915"
    digits_folder.mkdir(parents=True)
    full_invoice_folder.mkdir()
    (digits_folder / "外销.pdf").write_text("existing", encoding="utf-8")
    (full_invoice_folder / "提单.pdf").write_text("bill", encoding="utf-8")
    (full_invoice_folder / "外销.pdf").write_text("from full invoice folder", encoding="utf-8")
    processor = fc.TaxRefundProcessor(str(output_dir), config_path=tmp_path / "materials_config.json")

    target = processor.resolve_target_folder("LY25112915")

    assert Path(target) == digits_folder
    assert not full_invoice_folder.exists()
    assert (digits_folder / "提单.pdf").read_text(encoding="utf-8") == "bill"
    assert (digits_folder / "外销.pdf").read_text(encoding="utf-8") == "existing"
    assert (digits_folder / "外销_重复1.pdf").read_text(encoding="utf-8") == "from full invoice folder"


def test_resolve_target_folder_uses_existing_ly_prefixed_child_folder(tmp_path):
    output_dir = tmp_path / "output"
    prefixed_folder = output_dir / "LY25112915 上海客户"
    exact_folder = output_dir / "LY25112915"
    prefixed_folder.mkdir(parents=True)
    exact_folder.mkdir()
    (prefixed_folder / "外销.pdf").write_text("existing", encoding="utf-8")
    (exact_folder / "提单.pdf").write_text("bill", encoding="utf-8")
    (exact_folder / "外销.pdf").write_text("from exact folder", encoding="utf-8")
    processor = fc.TaxRefundProcessor(str(output_dir), config_path=tmp_path / "materials_config.json")

    target = processor.resolve_target_folder("LY25112915")

    assert Path(target) == prefixed_folder
    assert not exact_folder.exists()
    assert (prefixed_folder / "提单.pdf").read_text(encoding="utf-8") == "bill"
    assert (prefixed_folder / "外销.pdf").read_text(encoding="utf-8") == "existing"
    assert (prefixed_folder / "外销_重复1.pdf").read_text(encoding="utf-8") == "from exact folder"


def test_move_file_to_output_uses_existing_digits_folder(tmp_path):
    source = tmp_path / "LY25112915 提单.pdf"
    source.write_bytes(b"data")
    output_dir = tmp_path / "output"
    digits_folder = output_dir / "25112915"
    digits_folder.mkdir(parents=True)
    processor = fc.TaxRefundProcessor(str(output_dir), config_path=tmp_path / "materials_config.json")

    success, subfolder, message, dest_path, invoice_no, status = fc.move_file_to_output(
        str(source),
        str(output_dir),
        fc.InvoiceMatcher(["LY25112915"]),
        invoice_dir_resolver=processor.resolve_target_folder,
    )

    assert success is True
    assert status == "成功"
    assert invoice_no == "LY25112915"
    assert subfolder == "25112915"
    assert "25112915/" in message
    assert Path(dest_path).parent == digits_folder
    assert not (output_dir / "LY25112915").exists()


def test_move_file_to_output_uses_existing_ly_prefixed_child_folder(tmp_path):
    source = tmp_path / "LY25112915 提单.pdf"
    source.write_bytes(b"data")
    output_dir = tmp_path / "output"
    prefixed_folder = output_dir / "LY25112915 上海客户"
    prefixed_folder.mkdir(parents=True)
    processor = fc.TaxRefundProcessor(str(output_dir), config_path=tmp_path / "materials_config.json")

    success, subfolder, message, dest_path, invoice_no, status = fc.move_file_to_output(
        str(source),
        str(output_dir),
        fc.InvoiceMatcher(["LY25112915"]),
        invoice_dir_resolver=processor.resolve_target_folder,
    )

    assert success is True
    assert status == "成功"
    assert invoice_no == "LY25112915"
    assert subfolder == "LY25112915 上海客户"
    assert "LY25112915 上海客户/" in message
    assert Path(dest_path) == prefixed_folder / "LY25112915 提单.pdf"
    assert not (output_dir / "LY25112915").exists()


def test_move_file_to_output_matches_invoice_parent_folder_for_generic_material(tmp_path):
    watch_dir = tmp_path / "watch"
    source = watch_dir / "25112915" / "装箱单.pdf"
    source.parent.mkdir(parents=True)
    source.write_bytes(b"data")
    output_dir = tmp_path / "output"
    processor = fc.TaxRefundProcessor(str(output_dir), config_path=tmp_path / "materials_config.json")

    success, subfolder, message, dest_path, invoice_no, status = fc.move_file_to_output(
        str(source),
        str(output_dir),
        fc.InvoiceMatcher(["LY25112915"]),
        invoice_dir_resolver=processor.resolve_target_folder,
        source_root=str(watch_dir),
    )

    assert success is True
    assert status == "成功"
    assert invoice_no == "LY25112915"
    assert subfolder == "LY25112915"
    assert message == "成功: 装箱单.pdf → LY25112915/"
    assert Path(dest_path) == output_dir / "LY25112915" / "装箱单.pdf"


def test_move_file_to_output_adds_repeat_suffix_for_duplicate_generic_material(tmp_path):
    watch_dir = tmp_path / "watch"
    source = watch_dir / "25112915" / "装箱单.pdf"
    source.parent.mkdir(parents=True)
    source.write_text("new", encoding="utf-8")
    output_dir = tmp_path / "output"
    existing = output_dir / "LY25112915" / "装箱单.pdf"
    existing.parent.mkdir(parents=True)
    existing.write_text("old", encoding="utf-8")
    processor = fc.TaxRefundProcessor(str(output_dir), config_path=tmp_path / "materials_config.json")

    success, subfolder, message, dest_path, invoice_no, status = fc.move_file_to_output(
        str(source),
        str(output_dir),
        fc.InvoiceMatcher(["LY25112915"]),
        invoice_dir_resolver=processor.resolve_target_folder,
        source_root=str(watch_dir),
    )

    assert success is True
    assert status == "成功"
    assert invoice_no == "LY25112915"
    assert Path(dest_path) == output_dir / "LY25112915" / "装箱单_重复1.pdf"
    assert existing.read_text(encoding="utf-8") == "old"
    assert Path(dest_path).read_text(encoding="utf-8") == "new"


def test_load_finance_batch_invoices_reads_headers_first_column_multisheet_and_dedup(tmp_path):
    batch_path = tmp_path / "finance.xlsx"
    save_workbook(
        batch_path,
        {
            "有表头": [
                ["客户", "发票号"],
                ["A", "LY25112915"],
                ["A2", "25112915"],
                ["B", ""],
                ["C", "LY25112916"],
            ],
            "无表头": [
                ["LY25112917", "x"],
                ["LY25112915", "duplicate"],
                [None, "empty"],
            ],
        },
    )

    records = fc.load_finance_batch_invoices(str(batch_path))

    assert [item["invoice_no"] for item in records] == ["LY25112915", "LY25112916", "LY25112917"]
    assert records[0]["sheet"] == "有表头"
    assert records[0]["row"] == 2
    assert records[2]["sheet"] == "无表头"
    assert records[2]["row"] == 1


@pytest.mark.parametrize(
    "record,in_ledger,folder_exists,expected",
    [
        ({"缺少材料": "", "需人工确认文件": "", "状态": "已改名", "闭环状态": "强闭环通过"}, True, True, "齐了"),
        ({"缺少材料": "提单", "需人工确认文件": "", "状态": "缺材料", "闭环状态": "待复核"}, True, True, "缺材料"),
        ({"缺少材料": "", "需人工确认文件": "a.pdf", "状态": "待复核", "闭环状态": "待复核"}, True, True, "待复核"),
        ({"缺少材料": "", "需人工确认文件": "", "状态": "未找到文件夹", "闭环状态": ""}, True, False, "未找到文件夹"),
        ({"缺少材料": "", "需人工确认文件": "", "状态": "未找到文件夹", "闭环状态": ""}, False, False, "不在主台账"),
        ({"缺少材料": "", "需人工确认文件": "", "状态": "不通过", "闭环状态": "不通过"}, True, True, "不通过"),
    ],
)
def test_tax_status_from_record(record, in_ledger, folder_exists, expected):
    assert fc.tax_status_from_record(record, in_ledger, folder_exists) == expected


def test_write_status_to_workbook_adds_columns_multisheet_backup_and_invoice_equivalent(tmp_path):
    workbook_path = tmp_path / "ledger.xlsx"
    save_workbook(
        workbook_path,
        {
            "有表头": [["发票号", "客户"], ["25112915", "A"]],
            "无表头": [["25112916", "B"]],
        },
    )
    statuses = {
        "LY25112915": {
            "退税齐套状态": "齐了",
            "财务批次": "batch",
            "最后更新时间": "2026-05-21 10:00:00",
        },
        "LY25112916": {
            "退税齐套状态": "缺材料",
            "缺少材料": "提单",
            "最后更新时间": "2026-05-21 10:00:00",
        },
    }

    backup_path = fc.write_status_to_workbook(str(workbook_path), statuses)

    assert Path(backup_path).exists()
    wb = openpyxl.load_workbook(workbook_path)
    try:
        ws1 = wb["有表头"]
        headers1 = [cell.value for cell in ws1[1]]
        status_col1 = headers1.index("退税齐套状态") + 1
        assert ws1.cell(row=2, column=status_col1).value == "齐了"

        ws2 = wb["无表头"]
        headers2 = [cell.value for cell in ws2[1]]
        assert headers2[0] == "发票号"
        status_col2 = headers2.index("退税齐套状态") + 1
        missing_col2 = headers2.index("缺少材料") + 1
        assert ws2.cell(row=2, column=1).value == "25112916"
        assert ws2.cell(row=2, column=status_col2).value == "缺材料"
        assert ws2.cell(row=2, column=missing_col2).value == "提单"
    finally:
        wb.close()


def test_evaluate_invoice_batch_mode_has_no_declaration_report_prompt_or_rename_side_effects(monkeypatch, tmp_path):
    output_dir = tmp_path / "output"
    folder = output_dir / "LY25112915"
    folder.mkdir(parents=True)
    source_file = folder / "LY25112915 报关委托书.pdf"
    source_file.write_bytes(b"%PDF-1.4")

    def fake_pdf_text(filepath, max_pages=2):
        return "报关委托书 一般贸易 报关单编号 123456789012345"

    def prompt_callback(invoice_no):
        raise AssertionError(f"prompt should not be called for {invoice_no}")

    monkeypatch.setattr(fc, "pdf_text", fake_pdf_text)
    processor = fc.TaxRefundProcessor(
        str(output_dir),
        config_path=tmp_path / "materials_config.json",
        ledger_records={"LY25112915": {"发票号": "LY25112915"}},
    )

    record = processor.evaluate_invoice(
        "LY25112915",
        str(folder),
        allow_prompt=False,
        allow_rename=False,
        save_report=False,
        persist_declaration=False,
        prompt_callback=prompt_callback,
    )

    assert record["报关号"] == "123456789012345"
    assert folder.exists()
    assert source_file.exists()
    assert not (output_dir / "declaration_overrides.json").exists()
    assert not list(output_dir.glob("判断报告_*"))
    assert processor.declarations == {}


def test_sync_single_invoice_status_to_company_ledger(tmp_path):
    ledger_path = tmp_path / "ledger.xlsx"
    save_workbook(
        ledger_path,
        {"台账": [["发票号"], ["LY25112915"]]},
    )
    folder = tmp_path / "output" / "LY25112915"
    folder.mkdir(parents=True)
    record = {
        "发票号": "LY25112915",
        "当前文件夹": str(folder),
        "状态": "已改名",
        "缺少材料": "",
        "需人工确认文件": "",
        "闭环状态": "强闭环通过",
        "闭环分数": 100,
        "最后更新时间": "2026-05-22 10:00:00",
    }

    fc.sync_single_invoice_status_to_company_ledger(
        "LY25112915",
        record,
        str(ledger_path),
        batch_name="batch-a",
        create_backup=False,
    )

    assert cell_value_by_header(ledger_path, "台账", "LY25112915", "退税齐套状态") == "齐了"
    assert cell_value_by_header(ledger_path, "台账", "LY25112915", "财务批次") == "batch-a"


def test_company_ledger_sync_manager_batches_and_only_backs_up_once_per_day(tmp_path):
    ledger_path = tmp_path / "ledger.xlsx"
    save_workbook(ledger_path, {"台账": [["发票号"], ["LY25112915"], ["LY25112916"]]})
    folder = tmp_path / "output"
    (folder / "LY25112915").mkdir(parents=True)
    (folder / "LY25112916").mkdir(parents=True)
    manager = fc.CompanyLedgerSyncManager(str(ledger_path), debounce_seconds=30.0)

    for invoice_no in ["LY25112915", "LY25112916"]:
        manager.enqueue(invoice_no, {
            "当前文件夹": str(folder / invoice_no),
            "状态": "已改名",
            "缺少材料": "",
            "需人工确认文件": "",
            "闭环状态": "强闭环通过",
            "最后更新时间": "2026-05-22 10:00:00",
        })
    manager.flush()
    manager.enqueue("LY25112915", {
        "当前文件夹": str(folder / "LY25112915"),
        "状态": "已改名",
        "缺少材料": "",
        "需人工确认文件": "",
        "闭环状态": "强闭环通过",
        "最后更新时间": "2026-05-22 10:05:00",
    })
    manager.flush()

    assert cell_value_by_header(ledger_path, "台账", "LY25112915", "退税齐套状态") == "齐了"
    assert cell_value_by_header(ledger_path, "台账", "LY25112916", "退税齐套状态") == "齐了"
    assert len(list(tmp_path.glob("ledger.xlsx.bak_*"))) == 1


def test_finance_batch_processor_process_batch_creates_package_reports_and_statuses(monkeypatch, tmp_path):
    patch_app_paths(monkeypatch, tmp_path)
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    source_folder = output_dir / "LY25112915"
    source_folder.mkdir()
    source_file = source_folder / "LY25112915 采购合同.jpg"
    source_file.write_text("dummy", encoding="utf-8")

    ledger_path = fc.DEFAULT_COMPANY_LEDGER
    save_workbook(
        ledger_path,
        {
            "台账": [
                ["发票号", "出运日期", "目的国", "金额"],
                ["LY25112915", "2026-05-20", "美国", 100],
                ["LY25112916", "2026-05-20", "美国", 100],
            ]
        },
    )
    batch_path = tmp_path / "财务批次.xlsx"
    save_workbook(
        batch_path,
        {
            "批次": [
                ["发票号"],
                ["LY25112915"],
                ["25112915"],
                ["LY25112916"],
                ["LY99999999"],
            ]
        },
    )

    ledger_records = fc.load_invoice_records(str(ledger_path))
    tax_processor = fc.TaxRefundProcessor(
        str(output_dir),
        config_path=tmp_path / "materials_config.json",
        ledger_records=ledger_records,
    )
    settings = fc.default_app_settings()
    processor = fc.FinanceBatchProcessor(
        str(output_dir),
        str(ledger_path),
        tax_processor,
        settings,
    )

    summary = processor.process_batch(str(batch_path))

    batch_output = Path(summary["batch_output_dir"])
    assert summary["total"] == 3
    assert batch_output.exists()
    copied_folders = [path for path in batch_output.iterdir() if path.is_dir()]
    assert len(copied_folders) == 1
    assert (copied_folders[0] / source_file.name).exists()
    assert source_folder.exists()
    assert source_file.exists()
    assert Path(summary["report_xlsx"]).exists()
    assert Path(summary["report_json"]).exists()
    assert Path(summary["marked_batch_copy"]).exists()

    records = json.loads(Path(summary["report_json"]).read_text(encoding="utf-8"))
    status_by_invoice = {record["发票号"]: record["退税齐套状态"] for record in records}
    assert status_by_invoice["LY25112915"] == "缺材料"
    assert status_by_invoice["LY25112916"] == "未找到文件夹"
    assert status_by_invoice["LY99999999"] == "不在主台账"
    assert cell_value_by_header(ledger_path, "台账", "LY25112915", "退税齐套状态") == "缺材料"
    assert cell_value_by_header(ledger_path, "台账", "LY25112916", "退税齐套状态") == "未找到文件夹"
    assert cell_value_by_header(batch_path, "批次", "LY25112915", "退税齐套状态") == "缺材料"
    assert cell_value_by_header(batch_path, "批次", "LY25112916", "退税齐套状态") == "未找到文件夹"
    assert cell_value_by_header(batch_path, "批次", "LY99999999", "退税齐套状态") == "不在主台账"


def test_file_handler_sync_failure_does_not_affect_processing(monkeypatch, tmp_path):
    source = tmp_path / "incoming.pdf"
    source.write_bytes(b"data")
    dest = tmp_path / "output" / "LY25112915" / "incoming.pdf"
    dest.parent.mkdir(parents=True)
    logs = []

    class FakeExcelLogger:
        def __init__(self):
            self.records = []
            self.saved = 0

        def add_record(self, *args, **kwargs):
            self.records.append((args, kwargs))

        def save(self):
            self.saved += 1

    class FakeTaxProcessor:
        def resolve_target_folder(self, invoice_no):
            return str(dest.parent)

        def process_invoice(self, invoice_no, folder_path, prompt_callback=None):
            return {
                "发票号": invoice_no,
                "当前文件夹": folder_path,
                "状态": "缺材料",
                "缺少材料": "提单",
                "最后更新时间": "2026-05-22 10:00:00",
            }

    class FailingSyncManager:
        def enqueue(self, invoice_no, record):
            raise RuntimeError("boom")

    monkeypatch.setattr(fc, "wait_for_file_ready", lambda *args, **kwargs: (True, ""))
    monkeypatch.setattr(
        fc,
        "move_file_to_output",
        lambda *args, **kwargs: (True, "LY25112915", "moved", str(dest), "LY25112915", "成功"),
    )
    logger = FakeExcelLogger()
    handler = fc.FileHandler(
        output_dir=str(tmp_path / "output"),
        excel_logger=logger,
        invoice_matcher=fc.InvoiceMatcher(["LY25112915"]),
        executor=None,
        log_callback=lambda message, processed=False: logs.append((message, processed)),
        pending_set={str(source)},
        tax_processor=FakeTaxProcessor(),
        ledger_sync_manager=FailingSyncManager(),
    )

    handler._process(str(source))

    assert logger.saved >= 1
    assert str(source) not in handler.pending_set
    assert any("公司主台账状态同步失败" in message for message, _ in logs)
