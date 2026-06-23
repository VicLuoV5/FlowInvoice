import importlib
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def _install_import_stubs():
    if "fitz" not in sys.modules:
        sys.modules["fitz"] = types.SimpleNamespace()

    if "pandas" not in sys.modules:
        sys.modules["pandas"] = types.SimpleNamespace(DataFrame=lambda *a, **k: None, ExcelWriter=None)

    if "rapidocr_onnxruntime" not in sys.modules:
        sys.modules["rapidocr_onnxruntime"] = types.SimpleNamespace(RapidOCR=lambda: None)

    if "openpyxl.utils" not in sys.modules:
        sys.modules["openpyxl"] = types.ModuleType("openpyxl")
        sys.modules["openpyxl.utils"] = types.SimpleNamespace(get_column_letter=lambda idx: chr(64 + idx))
        sys.modules["openpyxl.styles"] = types.SimpleNamespace(
            Font=lambda *a, **k: None,
            PatternFill=lambda *a, **k: None,
            Alignment=lambda *a, **k: None,
            Border=lambda *a, **k: None,
            Side=lambda *a, **k: None,
        )


_install_import_stubs()
processor = importlib.import_module("core.processor")


class ProcessorRegressionTests(unittest.TestCase):
    def test_vat_total_without_tax_pair_requires_manual_tax_review(self):
        with tempfile.TemporaryDirectory() as tmp:
            pdf = Path(tmp) / "vat.pdf"
            pdf.write_bytes(b"%PDF")

            fake_fitz = _TextPdfFitz("发票开票日期2025年01月02日发票号码12345678价税合计100.00")

            with patch.object(processor, "fitz", fake_fitz), patch.object(processor, "RapidOCR", lambda: _FakeOcr("")):
                invoices, failures = processor.extract_invoices_data(tmp)

        self.assertEqual([], failures)
        self.assertEqual(1, len(invoices))
        invoice = invoices[0]
        self.assertEqual("100.00", invoice["价税合计(报销额)"])
        self.assertEqual("⚠️ 需手动核对", invoice["不含税金额"])
        self.assertEqual("⚠️ 需手动核对", invoice["税额"])
        self.assertLess(invoice["置信度(%)"], processor.config.CONFIDENCE_HIGH_THRESHOLD)
        self.assertIn("税额需人工核对", invoice["备注"])

    def test_scanned_pdf_pages_fall_back_to_ocr(self):
        with tempfile.TemporaryDirectory() as tmp:
            pdf = Path(tmp) / "scan.pdf"
            pdf.write_bytes(b"%PDF")

            text = "发票开票日期2025年02月03日发票号码87654321价税合计109.00金额100.00税额9.00"
            fake_fitz = _BlankPdfFitz()

            with patch.object(processor, "fitz", fake_fitz), patch.object(processor, "RapidOCR", lambda: _FakeOcr(text)):
                invoices, failures = processor.extract_invoices_data(tmp)

        self.assertEqual([], failures)
        self.assertEqual(1, len(invoices))
        self.assertEqual("增值税发票", invoices[0]["业务分类"])
        self.assertEqual("109.00", invoices[0]["价税合计(报销额)"])
        self.assertEqual("100.00", invoices[0]["不含税金额"])
        self.assertEqual("9.00", invoices[0]["税额"])


    def test_lodging_vat_invoice_uses_labeled_total_and_long_invoice_number(self):
        source_text = (
            "电子发票（增值税专用发票） 发票号码： 开票日期： "
            "26442000006829024456 2026年06月18日 上海商盛投资管理咨询有限公司 "
            "91310000566529130C 广东简桥酒店有限公司 91440404MAEKFQ2P0W "
            "项目名称 规格型号 单位 数量 单价 金额 税率/征收率 税额 "
            "*生产生活服务*住宿服务 晚 1 359.6415094339623 359.64 6% 21.58 "
            "价税合计（大写） 合计 （小写） 叁佰捌拾壹圆贰角贰分 ¥381.22 "
            "359.64 ¥ 21.58 ¥"
        )
        with tempfile.TemporaryDirectory() as tmp:
            pdf = Path(tmp) / "07.pdf"
            pdf.write_bytes(b"%PDF")
            fake_fitz = _TextPdfFitz(source_text)

            with patch.object(processor, "fitz", fake_fitz), patch.object(processor, "RapidOCR", lambda: _FakeOcr("")):
                invoices, failures = processor.extract_invoices_data(tmp)

        self.assertEqual([], failures)
        self.assertEqual(1, len(invoices))
        invoice = invoices[0]
        self.assertEqual("26442000006829024456", invoice["发票号码"])
        self.assertEqual("住宿发票", invoice["业务分类"])
        self.assertEqual("359.64", invoice["不含税金额"])
        self.assertEqual("21.58", invoice["税额"])
        self.assertEqual("381.22", invoice["价税合计(报销额)"])
        self.assertGreaterEqual(invoice["置信度(%)"], processor.config.CONFIDENCE_HIGH_THRESHOLD)


    def test_transport_invoice_preserves_number_boundaries_for_amounts(self):
        source_text = (
            "电子发票（普通发票） 发票号码：26127000000315704511 开票日期：2026年06月22日 "
            "项目名称 单价 数量 金额 税率/征收率 税额 "
            "*交通运输服务*客运服务 36.86 1 36.86 3% 1.11 "
            "合计 ¥36.86 ¥1.11 价税合计(大写) 叁拾柒圆玖角柒分 (小写)¥37.97"
        )
        with tempfile.TemporaryDirectory() as tmp:
            pdf = Path(tmp) / "01.pdf"
            pdf.write_bytes(b"%PDF")
            fake_fitz = _TextPdfFitz(source_text)

            with patch.object(processor, "fitz", fake_fitz), patch.object(processor, "RapidOCR", lambda: _FakeOcr("")):
                invoices, failures = processor.extract_invoices_data(tmp)

        self.assertEqual([], failures)
        self.assertEqual(1, len(invoices))
        invoice = invoices[0]
        self.assertEqual("26127000000315704511", invoice["发票号码"])
        self.assertEqual("36.86", invoice["不含税金额"])
        self.assertEqual("1.11", invoice["税额"])
        self.assertEqual("37.97", invoice["价税合计(报销额)"])


    def test_transport_invoice_with_adjustment_uses_final_chinese_total(self):
        source_text = (
            "电子发票（普通发票） 旅客运输服务 发票号码： 开票日期： "
            "价税合计（大写） （小写） 备 注 开票人： "
            "26327000001162103916 2026年06月22日 项目名称 数量 单价 金额 税率/征收率 税额 "
            "*交通运输服务*客运服务费 30.747573 1 30.75 3% 0.92 "
            "*交通运输服务*客运服务费 -9.52 3% -0.29 "
            "合计 ¥21.23 ¥0.63 贰拾壹圆捌角陆分 ¥21.86"
        )
        with tempfile.TemporaryDirectory() as tmp:
            pdf = Path(tmp) / "02.pdf"
            pdf.write_bytes(b"%PDF")
            fake_fitz = _TextPdfFitz(source_text)

            with patch.object(processor, "fitz", fake_fitz), patch.object(processor, "RapidOCR", lambda: _FakeOcr("")):
                invoices, failures = processor.extract_invoices_data(tmp)

        self.assertEqual([], failures)
        self.assertEqual(1, len(invoices))
        invoice = invoices[0]
        self.assertEqual("26327000001162103916", invoice["发票号码"])
        self.assertEqual("21.23", invoice["不含税金额"])
        self.assertEqual("0.63", invoice["税额"])
        self.assertEqual("21.86", invoice["价税合计(报销额)"])

    def test_trip_sheet_pdf_is_rejected_as_non_invoice_document(self):
        source_text = (
            "高德地图—打车——行程单 AMAP ITINERARY 申请时间：2026-06-22 "
            "行程时间：2026-06-17 23:23至2026-06-17 23:39 "
            "共计1单行程，合计37.97元 序号 服务商 车型 上车时间 城市 起点 终点 金额"
        )
        with tempfile.TemporaryDirectory() as tmp:
            pdf = Path(tmp) / "19.pdf"
            pdf.write_bytes(b"%PDF")
            fake_fitz = _TextPdfFitz(source_text)

            with patch.object(processor, "fitz", fake_fitz), patch.object(processor, "RapidOCR", lambda: _FakeOcr("")):
                invoices, failures = processor.extract_invoices_data(tmp)

        self.assertEqual([], invoices)
        self.assertEqual(1, len(failures))
        self.assertIn("非发票", failures[0]["reason"])

    def test_hotel_checkout_image_is_rejected_as_non_invoice_document(self):
        source_text = (
            "Home2Suites by Hilton Quanzhou INFORMATION INVOICE 结账单 GuestName "
            "BillNo./账单号 B01001950000069115 Date/日期 Description/项目 Room No/房 "
            "房费 295.00 Balance/余额 ￥0.00 本酒店控制范围 总金额已含税和服务费"
        )
        with tempfile.TemporaryDirectory() as tmp:
            jpg = Path(tmp) / "16.jpg"
            jpg.write_bytes(b"jpg")

            with patch.object(processor, "RapidOCR", lambda: _FakeOcr(source_text)):
                invoices, failures = processor.extract_invoices_data(tmp)

        self.assertEqual([], invoices)
        self.assertEqual(1, len(failures))
        self.assertIn("非发票", failures[0]["reason"])

    def test_merge_skips_corrupt_files_and_saves_valid_pages(self):
        with tempfile.TemporaryDirectory() as tmp:
            good = Path(tmp) / "01_good.pdf"
            bad = Path(tmp) / "02_bad.pdf"
            out = Path(tmp) / "merged.pdf"
            good.write_bytes(b"%PDF")
            bad.write_bytes(b"broken")

            fake_fitz = _MergeFitz(corrupt_name=bad.name)

            with patch.object(processor, "fitz", fake_fitz):
                success, msg = processor.merge_pdfs_logic(tmp, str(out))
                output_exists = out.exists()

        self.assertTrue(success)
        self.assertTrue(output_exists)
        self.assertEqual(1, fake_fitz.merged.page_count)
        self.assertIn("成功合并 1 份文件", msg)
        self.assertIn("1 个文件失败", msg)
        self.assertIn("02_bad.pdf", msg)


class _FakeOcr:
    def __init__(self, text):
        self.text = text

    def __call__(self, _path):
        if not self.text:
            return [], None
        return [[None, self.text, 0.98]], None


class _TextPdfFitz:
    def __init__(self, text):
        self.text = text

    def open(self, *args):
        return _PdfDoc([_TextPage(self.text)])


class _BlankPdfFitz:
    Matrix = lambda self, x, y: (x, y)

    def open(self, *args):
        return _PdfDoc([_BlankPage()])


class _PdfDoc:
    def __init__(self, pages):
        self.pages = pages

    def __iter__(self):
        return iter(self.pages)

    def close(self):
        pass


class _TextPage:
    rect = types.SimpleNamespace(width=100, height=200)
    number = 0

    def __init__(self, text):
        self.text = text

    def get_text(self, _mode):
        return self.text


class _BlankPage(_TextPage):
    def __init__(self):
        super().__init__("")

    def get_pixmap(self, **_kwargs):
        return _Pixmap()


class _Pixmap:
    def save(self, path):
        Path(path).write_bytes(b"png")


class _MergeFitz:
    Rect = lambda self, *args: args

    def __init__(self, corrupt_name):
        self.corrupt_name = corrupt_name
        self.merged = _MergedDoc()

    def open(self, *args):
        if not args:
            return self.merged
        path = Path(args[0])
        if path.name == self.corrupt_name:
            raise RuntimeError("cannot open file")
        return _PdfDoc([_MergePage()])


class _MergedDoc:
    def __init__(self):
        self.page_count = 0

    def new_page(self, width, height):
        self.page_count += 1
        return _OutputPage()

    def save(self, path):
        Path(path).write_bytes(b"%PDF-merged")

    def close(self):
        pass


class _MergePage:
    rect = types.SimpleNamespace(width=100, height=200)
    number = 0


class _OutputPage:
    def show_pdf_page(self, *_args, **_kwargs):
        pass


if __name__ == "__main__":
    unittest.main()
