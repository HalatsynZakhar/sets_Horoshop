from __future__ import annotations

import io
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook

from horoshop_sets import (
    ArticleMappings,
    CatalogIndex,
    PlanItem,
    SetRow,
    Settings,
    StateStore,
    build_excel_template,
    build_mappings_excel,
    build_state_excel,
    import_payload,
    import_results,
    leading_zeros_count,
    load_article_mappings,
    load_settings,
    normalize_article_code,
    parse_article_mappings,
    parse_excel_sets,
    prepare_plan,
    save_article_mappings,
)


class HoroshopSetsTests(unittest.TestCase):
    def excel_bytes(self, rows: list[tuple[object, object, object]]) -> bytes:
        workbook = Workbook()
        worksheet = workbook.active
        for row in rows:
            worksheet.append(row)
        output = io.BytesIO()
        workbook.save(output)
        workbook.close()
        return output.getvalue()

    def test_template_has_expected_columns(self) -> None:
        workbook = load_workbook(io.BytesIO(build_excel_template()), read_only=True)
        self.assertEqual(
            [cell.value for cell in next(workbook.active.iter_rows(max_row=1))],
            [
                "Артикул набору (необов'язково)",
                "Артикули відображення товарів",
                "Кінцева ціна",
                "Знижка %",
                "База розрахунку (Ціна / РРЦ)",
            ],
        )
        self.assertEqual(workbook.active["A2"].number_format, "@")
        self.assertEqual(workbook.active.title, "Набори")
        self.assertEqual(workbook["Інструкція"]["A1"].value, "Шаблон для створення та оновлення наборів")
        self.assertIn("Артикул набору необов'язковий", workbook["Інструкція"]["A4"].value)
        workbook.close()

    def test_excel_parses_price_and_articles(self) -> None:
        rows = parse_excel_sets(
            self.excel_bytes(
                [
                    ("Артикул набора", "Артикулы отображения товаров", "Цена набора"),
                    ("SET-1", "show-a; SHOW-B", "199,50"),
                ]
            )
        )
        self.assertEqual(rows, [SetRow("SET-1", ("show-a", "SHOW-B"), Decimal("199.50"), 2)])

    def test_excel_builds_article_when_it_is_blank(self) -> None:
        rows = parse_excel_sets(
            self.excel_bytes(
                [
                    ("Артикул набору (необов'язково)", "Артикули відображення товарів", "Кінцева ціна"),
                    (None, "show-a; SHOW-B", "199,50"),
                ]
            )
        )
        self.assertEqual(rows[0].article, "show-a.SHOW-B")

    def test_excel_supports_delete_action_without_price_or_products(self) -> None:
        rows = parse_excel_sets(
            self.excel_bytes(
                [
                    ("Артикул набору", "Артикули відображення товарів", "Кінцева ціна", "Видалити (Так)"),
                    ("SET-1", None, None, "Так"),
                ]
            )
        )
        self.assertEqual(rows[0].article, "SET-1")
        self.assertEqual(rows[0].action, "delete")
        self.assertIsNone(rows[0].discounted_price)

    def test_excel_treats_blank_delete_marker_as_upsert(self) -> None:
        rows = parse_excel_sets(
            self.excel_bytes(
                [
                    ("Артикул набору", "Артикули відображення товарів", "Кінцева ціна", "Видалити (Так)"),
                    ("SET-1", "A; B", 100, ""),
                ]
            )
        )
        self.assertEqual(rows[0].action, "upsert")
        self.assertEqual(rows[0].display_articles, ("A", "B"))

    def test_registry_export_has_explicit_delete_marker(self) -> None:
        workbook = load_workbook(
            io.BytesIO(build_state_excel([{"article": "SET-1", "display_articles": ["A", "B"], "discounted_price": "100"}])),
            read_only=True,
        )
        worksheet = workbook.active
        self.assertEqual(worksheet["D1"].value, "Видалити (Так)")
        self.assertIsNone(worksheet["D2"].value)
        workbook.close()

    def test_catalog_resolution_prefers_exact_then_case_insensitive(self) -> None:
        catalog = CatalogIndex.from_raw(
            [
                {"article": "REAL-A", "article_for_display": "Display-A"},
                {"article": "REAL-B", "article_for_display": "DISPLAY-B"},
            ]
        )
        plan = prepare_plan(
            [SetRow("SET-1", ("Display-A", "display-b"), Decimal("12.00"), 2)],
            catalog,
        )
        self.assertTrue(plan[0].ready)
        self.assertEqual(plan[0].products, ("REAL-A", "REAL-B"))

    def test_catalog_index_accepts_raw_api_export(self) -> None:
        catalog = CatalogIndex.from_raw(
            [{"article": "REAL-A", "article_for_display": "DISPLAY-A"}]
        )
        article, error = catalog.resolve_display_article("DISPLAY-A")
        self.assertEqual(article, "REAL-A")
        self.assertEqual(error, "")

    def test_ambiguous_display_article_and_product_article_collision_are_blocked(self) -> None:
        catalog = CatalogIndex.from_raw(
            [
                {"article": "PRODUCT-1", "article_for_display": "Same"},
                {"article": "PRODUCT-2", "article_for_display": "same"},
            ]
        )
        plan = prepare_plan(
            [
                SetRow("SET-1", ("sAmE", "other"), Decimal("12.00"), 2),
                SetRow("PRODUCT-1", ("Same", "other"), Decimal("12.00"), 3),
            ],
            catalog,
        )
        self.assertIn("не є унікальним", plan[0].error)
        self.assertIn("збігається", plan[1].error)

    def test_state_records_composition(self) -> None:
        item = PlanItem("SET-1", ("A", "B"), Decimal("10.00"), ("REAL-A", "REAL-B"), 2)
        with tempfile.TemporaryDirectory() as directory:
            store = StateStore(Path(directory) / "state.json")
            store.record(item, "synced", "Imported")
            store.save()
            reloaded = StateStore(Path(directory) / "state.json")
        self.assertEqual(reloaded.snapshot()[0]["products"], ["REAL-A", "REAL-B"])
        self.assertEqual(reloaded.snapshot()[0]["status"], "synced")

    def test_deletion_removes_active_set_and_keeps_history(self) -> None:
        item = PlanItem("SET-1", ("A", "B"), Decimal("10.00"), ("REAL-A", "REAL-B"), 2)
        with tempfile.TemporaryDirectory() as directory:
            store = StateStore(Path(directory) / "state.json")
            store.record(item, "synced", "Imported")
            store.record_deletion("SET-1", "deleted", "Deleted")
            store.save()
            reloaded = StateStore(Path(directory) / "state.json")
        self.assertEqual(reloaded.snapshot(), [])
        self.assertEqual(reloaded.history[-1]["action"], "delete")

    def test_payload_and_api_log_use_discounted_price_only(self) -> None:
        item = PlanItem("SET-1", ("A", "B"), Decimal("10.50"), ("REAL-A", "REAL-B"), 2)
        settings = Settings(
            domain="https://shop.example.com",
            host="0.0.0.0",
            port=8093,
            currency="UAH",
            title="Разом дешевше",
            batch_size=50,
            request_timeout_seconds=60,
            state_file=Path("state.json"),
        )
        payload = import_payload([item], settings)
        self.assertEqual(payload[0]["discountedPrice"], 10.5)
        self.assertEqual(payload[0]["title"], "Разом дешевше")
        self.assertNotIn("initialPrice", payload[0])
        self.assertEqual(
            import_results(
                {"status": "WARNING", "response": {"log": [{"article": "SET-1", "info": [{"code": 0, "message": "Imported"}]}]}}
            )["SET-1"],
            (True, "Imported"),
        )

    def test_settings_accept_utf8_bom_from_windows_notepad(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            config_path = Path(directory) / "config.json"
            config_path.write_text(
                '{"server":{"port":8093},"horoshop":{"domain":"https://shop.example.com","title":"Вместе дешевле"},"logging":{"public_log_path":"public-logs","public_log_name":"visible.log"}}',
                encoding="utf-8-sig",
            )
            settings = load_settings(config_path)
        self.assertEqual(settings.domain, "https://shop.example.com")
        self.assertEqual(settings.title, "Разом дешевше")
        self.assertEqual(settings.public_log_file, config_path.parent / "public-logs" / "visible.log")

    def test_settings_repair_unescaped_windows_public_log_path(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            config_path = Path(directory) / "config.json"
            config_path.write_text(
                r'{"horoshop":{"domain":"https://shop.example.com"},"logging":{"public_log_path":"C:\ShareFiles\public","public_log_name":"visible.log"}}',
                encoding="utf-8",
            )
            settings = load_settings(config_path)
            repaired_config = config_path.read_text(encoding="utf-8")
        self.assertEqual( settings.public_log_file, Path(r"C:\ShareFiles\public") / "visible.log")
        self.assertIn(r'"public_log_path":"C:\\ShareFiles\\public"', repaired_config)

    def test_calculates_price_from_current_prices_with_discount_and_rounds_to_integer(self) -> None:
        catalog = CatalogIndex.from_raw(
            [
                {"article": "REAL-A", "article_for_display": "A", "price": 250, "price_old": 300},
                {"article": "REAL-B", "article_for_display": "B", "price": 149, "price_old": 0},
            ]
        )
        rows = parse_excel_sets(
            self.excel_bytes(
                [
                    (
                        "Артикул набору (необов'язково)",
                        "Артикули відображення товарів",
                        "Кінцева ціна",
                        "Знижка %",
                        "База розрахунку (Ціна / РРЦ)",
                    ),
                    ("SET-CURR", "A; B", None, 10, "Ціна"),
                ]
            )
        )
        plan = prepare_plan(rows, catalog)
        self.assertTrue(plan[0].ready)
        self.assertEqual(plan[0].base_price_sum, Decimal("399.00"))
        # 399 * 0.9 = 359.1 -> 359
        self.assertEqual(plan[0].discounted_price, Decimal("359"))

    def test_calculates_price_from_rrc_sum_with_discount_and_half_up_rounding(self) -> None:
        catalog = CatalogIndex.from_raw(
            [
                {"article": "REAL-A", "article_for_display": "A", "price": 200, "price_old": 250},
                {"article": "REAL-B", "article_for_display": "B", "price": 180, "price_old": 225},
            ]
        )
        rows = parse_excel_sets(
            self.excel_bytes(
                [
                    (
                        "Артикул набору (необов'язково)",
                        "Артикули відображення товарів",
                        "Кінцева ціна",
                        "Знижка %",
                        "База розрахунку (Ціна / РРЦ)",
                    ),
                    (None, "A; B", None, 5, "РРЦ"),
                ]
            )
        )
        plan = prepare_plan(rows, catalog)
        self.assertTrue(plan[0].ready)
        self.assertEqual(plan[0].article, "A.B")
        self.assertEqual(plan[0].base_price_sum, Decimal("475.00"))
        # 475 * 0.95 = 451.25 -> 451; let's also verify 250 * 0.95 = 237.5 -> 238
        self.assertEqual(plan[0].discounted_price, Decimal("451"))

    def test_rrc_calculation_fails_when_any_product_lacks_old_price(self) -> None:
        catalog = CatalogIndex.from_raw(
            [
                {"article": "REAL-A", "article_for_display": "A", "price": 200, "price_old": 250},
                {"article": "REAL-B", "article_for_display": "B", "price": 180, "price_old": 0},
            ]
        )
        rows = parse_excel_sets(
            self.excel_bytes(
                [
                    ("Артикул набору", "Артикули відображення товарів", "Кінцева ціна"),
                    ("SET-RRC-ERR", "A; B", None),
                ]
            ),
            default_discount_percent=Decimal("15"),
            default_price_base="rrc",
        )
        plan = prepare_plan(rows, catalog)
        self.assertFalse(plan[0].ready)
        self.assertIn("не вказана стара ціна (РРЦ)", plan[0].error)
        self.assertIn("'B'", plan[0].error)

    def test_excel_uses_web_defaults_when_row_price_and_discount_are_blank(self) -> None:
        rows = parse_excel_sets(
            self.excel_bytes(
                [
                    ("Артикул набору", "Артикули відображення товарів", "Кінцева ціна"),
                    ("SET-DEF", "A; B", None),
                    ("SET-FIXED", "A; B", "320"),
                ]
            ),
            default_discount_percent=Decimal("12"),
            default_price_base="rrc",
        )
        self.assertIsNone(rows[0].discounted_price)
        self.assertEqual(rows[0].discount_percent, Decimal("12.00"))
        self.assertEqual(rows[0].price_base, "rrc")
        self.assertEqual(rows[1].discounted_price, Decimal("320.00"))
        self.assertIsNone(rows[1].discount_percent)

    def test_excel_defaults_to_zero_discount_when_price_and_discount_are_blank(self) -> None:
        rows = parse_excel_sets(
            self.excel_bytes(
                [
                    ("Артикул набору", "Артикули відображення товарів", "Кінцева ціна"),
                    ("SET-NO-PRICE", "A; B", None),
                    ("   ", "   ", "   "),
                ]
            )
        )
        self.assertEqual(len(rows), 1)
        self.assertIsNone(rows[0].discounted_price)
        self.assertEqual(rows[0].discount_percent, Decimal("0.00"))
        self.assertEqual(rows[0].price_base, "current")

    def test_excel_handles_percent_cells_negative_discount_and_3col_headers(self) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Артикул набору", "Артикули відображення товарів", "Кінцева ціна"])
        sheet.append(["SET-PCT-FMT", "A; B", None, 0.15, "РРЦ"])
        sheet.cell(row=2, column=4).number_format = "0%"
        sheet.append(["SET-NEG", "A; B", "-10%"])
        output = io.BytesIO()
        workbook.save(output)
        workbook.close()

        rows = parse_excel_sets(output.getvalue())
        self.assertEqual(len(rows), 2)
        self.assertIsNone(rows[0].discounted_price)
        self.assertEqual(rows[0].discount_percent, Decimal("15.00"))
        self.assertEqual(rows[0].price_base, "rrc")
        self.assertIsNone(rows[1].discounted_price)
        self.assertEqual(rows[1].discount_percent, Decimal("10.00"))

    def test_half_up_rounding_on_exact_half(self) -> None:
        catalog = CatalogIndex.from_raw(
            [
                {"article": "REAL-A", "article_for_display": "A", "price": 150},
                {"article": "REAL-B", "article_for_display": "B", "price": 100},
            ]
        )
        # 250 * (100 - 5) / 100 = 237.5 -> 238
        plan = prepare_plan(
            [SetRow("SET-HALF", ("A", "B"), None, 2, discount_percent=Decimal("5"), price_base="current")],
            catalog,
        )
        self.assertTrue(plan[0].ready)
        self.assertEqual(plan[0].discounted_price, Decimal("238"))

    def test_registry_export_preserves_discount_and_price_base_columns(self) -> None:
        data = build_state_excel(
            [
                {
                    "article": "SET-1",
                    "display_articles": ["A", "B"],
                    "discounted_price": "359",
                    "discount_percent": "10",
                    "price_base": "rrc",
                }
            ]
        )
        workbook = load_workbook(io.BytesIO(data), read_only=True)
        worksheet = workbook.active
        self.assertEqual(worksheet["E1"].value, "Знижка %")
        self.assertEqual(worksheet["F1"].value, "База розрахунку (Ціна / РРЦ)")
        self.assertEqual(worksheet["E2"].value, "10")
        self.assertEqual(worksheet["F2"].value, "РРЦ")
        workbook.close()

    def test_mappings_generate_cartesian_set_variations_with_site_display_articles(self) -> None:
        catalog = CatalogIndex.from_raw(
            [
                {"article": "REAL-A1", "article_for_display": "SITE-A1", "price": 200, "price_old": 250},
                {"article": "REAL-A2", "article_for_display": "SITE-A2", "price": 300, "price_old": 350},
                {"article": "REAL-B1", "article_for_display": "SITE-B1", "price": 100, "price_old": 120},
                {"article": "REAL-B2", "article_for_display": "SITE-B2", "price": 150, "price_old": 180},
            ]
        )
        mappings = parse_article_mappings(
            self.excel_bytes(
                [
                    ("Артикул", "Варіант 1", "Варіант 2"),
                    ("SELLER-A", "SITE-A1", "SITE-A2"),
                    ("SELLER-B", "SITE-B1", "SITE-B2"),
                ]
            )
        )
        rows = parse_excel_sets(
            self.excel_bytes(
                [
                    (
                        "Артикул набору (необов'язково)",
                        "Артикули відображення товарів",
                        "Кінцева ціна",
                        "Знижка %",
                        "База розрахунку (Ціна / РРЦ)",
                    ),
                    (None, "SELLER-A; SELLER-B", None, 10, "Ціна"),
                ]
            )
        )
        plan = prepare_plan(rows, catalog, normalize_articles=True, mappings=mappings)
        self.assertEqual(len(plan), 4)
        self.assertTrue(all(item.ready for item in plan))
        self.assertEqual(
            [item.article for item in plan],
            ["SITE-A1.SITE-B1", "SITE-A1.SITE-B2", "SITE-A2.SITE-B1", "SITE-A2.SITE-B2"],
        )
        self.assertEqual(
            [item.display_articles for item in plan],
            [
                ("SITE-A1", "SITE-B1"),
                ("SITE-A1", "SITE-B2"),
                ("SITE-A2", "SITE-B1"),
                ("SITE-A2", "SITE-B2"),
            ],
        )
        self.assertEqual(
            [item.products for item in plan],
            [
                ("REAL-A1", "REAL-B1"),
                ("REAL-A1", "REAL-B2"),
                ("REAL-A2", "REAL-B1"),
                ("REAL-A2", "REAL-B2"),
            ],
        )
        # Sums: 300->270, 350->315, 400->360, 450->405
        self.assertEqual(
            [item.discounted_price for item in plan],
            [Decimal("270"), Decimal("315"), Decimal("360"), Decimal("405")],
        )

    def test_mappings_support_internal_horoshop_articles_and_normalization(self) -> None:
        catalog = CatalogIndex.from_raw(
            [
                {"article": "TECH/1 (A)", "article_for_display": "Visible-1", "price": 500},
                {"article": "TECH-2", "article_for_display": "Display-2", "price": 300},
            ]
        )
        mappings = parse_article_mappings(
            self.excel_bytes(
                [
                    ("Постачальник", "Технічний"),
                    ("SUP-1", "TECH/1 (A)"),
                ]
            )
        )
        row = SetRow("SUP-1.Display.2", ("SUP-1", "Display.2"), Decimal("700"), 2)
        plan = prepare_plan([row], catalog, normalize_articles=True, mappings=mappings)
        self.assertEqual(len(plan), 1)
        self.assertTrue(plan[0].ready)
        self.assertEqual(plan[0].article, "Visible-1.Display-2")
        self.assertEqual(plan[0].display_articles, ("Visible-1", "Display-2"))
        self.assertEqual(plan[0].products, ("TECH/1 (A)", "TECH-2"))

    def test_reimporting_exported_registry_with_cluster_items_does_not_conflict(self) -> None:
        catalog = CatalogIndex.from_raw(
            [
                {"article": "REAL-A1", "article_for_display": "SITE-A1", "price": 200},
                {"article": "REAL-A2", "article_for_display": "SITE-A2", "price": 300},
                {"article": "REAL-B", "article_for_display": "SITE-B", "price": 100},
            ]
        )
        mappings = ArticleMappings(
            {
                "SITE-A1": ("SITE-A1", "SITE-A2"),
                "SITE-A2": ("SITE-A2", "SITE-A1"),
            }
        )
        rows = [
            SetRow("SITE-A1.SITE-B", ("SITE-A1", "SITE-B"), Decimal("250"), 2),
            SetRow("SITE-A2.SITE-B", ("SITE-A2", "SITE-B"), Decimal("350"), 3),
        ]
        plan = prepare_plan(rows, catalog, normalize_articles=True, mappings=mappings)
        self.assertEqual(len(plan), 2)
        self.assertTrue(all(item.ready for item in plan))
        self.assertEqual([item.article for item in plan], ["SITE-A1.SITE-B", "SITE-A2.SITE-B"])
        self.assertEqual([item.discounted_price for item in plan], [Decimal("250"), Decimal("350")])

    def test_leading_zeros_preserved_and_protected_in_mappings_and_catalog(self) -> None:
        self.assertEqual(normalize_article_code("04040"), "04040")
        self.assertEqual(normalize_article_code("4040"), "4040")
        self.assertEqual(leading_zeros_count("04040"), 1)
        self.assertEqual(leading_zeros_count("4040"), 0)

        with self.assertRaisesRegex(ValueError, "відрізняються лише початковими нулями"):
            parse_article_mappings(
                self.excel_bytes(
                    [
                        ("Варіант 1", "Варіант 2"),
                        ("04040", "4040"),
                    ]
                )
            )

        catalog = CatalogIndex.from_raw(
            [
                {"article": "INT-04040", "article_for_display": "04040", "price": 100},
            ]
        )
        prod_0, _ = catalog.resolve_product("04040", use_normalized=True)
        prod_no_0, err = catalog.resolve_product("4040", use_normalized=True)
        self.assertIsNotNone(prod_0)
        self.assertIsNone(prod_no_0)
        self.assertIn("не знайдений у каталозі", err)

    def test_mappings_excel_roundtrip_and_server_clear_backup(self) -> None:
        import asyncio
        import sets_server

        mappings = ArticleMappings(
            {
                "04040": ("04040", "UUID-04040"),
                "UUID-04040": ("UUID-04040", "04040"),
            }
        )
        excel_bytes = build_mappings_excel(mappings)
        restored = parse_article_mappings(excel_bytes)
        self.assertEqual(restored.targets_for("04040", False), ("04040", "UUID-04040"))

        with tempfile.TemporaryDirectory() as temp_dir:
            test_file = Path(temp_dir) / "data" / "article_mappings.json"
            save_article_mappings(test_file, mappings)
            original_file = sets_server.MAPPINGS_FILE
            try:
                sets_server.MAPPINGS_FILE = test_file
                info_before = sets_server.mappings_info()
                self.assertEqual(info_before["sources"], 2)
                self.assertEqual(info_before["rules"], 4)

                export_res = sets_server.export_mappings()
                self.assertEqual(
                    export_res.media_type,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                self.assertTrue(len(export_res.body) > 0)

                clear_res = asyncio.run(sets_server.clear_mappings())
                self.assertEqual(clear_res["status"], "cleared")
                self.assertTrue(clear_res["backup"].endswith(".bak"))

                info_after = sets_server.mappings_info()
                self.assertEqual(info_after["sources"], 0)
                self.assertEqual(info_after["rules"], 0)

                backup_path = test_file.parent / clear_res["backup"]
                self.assertTrue(backup_path.exists())
                backup_mappings = load_article_mappings(backup_path)
                self.assertEqual(backup_mappings.targets_for("04040", False), ("04040", "UUID-04040"))
            finally:
                sets_server.MAPPINGS_FILE = original_file


if __name__ == "__main__":
    unittest.main()
