#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import unittest
from decimal import Decimal

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "skills", "dealer-ops-suite", "scripts"))

from settlement_rules import (  # noqa: E402
    RateBook,
    RuleViolation,
    calculate_consignment,
    calculate_monthly_buyout,
    normalize_business_type,
)


D = Decimal


def prices(rows=None):
    return RateBook(rows or [
        {"barcode": "690000000001", "value": "50", "start": "2026-10-01", "end": "2026-10-15", "source": "合同附件v1"},
        {"barcode": "690000000001", "value": "55", "start": "2026-10-16", "end": None, "source": "调价政策v2"},
    ])


def decals(rows=None):
    return RateBook(rows or [
        {"barcode": "690000000001", "value": "2", "start": "2026-10-01", "end": "2026-10-15"},
        {"barcode": "690000000001", "value": "3", "start": "2026-10-16", "end": None},
    ], kind="贴花标准")


class ConsignmentRulesTest(unittest.TestCase):
    def setUp(self):
        self.sales = [
            {"id": "S1", "barcode": "690000000001", "date": "2026-10-10", "quantity": 40},
            {"id": "S2", "barcode": "690000000001", "date": "2026-10-20", "quantity": 20},
        ]
        self.kwargs = dict(
            period_start="2026-10-01", period_end="2026-10-31",
            opening={"690000000001": 100}, inventory={"690000000001": 37},
            addback={}, sales=self.sales, prices=prices(), decals=decals(),
        )

    def test_mid_month_price_is_applied_per_sale_date(self):
        result = calculate_consignment(**self.kwargs)
        self.assertEqual(result["gross_sales_amount"], D("3100.00"))
        self.assertEqual(result["settlement_amount"], D("3100.00"))
        self.assertEqual([x["unit_price"] for x in result["priced_lines"]], [D("50.00"), D("55.00")])

    def test_shortage_is_listed_at_closing_price_but_not_settled(self):
        result = calculate_consignment(**self.kwargs)
        row = result["quantity_reconciliation"][0]
        self.assertEqual(row["theoretical_end_qty"], D("40"))
        self.assertEqual(row["variance_qty"], D("3"))
        self.assertEqual(row["variance_display_amount"], D("165.00"))
        self.assertEqual(row["status"], "短缺·待双方核对")
        self.assertFalse(result["variance_included"])
        self.assertEqual(result["settlement_amount"], D("3100.00"))

    def test_single_period_authorization_is_required_and_not_inherited(self):
        kwargs = dict(self.kwargs)
        kwargs["variance_authorization"] = {
            "approved": True, "period_start": "2026-09-01", "period_end": "2026-09-30",
        }
        with self.assertRaises(RuleViolation) as caught:
            calculate_consignment(**kwargs)
        self.assertEqual(caught.exception.code, "VARIANCE_AUTHORIZATION_PERIOD_MISMATCH")

        kwargs["variance_authorization"].update(period_start="2026-10-01", period_end="2026-10-31")
        result = calculate_consignment(**kwargs)
        self.assertTrue(result["variance_included"])
        self.assertEqual(result["settlement_amount"], D("3265.00"))

        next_period = dict(self.kwargs)
        next_period.update(period_start="2026-11-01", period_end="2026-11-30", sales=[
            {"id": "N1", "barcode": "690000000001", "date": "2026-11-02", "quantity": 60},
        ], inventory={"690000000001": 37})
        next_result = calculate_consignment(**next_period)
        self.assertFalse(next_result["variance_included"])
        self.assertEqual(next_result["settlement_amount"], D("3300.00"))

    def test_surplus_has_negative_display_amount(self):
        kwargs = dict(self.kwargs, inventory={"690000000001": 45})
        result = calculate_consignment(**kwargs)
        row = result["quantity_reconciliation"][0]
        self.assertEqual(row["variance_qty"], D("-5"))
        self.assertEqual(row["variance_display_amount"], D("-275.00"))
        self.assertEqual(row["status"], "盘盈·待双方核对")

    def test_no_threshold_blocks_large_variance(self):
        kwargs = dict(self.kwargs, inventory={"690000000001": 0})
        result = calculate_consignment(**kwargs)
        self.assertEqual(result["quantity_reconciliation"][0]["variance_qty"], D("40"))
        self.assertEqual(result["settlement_amount"], D("3100.00"))

    def test_return_uses_original_sale_price_and_original_decal(self):
        sales = self.sales + [{
            "id": "R1", "kind": "return", "barcode": "690000000001", "date": "2026-10-25",
            "quantity": 5, "original_sale_id": "S1", "original_decal_charged": True,
        }]
        result = calculate_consignment(**dict(self.kwargs, sales=sales, inventory={"690000000001": 42}))
        self.assertEqual(result["return_amount"], D("-250.00"))
        self.assertEqual(result["sales_settlement_amount"], D("2850.00"))
        self.assertEqual(result["decal_sale_amount"], D("180.00"))
        self.assertEqual(result["decal_return_amount"], D("-10.00"))
        self.assertEqual(result["decal_amount"], D("170.00"))
        self.assertEqual(result["quantity_reconciliation"][0]["net_sales_qty"], D("55"))

    def test_return_without_original_price_evidence_is_blocked(self):
        sales = self.sales + [{
            "id": "R1", "kind": "return", "barcode": "690000000001", "date": "2026-10-25", "quantity": 5,
        }]
        with self.assertRaises(RuleViolation) as caught:
            calculate_consignment(**dict(self.kwargs, sales=sales))
        self.assertEqual(caught.exception.code, "ORIGINAL_PRICE_REQUIRED")

    def test_return_without_original_decal_evidence_is_blocked(self):
        sales = self.sales + [{
            "id": "R1", "kind": "return", "barcode": "690000000001", "date": "2026-10-25",
            "quantity": 5, "original_sale_id": "S1",
        }]
        with self.assertRaises(RuleViolation) as caught:
            calculate_consignment(**dict(self.kwargs, sales=sales))
        self.assertEqual(caught.exception.code, "ORIGINAL_DECAL_UNKNOWN")

    def test_string_false_does_not_reverse_decal(self):
        sales = self.sales + [{"id": "R1", "kind": "return", "barcode": "690000000001",
                              "date": "2026-10-25", "quantity": 5, "original_sale_id": "S1",
                              "original_decal_charged": "否"}]
        result = calculate_consignment(**dict(self.kwargs, sales=sales, inventory={"690000000001": 42}))
        self.assertEqual(result["decal_return_amount"], D("0.00"))
        self.assertEqual(result["decal_amount"], D("180.00"))

    def test_original_sale_barcode_mismatch_is_blocked(self):
        returns = [{"id": "R1", "barcode": "690000000001", "date": "2026-10-26", "quantity": 5,
                    "original_receipt_id": "P1", "original_sale_id": "OTHER", "original_decal_charged": True}]
        original_sales = [{"id": "OTHER", "barcode": "690000000002", "date": "2026-10-12",
                           "quantity": 5, "unit_price": 50}]
        kwargs = dict(
            period_start="2026-10-01", period_end="2026-10-31",
            receipts=[{"id": "P1", "barcode": "690000000001", "date": "2026-10-10", "quantity": 40}],
            returns=returns, sales=[], original_sales=original_sales, prices=prices(), decals=decals(),
            buyout_date="2026-11-05", payment_due_date="2026-11-15",
        )
        with self.assertRaises(RuleViolation) as caught:
            calculate_monthly_buyout(**kwargs)
        self.assertEqual(caught.exception.code, "ORIGINAL_BARCODE_MISMATCH")

    def test_no_forced_previous_period_closing_reconciliation(self):
        result = calculate_consignment(**self.kwargs)
        self.assertIn("不强制与上期台账期末衔接", result["opening_source"])
        self.assertNotIn("previous_closing_qty", result)


class RateBookTest(unittest.TestCase):
    def test_gap_is_blocked_on_transaction_date(self):
        book = RateBook([{"barcode": "690000000001", "value": 50, "start": "2026-10-01", "end": "2026-10-15"}])
        with self.assertRaises(RuleViolation) as caught:
            book.rate_on("690000000001", "2026-10-16")
        self.assertEqual(caught.exception.code, "NO_UNIQUE_RATE")

    def test_period_gap_is_blocked_even_without_a_transaction_in_gap(self):
        book = RateBook([
            {"barcode": "690000000001", "value": 50, "start": "2026-10-01", "end": "2026-10-15"},
            {"barcode": "690000000001", "value": 55, "start": "2026-10-20", "end": None},
        ])
        with self.assertRaises(RuleViolation) as caught:
            book.validate_period("690000000001", "2026-10-01", "2026-10-31")
        self.assertEqual(caught.exception.code, "RATE_COVERAGE_GAP")

    def test_overlap_is_blocked_at_book_construction(self):
        with self.assertRaises(RuleViolation) as caught:
            RateBook([
                {"barcode": "690000000001", "value": 50, "start": "2026-10-01", "end": "2026-10-20"},
                {"barcode": "690000000001", "value": 55, "start": "2026-10-16", "end": None},
            ])
        self.assertEqual(caught.exception.code, "OVERLAPPING_RATE_INTERVALS")

    def test_missing_transaction_date_is_blocked(self):
        kwargs = dict(
            period_start="2026-10-01", period_end="2026-10-31", opening={}, inventory={}, addback={},
            sales=[{"id": "S1", "barcode": "690000000001", "quantity": 1}], prices=prices(),
        )
        with self.assertRaises(RuleViolation) as caught:
            calculate_consignment(**kwargs)
        self.assertEqual(caught.exception.code, "MISSING_BUSINESS_DATE")


class MonthlyBuyoutRulesTest(unittest.TestCase):
    def setUp(self):
        self.kwargs = dict(
            period_start="2026-10-01", period_end="2026-10-31",
            receipts=[
                {"id": "P1", "barcode": "690000000001", "date": "2026-10-10", "quantity": 40},
                {"id": "P2", "barcode": "690000000001", "date": "2026-10-20", "quantity": 20},
            ],
            returns=[], sales=[
                {"id": "S1", "barcode": "690000000001", "date": "2026-10-12", "quantity": 8},
                {"id": "S2", "barcode": "690000000001", "date": "2026-10-22", "quantity": 7},
            ],
            prices=prices(), decals=decals(), buyout_date="2026-11-05", payment_due_date="2026-11-15",
        )

    def test_buyout_uses_net_receipts_not_sales(self):
        result = calculate_monthly_buyout(**self.kwargs)
        self.assertEqual(result["buyout_qty"], D("60"))
        self.assertEqual(result["buyout_amount"], D("3100.00"))
        self.assertEqual(result["sales_qty_for_decal_only"], D("15"))
        self.assertEqual(result["decal_amount"], D("37.00"))
        self.assertEqual(result["buyout_date"], "2026-11-05")
        self.assertEqual(result["payment_due_date"], "2026-11-15")

    def test_sales_change_only_changes_decal(self):
        first = calculate_monthly_buyout(**self.kwargs)
        kwargs = dict(self.kwargs, sales=[])
        second = calculate_monthly_buyout(**kwargs)
        self.assertEqual(second["buyout_amount"], first["buyout_amount"])
        self.assertEqual(second["buyout_qty"], first["buyout_qty"])
        self.assertEqual(second["decal_amount"], D("0.00"))

    def test_return_reduces_receipt_qty_at_original_receipt_price(self):
        returns = [{
            "id": "R1", "barcode": "690000000001", "date": "2026-10-26", "quantity": 5,
            "original_receipt_id": "P1", "original_sale_id": "S1", "original_decal_charged": True,
        }]
        result = calculate_monthly_buyout(**dict(self.kwargs, returns=returns))
        self.assertEqual(result["buyout_qty"], D("55"))
        self.assertEqual(result["return_amount"], D("-250.00"))
        self.assertEqual(result["buyout_amount"], D("2850.00"))
        self.assertEqual(result["decal_return_amount"], D("-10.00"))
        self.assertEqual(result["decal_amount"], D("27.00"))

    def test_buyout_and_payment_dates_are_not_merged(self):
        result = calculate_monthly_buyout(**self.kwargs)
        self.assertNotEqual(result["buyout_date"], result["payment_due_date"])


class AmountRoundingTest(unittest.TestCase):
    def test_consignment_total_sums_individually_rounded_lines(self):
        rate_rows = [{"barcode": "690000000001", "value": "0.005",
                      "start": "2026-10-01", "end": None, "source": "测试价"}]
        sales = [{"id": f"S{index}", "barcode": "690000000001", "date": "2026-10-10", "quantity": 1}
                 for index in range(4)]
        result = calculate_consignment(
            period_start="2026-10-01", period_end="2026-10-31",
            opening={"690000000001": 4}, inventory={"690000000001": 0}, addback={},
            sales=sales, prices=prices(rate_rows), decals=decals([dict(rate_rows[0], value="0")]),
        )
        self.assertTrue(all(line["amount"] == D("0.01") for line in result["priced_lines"]))
        self.assertEqual(sum((line["amount"] for line in result["priced_lines"]), D("0")), D("0.04"))
        self.assertEqual(result["settlement_amount"], D("0.04"))

    def test_monthly_total_sums_individually_rounded_lines(self):
        rate_rows = [{"barcode": "690000000001", "value": "0.005",
                      "start": "2026-10-01", "end": None, "source": "测试价"}]
        receipts = [{"id": f"P{index}", "barcode": "690000000001", "date": "2026-10-10", "quantity": 1}
                    for index in range(4)]
        result = calculate_monthly_buyout(
            period_start="2026-10-01", period_end="2026-10-31", receipts=receipts,
            returns=[], sales=[], prices=prices(rate_rows), decals=None,
            buyout_date="2026-11-05", payment_due_date="2026-11-15",
        )
        self.assertTrue(all(line["amount"] == D("0.01") for line in result["priced_receipts"]))
        self.assertEqual(sum((line["amount"] for line in result["priced_receipts"]), D("0")), D("0.04"))
        self.assertEqual(result["buyout_amount"], D("0.04"))


class RoutingTest(unittest.TestCase):
    def test_direct_buyout_does_not_enter_reconciliation(self):
        with self.assertRaises(RuleViolation) as caught:
            normalize_business_type("直接买断")
        self.assertEqual(caught.exception.code, "DIRECT_BUYOUT_NOT_RECONCILED")

    def test_explicit_modes(self):
        self.assertEqual(normalize_business_type("铺货代销"), "consignment")
        self.assertEqual(normalize_business_type("月结买断"), "monthly_buyout")


if __name__ == "__main__":
    unittest.main()
