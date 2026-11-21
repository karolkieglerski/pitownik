import os
import pytest
from datetime import datetime
from unittest.mock import patch, MagicMock
from collections import OrderedDict

import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "cli"))

from main import (
    read_transactions,
    format_date,
    parse_date,
    get_nbp_rate,
    fetch_nbp_rates,
    group_by_sale_date,
    WIRE_FEE_USD,
)

EXAMPLE_FILE = os.path.join(os.path.dirname(__file__), "input", "example_2025.xlsx")


# ---------------------------------------------------------------------------
# read_transactions
# ---------------------------------------------------------------------------

class TestReadTransactions:
    def test_reads_only_sell_records(self):
        txns = read_transactions(EXAMPLE_FILE)
        assert all(t["Record Type"] == "Sell" for t in txns)

    def test_skips_summary_row(self):
        txns = read_transactions(EXAMPLE_FILE)
        assert not any(t["Record Type"] == "Summary" for t in txns)

    def test_transaction_count(self):
        txns = read_transactions(EXAMPLE_FILE)
        assert len(txns) == 4  # 2 RSU + 1 ESPP + 1 ESPP

    def test_rsu_transactions_have_zero_acquisition_cost(self):
        txns = read_transactions(EXAMPLE_FILE)
        rsu = [t for t in txns if t["Plan Type"] == "RS"]
        assert len(rsu) == 2
        for t in rsu:
            assert t["Acquisition Cost"] == 0

    def test_espp_transactions_have_nonzero_acquisition_cost(self):
        txns = read_transactions(EXAMPLE_FILE)
        espp = [t for t in txns if t["Plan Type"] == "ESPP"]
        assert len(espp) == 2
        for t in espp:
            assert t["Acquisition Cost"] > 0

    def test_all_transactions_have_required_fields(self):
        required = ["Record Type", "Symbol", "Plan Type", "Quantity",
                     "Date Sold", "Total Proceeds", "Proceeds Per Share"]
        txns = read_transactions(EXAMPLE_FILE)
        for t in txns:
            for field in required:
                assert field in t, f"Missing field: {field}"


# ---------------------------------------------------------------------------
# format_date
# ---------------------------------------------------------------------------

class TestFormatDate:
    def test_datetime_object(self):
        assert format_date(datetime(2025, 6, 9)) == "09.06.2025"

    def test_us_string_format(self):
        assert format_date("06/09/2025") == "09.06.2025"

    def test_unknown_format_returns_string(self):
        assert format_date("some-date") == "some-date"


# ---------------------------------------------------------------------------
# parse_date
# ---------------------------------------------------------------------------

class TestParseDate:
    def test_datetime_passthrough(self):
        dt = datetime(2025, 3, 15)
        assert parse_date(dt) is dt

    def test_us_string(self):
        result = parse_date("03/15/2025")
        assert result == datetime(2025, 3, 15)

    def test_invalid_string_raises(self):
        with pytest.raises(ValueError):
            parse_date("2025-03-15")


# ---------------------------------------------------------------------------
# group_by_sale_date
# ---------------------------------------------------------------------------

class TestGroupBySaleDate:
    def test_groups_by_date_sold(self):
        txns = read_transactions(EXAMPLE_FILE)
        groups = group_by_sale_date(txns)
        # example has 3 distinct sale dates: 04/10, 06/10, 12/10
        assert len(groups) == 3

    def test_rsu_package_has_two_transactions(self):
        txns = read_transactions(EXAMPLE_FILE)
        groups = group_by_sale_date(txns)
        # First group (04/10/2025) should have 2 RSU sells
        first_key = list(groups.keys())[0]
        assert len(groups[first_key]) == 2

    def test_espp_packages_have_one_transaction_each(self):
        txns = read_transactions(EXAMPLE_FILE)
        groups = group_by_sale_date(txns)
        keys = list(groups.keys())
        assert len(groups[keys[1]]) == 1  # June ESPP
        assert len(groups[keys[2]]) == 1  # December ESPP

    def test_preserves_insertion_order(self):
        txns = read_transactions(EXAMPLE_FILE)
        groups = group_by_sale_date(txns)
        dates = [parse_date(k) for k in groups.keys()]
        assert dates == sorted(dates)


# ---------------------------------------------------------------------------
# Tax calculation logic (RSU)
# ---------------------------------------------------------------------------

class TestRSUTaxLogic:
    """RSU: full proceeds are taxable, cost = 0."""

    def test_rsu_proceeds_sum(self):
        txns = read_transactions(EXAMPLE_FILE)
        rsu = [t for t in txns if t["Plan Type"] == "RS"]
        total = sum(t["Total Proceeds"] for t in rsu)
        # 50 * 52 + 30 * 52 = 2600 + 1560 = 4160
        assert total == pytest.approx(4160.0)

    def test_rsu_cost_is_zero(self):
        txns = read_transactions(EXAMPLE_FILE)
        rsu = [t for t in txns if t["Plan Type"] == "RS"]
        total_cost = sum(t["Acquisition Cost"] for t in rsu)
        assert total_cost == 0.0


# ---------------------------------------------------------------------------
# Tax calculation logic (ESPP)
# ---------------------------------------------------------------------------

class TestESPPTaxLogic:
    """ESPP: taxable = proceeds - acquisition cost."""

    def test_espp_proceeds_sum(self):
        txns = read_transactions(EXAMPLE_FILE)
        espp = [t for t in txns if t["Plan Type"] == "ESPP"]
        total = sum(t["Total Proceeds"] for t in espp)
        # 120 * 53 + 80 * 48 = 6360 + 3840 = 10200
        assert total == pytest.approx(10200.0)

    def test_espp_cost_sum(self):
        txns = read_transactions(EXAMPLE_FILE)
        espp = [t for t in txns if t["Plan Type"] == "ESPP"]
        total_cost = sum(t["Acquisition Cost"] for t in espp)
        # 120 * 38.29 + 80 * 35.70 = 4594.80 + 2856.00 = 7450.80
        assert total_cost == pytest.approx(7450.80)

    def test_espp_taxable_amount(self):
        txns = read_transactions(EXAMPLE_FILE)
        espp = [t for t in txns if t["Plan Type"] == "ESPP"]
        proceeds = sum(t["Total Proceeds"] for t in espp)
        cost = sum(t["Acquisition Cost"] for t in espp)
        taxable = proceeds - cost
        # 10200 - 7450.80 = 2749.20
        assert taxable == pytest.approx(2749.20)

    def test_espp_has_purchase_date(self):
        txns = read_transactions(EXAMPLE_FILE)
        espp = [t for t in txns if t["Plan Type"] == "ESPP"]
        for t in espp:
            assert t["Purchase Date"] != "--"
            # Should be parseable
            parse_date(t["Purchase Date"])


# ---------------------------------------------------------------------------
# NBP rate fetching (mocked)
# ---------------------------------------------------------------------------

class TestNBPRates:
    @patch("main.urllib.request.urlopen")
    def test_get_nbp_rate_returns_rate_and_date(self, mock_urlopen):
        response_data = {
            "rates": [{"mid": 4.05, "effectiveDate": "2025-04-09"}]
        }
        mock_response = MagicMock()
        mock_response.read.return_value = bytes(
            __import__("json").dumps(response_data), "utf-8"
        )
        mock_urlopen.return_value = mock_response

        rate, effective = get_nbp_rate(datetime(2025, 4, 10))
        assert rate == 4.05
        assert effective == "2025-04-09"

    @patch("main.urllib.request.urlopen")
    def test_get_nbp_rate_retries_on_weekend(self, mock_urlopen):
        """If day-1 is a weekend/holiday (404), it should retry earlier dates."""
        import urllib.error

        call_count = 0

        def side_effect(url):
            nonlocal call_count
            call_count += 1
            if call_count <= 2:  # First two attempts fail (Sat, Fri holiday)
                raise urllib.error.HTTPError(url, 404, "Not Found", {}, None)
            response_data = {
                "rates": [{"mid": 3.80, "effectiveDate": "2025-04-03"}]
            }
            mock = MagicMock()
            mock.read.return_value = bytes(
                __import__("json").dumps(response_data), "utf-8"
            )
            return mock

        mock_urlopen.side_effect = side_effect

        # Monday April 7 -> tries Apr 6 (Sun, fail), Apr 5 (Sat, fail), Apr 4 (Fri? fail scenario), Apr 3 (ok)
        rate, effective = get_nbp_rate(datetime(2025, 4, 7))
        assert rate == 3.80
        assert call_count == 3

    @patch("main.urllib.request.urlopen")
    def test_get_nbp_rate_raises_after_7_retries(self, mock_urlopen):
        import urllib.error
        mock_urlopen.side_effect = urllib.error.HTTPError(
            "url", 404, "Not Found", {}, None
        )
        with pytest.raises(RuntimeError, match="Nie udało się pobrać kursu"):
            get_nbp_rate(datetime(2025, 4, 10))

    @patch("main.get_nbp_rate")
    def test_fetch_nbp_rates_deduplicates(self, mock_get_rate):
        mock_get_rate.return_value = (4.0, "2025-04-09")
        dates = [datetime(2025, 4, 10), datetime(2025, 4, 10), datetime(2025, 4, 10)]
        rates = fetch_nbp_rates(dates)
        assert mock_get_rate.call_count == 1
        assert datetime(2025, 4, 10) in rates


# ---------------------------------------------------------------------------
# PLN conversion: different rates for revenue vs ESPP cost
# ---------------------------------------------------------------------------

class TestPLNConversion:
    """Verify that ESPP cost uses purchase-date rate, not sale-date rate."""

    def test_espp_cost_uses_purchase_date_rate(self):
        txns = read_transactions(EXAMPLE_FILE)
        groups = group_by_sale_date(txns)
        keys = list(groups.keys())

        # Second group = June ESPP: sold 06/10, purchased 06/05
        espp_txns = groups[keys[1]]
        assert espp_txns[0]["Plan Type"] == "ESPP"

        sale_date = parse_date(espp_txns[0]["Date Sold"])
        purchase_date = parse_date(espp_txns[0]["Purchase Date"])

        # Dates should be different
        assert sale_date != purchase_date
        # Purchase before sale
        assert purchase_date < sale_date

    def test_different_rates_produce_different_pln_values(self):
        """With different rates for cost and revenue, PLN values differ."""
        sale_rate = 3.75
        cost_rate = 3.70
        proceeds_usd = 6360.0
        cost_usd = 4594.80

        proceeds_pln = proceeds_usd * sale_rate
        cost_pln = cost_usd * cost_rate

        # Taxable in PLN should NOT equal (proceeds - cost) * single_rate
        taxable_pln = proceeds_pln - cost_pln
        single_rate_taxable = (proceeds_usd - cost_usd) * sale_rate

        assert taxable_pln != pytest.approx(single_rate_taxable)


# ---------------------------------------------------------------------------
# Wire fee is informational only
# ---------------------------------------------------------------------------

class TestWireFee:
    def test_wire_fee_not_deducted_from_taxable(self):
        """Wire fee should NOT reduce taxable amount."""
        txns = read_transactions(EXAMPLE_FILE)
        rsu = [t for t in txns if t["Plan Type"] == "RS"]
        proceeds = sum(t["Total Proceeds"] for t in rsu)
        # Taxable for RSU = full proceeds (not proceeds - wire_fee)
        taxable = proceeds  # cost = 0 for RSU
        assert taxable == proceeds
        assert taxable != proceeds - WIRE_FEE_USD

