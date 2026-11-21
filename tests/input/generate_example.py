"""Generate an anonymized example E*TRADE Gains & Losses xlsx file."""
import os
import openpyxl

HEADERS = [
    'Record Type', 'Symbol', 'Plan Type', 'Quantity',
    'Date Acquired', 'Date Acquired (Wash Sale Toggle = On)',
    'Acquisition Cost', 'Acquisition Cost Per Share',
    'Ordinary Income Recognized', 'Ordinary Income Recognized Per Share',
    'Adjusted Cost Basis', 'Adjusted Cost Basis Per Share',
    'Date Sold', 'Total Proceeds', 'Proceeds Per Share',
    'Deferred Loss', 'Gain/Loss', 'Gain/Loss (Wash Sale Toggle = On)',
    'Adjusted Gain/Loss', 'Adjusted Gain (Loss) Per Share',
    'Capital Gains Status', 'Wash Sale Adjusted Capital Gains Status',
    'Total Wash Sale Adjustment Amount', 'Wash Sale Adjustment Amount Per Share',
    'Total Wash Sale Adjusted Cost Basis', 'Wash Sale Adjusted Cost Basis Per Share',
    'Total Wash Sale Adjusted Gain/Loss', 'Wash Sale Adjusted Gain/Loss Per Share',
    'Order Type', 'Covered Status', 'Qualified Plan?', 'Disposition Type',
    'Type', 'Grant Date', 'Grant Date FMV', 'Discount Amount',
    'Purchase Date', 'Purchase Date Fair Mkt. Value', 'Purchase Price',
    'Grant Number', '83(b) Election', 'Vest Date', 'Vest Date FMV',
    'Exercise Date', 'Exercise Date FMV', 'Grant Price', 'Order Number',
]

NONE_47 = [None] * (len(HEADERS) - 4)  # pad to full row width


def make_summary(total_qty):
    row = [None] * len(HEADERS)
    row[0] = 'Summary'
    row[3] = total_qty
    return row


def make_rsu(qty, date_acquired, vest_fmv, date_sold, price_sold,
             grant_date, grant_fmv, grant_number, order_number):
    proceeds = round(qty * price_sold, 2)
    ordinary = round(qty * vest_fmv, 2)
    gain = round(proceeds - ordinary, 2)
    return [
        'Sell', 'ACME', 'RS', qty,
        date_acquired, date_acquired,
        0, 0,
        ordinary, vest_fmv,
        ordinary, vest_fmv,
        date_sold, proceeds, price_sold,
        0, proceeds, proceeds, gain, round(gain / qty, 6),
        'Short', 'Short', 0, 0, ordinary, vest_fmv, gain, round(gain / qty, 6),
        'Sell Restricted Stock', None, None, '--',
        'Restricted Stock Unit', grant_date, grant_fmv, 0,
        '--', '0.0', 0, grant_number, 'N',
        date_acquired, vest_fmv, '--', 0, 0, order_number,
    ]


def make_espp(qty, purchase_date, purchase_fmv, purchase_price,
              date_sold, price_sold, grant_date, grant_fmv, discount,
              grant_number, order_number):
    acq_cost = round(qty * purchase_price, 2)
    proceeds = round(qty * price_sold, 2)
    ordinary = round(qty * (purchase_fmv - purchase_price), 2)
    adj_basis = round(acq_cost + ordinary, 2)
    gain = round(proceeds - acq_cost, 2)
    adj_gain = round(proceeds - adj_basis, 2)
    return [
        'Sell', 'ACME', 'ESPP', qty,
        purchase_date, purchase_date,
        acq_cost, purchase_price,
        ordinary, round(ordinary / qty, 3),
        adj_basis, round(adj_basis / qty, 3),
        date_sold, proceeds, price_sold,
        0, gain, gain, adj_gain, round(adj_gain / qty, 3),
        'Short', 'Short', 0, 0, adj_basis, round(adj_basis / qty, 3), adj_gain, round(adj_gain / qty, 3),
        'Sell ESPP', None, 'N', 'Disqualifying',
        'Employee Stock Purchase Plan', grant_date, grant_fmv, discount,
        purchase_date, purchase_fmv, purchase_price, grant_number, 'N',
        '--', 0, '--', 0, 0, order_number,
    ]


def main():
    wb = openpyxl.Workbook()
    ws = wb.active

    ws.append(HEADERS)

    # Total quantity across all sells
    ws.append(make_summary(280))

    # --- RSU sells (April 2025) ---
    # 50 shares vested 03/15/2025 @ $47, sold 04/10/2025 @ $52
    ws.append(make_rsu(50, '03/15/2025', 47.0, '04/10/2025', 52.0,
                        '03/15/2024', 45.0, '10001000', 90000001))

    # 30 shares vested 03/15/2025 @ $47, sold 04/10/2025 @ $52
    ws.append(make_rsu(30, '03/15/2025', 47.0, '04/10/2025', 52.0,
                        '06/15/2024', 42.0, '10001001', 90000001))

    # --- ESPP sell (June 2025) ---
    # 120 shares, purchased 06/05/2025 @ $38.29 (FMV $45.05, 15% discount)
    # sold 06/10/2025 @ $53.00
    ws.append(make_espp(120, '06/05/2025', 45.05, 38.29,
                         '06/10/2025', 53.0,
                         '01/01/2025', 50.0, 7.5,
                         '20001000', 90000003))

    # --- ESPP sell (December 2025) ---
    # 80 shares, purchased 12/05/2025 @ $35.70 (FMV $42.00, 15% discount)
    # sold 12/10/2025 @ $48.00
    ws.append(make_espp(80, '12/05/2025', 42.0, 35.70,
                         '12/10/2025', 48.0,
                         '07/01/2025', 44.0, 6.6,
                         '20001001', 90000004))

    script_dir = os.path.dirname(os.path.abspath(__file__))
    out = os.path.join(script_dir, 'example', 'example_2025.xlsx')
    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb.save(out)
    print(f"Saved {out}")


if __name__ == '__main__':
    main()

