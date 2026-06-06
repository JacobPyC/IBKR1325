import argparse
import re
import csv
import datetime
import math
import os

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment

from aux_functions import get_date_format, get_trades_col_names, get_dividends_col_names
from exchange_rates import ExchangeRateProvider


# ---------------------------------------------------------------------------
# Styling constants
# ---------------------------------------------------------------------------
_H1_FILL        = PatternFill(start_color='1F6B3A', end_color='1F6B3A', fill_type='solid')
_H2_FILL        = PatternFill(start_color='1F3E6B', end_color='1F3E6B', fill_type='solid')
_H1_TOTALS_FILL = PatternFill(start_color='E8F4EA', end_color='E8F4EA', fill_type='solid')
_H2_TOTALS_FILL = PatternFill(start_color='E8EEF8', end_color='E8EEF8', fill_type='solid')
_HEADER_FONT    = Font(bold=True, color='FFFFFF', size=12)
_TOTALS_FONT    = Font(bold=True, size=11)
_TABLE_GAP      = 1   # blank rows between H1 and H2 blocks


def round_half_up(x):
    if x >= 0:
        return math.floor(x + 0.5)
    return math.ceil(x - 0.5)


def _style_header_row(sheet, row, col_start, col_end, label, fill):
    for col in range(col_start, col_end + 1):
        cell = sheet.cell(row=row, column=col)
        cell.fill = fill
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal='center', vertical='center')
    label_cell = sheet.cell(row=row, column=col_start)
    label_cell.value = label
    label_cell.alignment = Alignment(horizontal='left', vertical='center')
    sheet.row_dimensions[row].height = 22


def _style_totals_row(sheet, row, col_start, col_end, fill):
    for col in range(col_start, col_end + 1):
        cell = sheet.cell(row=row, column=col)
        cell.fill = fill
        cell.font = _TOTALS_FONT


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def parse_number(number_string):
    return float(number_string.replace(',', ''))


def get_row_number(row, col_names, col_name, default=0):
    if col_name not in col_names:
        return default
    value = row[col_names[col_name]]
    if value == '':
        return default
    return parse_number(value)


def get_multiplier(asset_category):
    if asset_category == 'Equity and Index Options':
        return 100
    return 1


def _is_h1(dt):
    """Return True if the datetime falls in January-June."""
    return dt.month <= 6


def _in_half(dt, half_label):
    """Return True if dt belongs to the requested half-year."""
    return _is_h1(dt) if half_label == 'H1' else not _is_h1(dt)


def _csv_path(file_dir, csv_file_name):
    return f'{file_dir}/{csv_file_name}.csv'

def calculate_taxable_profit(profit, profit_trivial, profit_adjusted):
    if profit >= 0:
        return max(min(profit_trivial, profit_adjusted), 0)
    return min(max(profit_trivial, profit_adjusted), 0)






def resolve_opening_info(closed_lot_trades, opening_orders, multiplier):
    """
    For each ClosedLot, compute weighted open_price and allocated_open_fee
    by matching against opening Order rows. Handles mixed lots (spanning
    multiple opening orders) by deducting matched quantities greedily.
    Returns a list of dicts with 'open_price' and 'allocated_open_fee' per lot.
    """
    if not opening_orders:
        return [None] * len(closed_lot_trades)

    # Build mutable list of remaining quantities per opening order
    remaining = [
        {**info, 'remaining': abs(info['quantity'])}
        for info in opening_orders
    ]

    results = []
    for lot in closed_lot_trades:
        lot_qty = abs(lot['quantity'])
        lot_date = lot['datetime'].date() if hasattr(lot['datetime'], 'date') else lot['datetime']
        ticker = lot['ticker']

        # Filter candidates by ticker and date
        candidates = [r for r in remaining
                      if r['ticker'] == ticker
                      and (r['datetime'].date() if hasattr(r['datetime'], 'date') else r['datetime']) == lot_date
                      and r['remaining'] > 0]

        if not candidates:
            results.append(None)
            continue

        # Check if single candidate covers the entire lot exactly
        lot_basis = abs(lot['basis'])
        matched_single = None
        for c in candidates:
            fee_per_share = abs(c['fee']) / abs(c['quantity'])
            if lot['quantity'] > 0:  # long
                expected_basis = (abs(c['price']) * multiplier + fee_per_share) * lot_qty
            else:  # short
                expected_basis = (abs(c['price']) * multiplier - fee_per_share) * lot_qty
            if abs(expected_basis - lot_basis) < 0.005:
                matched_single = c
                break

        if matched_single:
            matched_single['remaining'] -= lot_qty
            results.append({
                'open_price': abs(matched_single['price']),
                'allocated_open_fee': abs(matched_single['fee']) * lot_qty / abs(matched_single['quantity']),
            })
        else:
            # Mixed lot — allocate greedily from candidates
            weighted_price = 0.0
            total_fee = 0.0
            qty_left = lot_qty
            for c in candidates:
                take = min(c['remaining'], qty_left)
                if take <= 0:
                    continue
                weighted_price += abs(c['price']) * take
                total_fee += abs(c['fee']) * take / abs(c['quantity'])
                c['remaining'] -= take
                qty_left -= take
                if qty_left <= 0:
                    break
            if qty_left > 0:
                # Couldn't fully allocate — fallback
                results.append(None)
            else:
                results.append({
                    'open_price': weighted_price / lot_qty,
                    'allocated_open_fee': total_fee,
                })

    return results


def build_closed_lot_dict(closed_lot_trade, closing_trade, allocated_close_fee, col_names, rate_provider, opening_resolution=None):
    multiplier = get_multiplier(closed_lot_trade['asset_category'])
    lot_quantity = closed_lot_trade['quantity']
    abs_quantity = abs(lot_quantity)
    position_type = 'long' if lot_quantity > 0 else 'short'

    opening_trade_price = opening_resolution['open_price'] if opening_resolution else None
    opening_trade_fee = opening_resolution['allocated_open_fee'] if opening_resolution else None

    opening_basis = closed_lot_trade['basis']
    if opening_basis == 0:
        opening_basis = lot_quantity * closed_lot_trade['price'] * multiplier

    close_gross_value = abs_quantity * closing_trade['price'] * multiplier
    if position_type == 'long':
        original_value = abs(opening_basis)
        consideration_value = close_gross_value + allocated_close_fee
        form_open_datetime = closed_lot_trade['datetime']
        tax_event_datetime = closing_trade['datetime']
        open_rate_datetime = form_open_datetime
        close_rate_datetime = tax_event_datetime
        consideration_ils_rate_datetime = tax_event_datetime
        profit_trivial_rates = (consideration_value, consideration_ils_rate_datetime,
                                original_value, open_rate_datetime)
    else:
        open_price_clean = opening_resolution['open_price'] if opening_resolution else closed_lot_trade['price']
        open_gross_value = abs_quantity * open_price_clean * multiplier
        original_value = close_gross_value - allocated_close_fee
        consideration_value = abs(opening_basis)
        form_open_datetime = closed_lot_trade['datetime']   # short-sell date (original open)
        tax_event_datetime = closing_trade['datetime']      # buy-to-cover date (close)
        open_rate_datetime = form_open_datetime
        close_rate_datetime = tax_event_datetime
        consideration_ils_rate_datetime = tax_event_datetime
        profit_trivial_rates = (consideration_value, closed_lot_trade['datetime'],
                                original_value, tax_event_datetime)

    open_currency_factor = rate_provider.get_rate(closed_lot_trade['currency'], open_rate_datetime)
    print(f"Ticker: {closed_lot_trade['ticker']}")
    close_currency_factor = rate_provider.get_rate(closed_lot_trade['currency'], close_rate_datetime)
    consideration_ils_factor = rate_provider.get_rate(
        closed_lot_trade['currency'], consideration_ils_rate_datetime)
    currency_factor_ratio = (
        close_currency_factor / open_currency_factor
        if open_currency_factor != 0 else 0)

    open_value_ILS = original_value * open_currency_factor
    open_value_ILS_adjusted_forex = open_value_ILS * currency_factor_ratio
    close_value_ILS = consideration_value * consideration_ils_factor

    profit = consideration_value - original_value
    trivial_consideration, trivial_consideration_date, trivial_original, trivial_original_date = profit_trivial_rates
    profit_trivial_ILS = (
        trivial_consideration * rate_provider.get_rate(closed_lot_trade['currency'], trivial_consideration_date)
        - trivial_original * rate_provider.get_rate(closed_lot_trade['currency'], trivial_original_date)
    )
    profit_adjusted_forex = close_value_ILS - open_value_ILS_adjusted_forex

    return {
        'currency': closed_lot_trade['currency'],
        'ticker': closed_lot_trade['ticker'],
        'tax_event_datetime': tax_event_datetime,
        'form_open_datetime': form_open_datetime,
        'open_date': form_open_datetime.strftime("%d/%m/%Y"),
        'close_date': tax_event_datetime.strftime("%d/%m/%Y"),
        'quantity': lot_quantity,
        'position_type': position_type,
        'open_price': closing_trade['price'] if position_type == 'short' else (opening_trade_price if opening_trade_price is not None else closed_lot_trade['price']),
        'close_price': (opening_trade_price if opening_trade_price is not None else closed_lot_trade['price']) if position_type == 'short' else closing_trade['price'],
        'allocated_open_fee': opening_trade_fee if opening_trade_fee is not None else 0,
        'allocated_close_fee': allocated_close_fee,
        'open_value': original_value,
        'close_value': consideration_value,
        'open_currency_factor': open_currency_factor,
        'close_currency_factor': close_currency_factor,
        'currency_factor_ratio': currency_factor_ratio,
        'open_value_ILS': open_value_ILS,
        'open_value_ILS_adjusted_forex': open_value_ILS_adjusted_forex,
        'close_value_ILS': close_value_ILS,
        'gross_sale_value_ILS': open_gross_value * open_currency_factor if position_type == 'short' else close_gross_value * consideration_ils_factor,        'profit': profit,
        'profit_trivial_ILS': profit_trivial_ILS,
        'profit_ILS_forex': calculate_taxable_profit(profit, profit_trivial_ILS, profit_adjusted_forex),
    }


def parse_trade_row(row, col_names, date_slash_format):
    datetime_string = row[col_names['datetime']]
    date_format = get_date_format(datetime_string, date_slash_format=date_slash_format)
    return {
        'trade_type': row[col_names['trade_type']],
        'asset_category': row[col_names['asset_category']],
        'currency': row[col_names['currency']],
        'ticker': row[col_names['ticker']],
        'datetime': datetime.datetime.strptime(datetime_string, date_format),
        'quantity': parse_number(row[col_names['quantity']]),
        'price': get_row_number(row, col_names, 'price'),
        'fee': get_row_number(row, col_names, 'fee'),
        'basis': get_row_number(row, col_names, 'basis'),
    }


def flush_closing_trade(closing_trade, closed_lot_trades, closed_lots_list, closed_lots_datetime_list,
                        col_names, rate_provider, opening_orders=None):
    if closing_trade is None or not closed_lot_trades:
        return True
    multiplier = get_multiplier(closed_lot_trades[0]['asset_category'])
    opening_resolutions = resolve_opening_info(closed_lot_trades, opening_orders or [], multiplier)
    total_closed_quantity = sum(abs(t['quantity']) for t in closed_lot_trades)
    for closed_lot_trade, opening_resolution in zip(closed_lot_trades, opening_resolutions):
        allocated_close_fee = (
            closing_trade['fee'] * abs(closed_lot_trade['quantity']) / total_closed_quantity
            if total_closed_quantity != 0 else 0
        )
        closed_lot_dict = build_closed_lot_dict(
            closed_lot_trade, closing_trade, allocated_close_fee, col_names, rate_provider, opening_resolution)
        closed_lots_list.append(closed_lot_dict)
        closed_lots_datetime_list.append(closed_lot_dict['tax_event_datetime'])


# ---------------------------------------------------------------------------
# Extractors
# ---------------------------------------------------------------------------
def extract_trades_data_from_csv(file_dir, csv_file_name, verbosity=0, date_slash_format='normal'):
    csv_file = _csv_path(file_dir, csv_file_name)
    rate_provider = ExchangeRateProvider()
    col_names = get_trades_col_names(csv_file)

    closed_lots_list = []
    closed_lots_datetime_list = []
    inds_sorted_close_dates = []

    if col_names:

        # Pass 1: collect all closing trades and opening orders from CSV
        closing_trade_pairs = []  # list of (closing_trade, closed_lot_trades)
        opening_orders = []
        with open(csv_file, 'r') as read_obj:
            csv_reader = csv.reader(read_obj)
            current_closing_trade = None
            current_closed_lot_trades = []
            for row in csv_reader:
                if verbosity == 1:
                    print(row)
                if (row[col_names['main']] == 'Trades' and row[col_names['header']] == 'Data'
                        and row[col_names['asset_category']] in ['Stocks', 'Equity and Index Options']):
                    trade_type = row[col_names['trade_type']]
                    if trade_type == 'Order':
                        open_close = row[col_names['code']] if 'code' in col_names else ''
                        if 'O' in open_close and 'C' not in open_close:
                            parsed_open = parse_trade_row(row, col_names, date_slash_format)
                            opening_orders.append({
                                'ticker': parsed_open['ticker'],
                                'datetime': parsed_open['datetime'],
                                'price': parsed_open['price'],
                                'fee': parsed_open['fee'],
                                'quantity': parsed_open['quantity'],
                            })
                    elif trade_type == 'Trade':
                        open_close = row[col_names['code']] if 'code' in col_names else ''
                        if 'O' in open_close and 'C' not in open_close:
                            pass  # handled by Order row
                        else:
                            if current_closing_trade is not None:
                                closing_trade_pairs.append((current_closing_trade, current_closed_lot_trades))
                            current_closing_trade = parse_trade_row(row, col_names, date_slash_format)
                            current_closed_lot_trades = []
                    elif 'ClosedLot' in trade_type:
                        current_closed_lot_trades.append(parse_trade_row(row, col_names, date_slash_format))
            if current_closing_trade is not None:
                closing_trade_pairs.append((current_closing_trade, current_closed_lot_trades))

        # Pass 2: group all ClosedLots by ticker, resolve opening info per ticker
        from collections import defaultdict
        ticker_lots = defaultdict(list)  # ticker -> list of (lot, closing_trade, allocated_close_fee)
        for closing_trade, closed_lot_trades in closing_trade_pairs:
            total_closed_quantity = sum(abs(t['quantity']) for t in closed_lot_trades)
            for lot in closed_lot_trades:
                allocated_close_fee = (
                    closing_trade['fee'] * abs(lot['quantity']) / total_closed_quantity
                    if total_closed_quantity != 0 else 0
                )
                ticker_lots[lot['ticker']].append((lot, closing_trade, allocated_close_fee))

        # Resolve opening info per ticker (all lots together, in CSV order)
        ticker_resolutions = {}
        for ticker, lot_tuples in ticker_lots.items():
            lots = [t[0] for t in lot_tuples]
            ticker_opening_orders = [o for o in opening_orders if o['ticker'] == ticker]
            multiplier = get_multiplier(lots[0]['asset_category'])
            resolutions = resolve_opening_info(lots, ticker_opening_orders, multiplier)
            ticker_resolutions[ticker] = resolutions

        # Pass 3: build closed_lot_dicts using resolved opening info
        ticker_resolution_index = defaultdict(int)
        for closing_trade, closed_lot_trades in closing_trade_pairs:
            for lot, closing_trade_inner, allocated_close_fee in [
                (lot, closing_trade, closing_trade['fee'] * abs(lot['quantity']) / sum(abs(t['quantity']) for t in closed_lot_trades)
                 if sum(abs(t['quantity']) for t in closed_lot_trades) != 0 else 0)
                for lot in closed_lot_trades
            ]:
                ticker = lot['ticker']
                idx = ticker_resolution_index[ticker]
                resolution = ticker_resolutions[ticker][idx]
                ticker_resolution_index[ticker] += 1
                closed_lot_dict = build_closed_lot_dict(
                    lot, closing_trade, allocated_close_fee, col_names, rate_provider, resolution)
                closed_lots_list.append(closed_lot_dict)
                closed_lots_datetime_list.append(closed_lot_dict['tax_event_datetime'])


        if verbosity == 1:
            for d in closed_lots_list:
                msg = (
                    f"ticker {d['ticker']}: currency {d['currency']}, "
                    f"open_date {d['open_date']}, close_date {d['close_date']}, "
                    f"position_type: {d['position_type']}, quantity={d['quantity']}, "
                    f"open_value={d['open_value']}, close_value={d['close_value']}, "
                    f"profit={d['profit']}, forex rate open={d['open_currency_factor']}, "
                    f"close={d['close_currency_factor']}"
                )
                print(msg)

        inds_sorted_close_dates = sorted(
            range(len(closed_lots_datetime_list)), key=lambda i: closed_lots_datetime_list[i]
        )

    else:
        print('no trades exist in the file.')
    return closed_lots_list, inds_sorted_close_dates


def extract_dividends_data_from_csv(file_dir, csv_file_name, verbosity=0, date_slash_format='normal'):
    """
    Single-pass extraction of Dividends and Withholding Tax rows.
    Returns (dividends_list, unmatched_wht):
      - dividends_list: current-year dividends with WHT attached
      - unmatched_wht: WHT rows with no matching dividend (no matching dividend in this CSV)
    """
    csv_file = _csv_path(file_dir, csv_file_name)
    rate_provider = ExchangeRateProvider()
    col_names = get_dividends_col_names(csv_file)

    if not col_names:
        print('no dividends exist in the file.')
        return [], []

    dividends_list = []
    all_wht_rows = []

    with open(csv_file, 'r') as read_obj:
        csv_reader = csv.reader(read_obj)
        for row in csv_reader:
            if verbosity == 1:
                print(row)
            if row[col_names['main']] not in ('Dividends', 'Withholding Tax') or row[col_names['header']] != 'Data':
                continue
            currency = row[col_names['currency']]
            if 'Total' in currency:
                continue
            datetime_string = row[col_names['datetime']]
            date_format = get_date_format(datetime_string, date_slash_format=date_slash_format)
            dt = datetime.datetime.strptime(datetime_string, date_format)
            date_str = dt.strftime("%d/%m/%Y")
            ticker = row[col_names['ticker']].split('(')[0]
            amount = float(row[col_names['amount']])

            if row[col_names['main']] == 'Dividends':
                if (dividends_list
                        and dividends_list[-1]['ticker'] == ticker
                        and dividends_list[-1]['date'] == date_str):
                    dividends_list[-1]['amount'] += amount
                else:
                    dividends_list.append({
                        'currency': currency, 'datetime': dt, 'date': date_str,
                        'ticker': ticker, 'amount': amount, 'withholding_tax': 0,
                    })
            else:  # Withholding Tax — buffer, matched after the loop
                all_wht_rows.append({
                    'ticker': ticker, 'date': date_str, 'datetime': dt,
                    'currency': currency, 'amount': amount,
                })

    dividend_keys = {(d['ticker'], d['date']) for d in dividends_list}

    unmatched_wht_raw = []
    for wht in all_wht_rows:
        if (wht['ticker'], wht['date']) in dividend_keys:
            for d in dividends_list:
                if d['ticker'] == wht['ticker'] and d['date'] == wht['date']:
                    d['withholding_tax'] += wht['amount']
                    break
        else:
            unmatched_wht_raw.append(wht)

    for d in dividends_list:
        d['currency_factor'] = rate_provider.get_rate(d['currency'], d['datetime'])
        d['dividend'] = d['amount']
        d['dividend_ILS'] = d['dividend'] * d['currency_factor']
        d['withholding_tax_ILS'] = d['withholding_tax'] * d['currency_factor']

    unmatched_wht = []
    for wht in unmatched_wht_raw:
        rate = rate_provider.get_rate(wht['currency'], wht['datetime'])
        unmatched_wht.append({
            'date': wht['date'], 'datetime': wht['datetime'],
            'ticker': wht['ticker'], 'currency': wht['currency'],
            'amount': wht['amount'], 'rate': rate,
            'amount_ils': wht['amount'] * rate,
        })

    return dividends_list, unmatched_wht
def extract_other_fees_data_from_csv(file_dir, csv_file_name, verbosity=0, date_slash_format='normal'):
    """
    Extract individual fee and interest rows from the IB CSV.
    Returns dict with:
      - 'fees':            list of {date, description, currency, amount, rate, amount_ils}
      - 'interest_debit':  list of {date, description, currency, amount, rate, amount_ils}
      - 'interest_credit': list of {date, description, currency, amount, rate, amount_ils}
    Borrow Fee rows go into fees.
    """
    csv_file = _csv_path(file_dir, csv_file_name)
    rate_provider = ExchangeRateProvider()

    result = {
        'commissions':    [],
        'fees':           [],
        'interest_debit': [],
        'interest_credit':[],
    }

    with open(csv_file, 'r') as read_obj:
        csv_reader = csv.reader(read_obj)
        for row in csv_reader:
            if len(row) < 4:
                continue
            section  = row[0].strip()
            row_type = row[1].strip()
            if row_type != 'Data':
                continue

            # Commission Details: forex pairs only (e.g. USD.ILS, EUR.USD, EUR.ILS)
            if section == 'Commission Details':
                if len(row) < 8:
                    continue
                symbol = row[4].strip()
                if not symbol:
                    continue
                parts = symbol.split('.')
                if len(parts) != 2 or not parts[0].isalpha() or not parts[1].isalpha():
                    continue
                base_currency = parts[0]
                datetime_string = row[5].strip()
                if not datetime_string:
                    continue
                try:
                    date_format = get_date_format(datetime_string, date_slash_format=date_slash_format)
                    dt = datetime.datetime.strptime(datetime_string, date_format)
                    amount = float(row[7])
                except (ValueError, IndexError, KeyError):
                    continue
                if amount == 0:
                    continue
                fx = rate_provider.get_rate(base_currency, dt)
                result['commissions'].append({
                    'date': dt.strftime('%d/%m/%Y'),
                    'description': symbol,
                    'currency': base_currency,
                    'amount': amount,
                    'rate': fx,
                    'amount_ils': amount * fx,
                    'category': 'Commission',
                })

            # Fees: ADR fees, Snapshot, market-data, etc.
            elif section == 'Fees':
                if len(row) < 7:
                    continue
                currency_val = row[3].strip()
                if 'Total' in currency_val or not currency_val:
                    continue
                datetime_string = row[4].strip()
                if not datetime_string:
                    continue
                try:
                    date_format = get_date_format(datetime_string, date_slash_format=date_slash_format)
                    dt = datetime.datetime.strptime(datetime_string, date_format)
                    amount = float(row[6])
                except (ValueError, IndexError, KeyError):
                    continue
                if amount == 0:
                    continue
                fx = rate_provider.get_rate(currency_val, dt)
                description = row[5].strip() if len(row) > 5 else ''
                result['fees'].append({
                    'date': dt.strftime('%d/%m/%Y'),
                    'description': description,
                    'currency': currency_val,
                    'amount': amount,
                    'rate': fx,
                    'amount_ils': amount * fx,
                    'category': 'Fee',
                })

            # Interest: credit + SYEP, borrow fees go to fees
            elif section == 'Interest':
                if len(row) < 6:
                    continue
                currency_val = row[2].strip()
                if 'Total' in currency_val or not currency_val:
                    continue
                datetime_string = row[3].strip()
                if not datetime_string:
                    continue
                try:
                    date_format = get_date_format(datetime_string, date_slash_format=date_slash_format)
                    dt = datetime.datetime.strptime(datetime_string, date_format)
                    amount = float(row[5])
                except (ValueError, IndexError, KeyError):
                    continue
                if amount == 0:
                    continue
                fx = rate_provider.get_rate(currency_val, dt)
                description = row[4].strip() if len(row) > 4 else ''
                entry = {
                    'date': dt.strftime('%d/%m/%Y'),
                    'description': description,
                    'currency': currency_val,
                    'amount': amount,
                    'rate': fx,
                    'amount_ils': amount * fx,
                    'category': 'Interest',
                }
                if 'Borrow Fee' in description:
                    entry['category'] = 'Fee'
                    result['fees'].append(entry)
                elif 'Debit Interest' in description:
                    result['interest_debit'].append(entry)
                else:
                    result['interest_credit'].append(entry)

    return result


# ---------------------------------------------------------------------------
# Excel writers
# ---------------------------------------------------------------------------
def _write_capital_gains_half(sheet, sheet_name, closed_lots_list, inds_sorted_close_dates,
                               half_label, start_row):
    """Write one half-year block of capital gains. Returns (next_free_row, net, sell, profit, loss)."""
    total_profit_and_loss_ILS = 0
    total_profit_ILS = 0
    total_loss_ILS = 0
    total_sell_amount_ILS = 0
    ind_line = 0

    for ind_sort in inds_sorted_close_dates:
        d = closed_lots_list[ind_sort]
        if not _in_half(d['tax_event_datetime'], half_label):
            continue

        num_row = start_row + ind_line
        sheet['A' + str(num_row)] = ind_line + 1
        sheet['B' + str(num_row)] = d['ticker']
        sheet['D' + str(num_row)] = d['currency']
        sheet['F' + str(num_row)] = d['open_date']
        sheet['M' + str(num_row)] = d['close_date']
        sheet['G' + str(num_row)] = d['open_value']
        sheet['E' + str(num_row)] = d['close_value']
        sheet['H' + str(num_row)] = d['open_value_ILS']
        sheet['I' + str(num_row)] = d['open_currency_factor']
        sheet['J' + str(num_row)] = d['close_currency_factor']
        sheet['N' + str(num_row)] = d['close_value_ILS']

        if sheet_name == 'Capital Gains (FOREX adjusted)':
            sheet['K' + str(num_row)] = d['currency_factor_ratio']
            sheet['L' + str(num_row)] = d['open_value_ILS_adjusted_forex']
            profit_ILS_name = 'profit_ILS_forex'

        else:
            raise ValueError('invalid sheet_name', sheet_name)

        if profit_ILS_name not in d:
            ind_line += 1
            continue

        value = d[profit_ILS_name]
        if value >= 0:
            sheet['O' + str(num_row)] = value
            total_profit_ILS += value
        else:
            sheet['P' + str(num_row)] = value
            total_loss_ILS += value
        total_profit_and_loss_ILS += value
        total_sell_amount_ILS += d['gross_sale_value_ILS']

        sheet['S' + str(num_row)] = d['position_type']
        sheet['T' + str(num_row)] = abs(d['quantity'])
        sheet['U' + str(num_row)] = d['open_price']
        sheet['V' + str(num_row)] = d['close_price']
        sheet['W' + str(num_row)] = abs(d['allocated_open_fee'])
        sheet['X' + str(num_row)] = abs(d['allocated_close_fee'])
        sheet['Y' + str(num_row)] = d['profit']

        ind_line += 1

    return start_row + ind_line, total_profit_and_loss_ILS, total_sell_amount_ILS, total_profit_ILS, total_loss_ILS


def _write_dividends_half(sheet, dividends_list, half_label, start_row):
    """Write one half-year block of dividends. Returns (next_free_row, div_ils, tax_ils, net_ils)."""
    total_dividends_ILS = 0
    withholding_tax_ILS = 0
    ind_line = 0

    for d in dividends_list:
        if not _in_half(d['datetime'], half_label):
            continue

        num_row = start_row + ind_line
        sheet['B' + str(num_row)] = ind_line + 1
        sheet['C' + str(num_row)] = d['date']
        sheet['D' + str(num_row)] = d['ticker']
        sheet['E' + str(num_row)] = d['currency']
        sheet['F' + str(num_row)] = d['dividend']
        sheet['G' + str(num_row)] = d['withholding_tax']
        sheet['H' + str(num_row)] = d['currency_factor']
        sheet['I' + str(num_row)] = d['dividend_ILS']
        sheet['J' + str(num_row)] = d['withholding_tax_ILS']

        total_dividends_ILS += d['dividend_ILS']
        withholding_tax_ILS += abs(d['withholding_tax_ILS'])
        ind_line += 1

    return start_row + ind_line, total_dividends_ILS, withholding_tax_ILS


def _write_unmatched_wht_table(sheet, unmatched_wht, start_row, col_start=2, col_end=10):
    """
    Write a table of prior-year withholding taxes (recovered in the current year)
    onto the Dividends sheet, starting at start_row.
    Columns: B=index, C=date, D=ticker, E=currency, F=amount(USD), G=rate, H=amount(ILS), total row.
    """
    from openpyxl.styles import Border, Side

    PRIOR_FILL  = PatternFill(start_color='6B3A1F', end_color='6B3A1F', fill_type='solid')
    TOTAL_FILL  = PatternFill(start_color='F4EAE8', end_color='F4EAE8', fill_type='solid')

    thin = Side(style='thin')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    _style_header_row(sheet, start_row, col_start, col_end,
                      'מס במקור ללא דיבידנד תואם (Unmatched Withholding Tax)', PRIOR_FILL)

    data_start = start_row + 1
    total_ils = 0

    for i, wht in enumerate(unmatched_wht):
        row = data_start + i
        sheet[f'B{row}'] = i + 1
        sheet[f'C{row}'] = wht['date']
        sheet[f'D{row}'] = wht['ticker']
        sheet[f'E{row}'] = wht['currency']
        sheet[f'F{row}'] = abs(wht['amount'])
        sheet[f'H{row}'] = wht['rate']
        sheet[f'I{row}'] = abs(round_half_up(wht['amount_ils']))
        for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
            sheet[f'{col}{row}'].border = border
        total_ils += abs(wht['amount_ils'])

    total_row = data_start + len(unmatched_wht)
    sheet[f'B{total_row}'] = 'סה"כ מס במקור שנים קודמות'
    sheet[f'I{total_row}'] = round_half_up(total_ils)
    _style_totals_row(sheet, total_row, col_start, col_end, TOTAL_FILL)
    for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
        sheet[f'{col}{total_row}'].border = border

    return total_row + 1


def _write_other_fees_sheet(xfile, other_fees_data):
    """Write individual rows for all fees and interest, then three summary totals."""
    ws = xfile['Fees & Interest Income']

    # Combine all entries into one sorted list, tagged by category
    all_fees     = other_fees_data['commissions'] + other_fees_data['fees']
    all_debit    = other_fees_data['interest_debit']
    all_credit   = other_fees_data['interest_credit']
    all_entries  = sorted(all_fees + all_debit + all_credit, key=lambda x: datetime.datetime.strptime(x['date'], '%d/%m/%Y'))

    from openpyxl.styles import Border, Side, Font
    thin = Side(style='thin')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    bold_font = Font(bold=True)
    cols = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']

    def get_type(entry):
        return entry['category']

    # Write individual rows starting at row 6
    for i, entry in enumerate(all_entries):
        row = 6 + i
        ws[f'B{row}'] = i + 1
        ws[f'C{row}'] = entry['date']
        ws[f'D{row}'] = get_type(entry)
        desc = entry['description']
        # Strip currency prefix (e.g. "USD ") and date suffix (e.g. " for Jan-2025")
        desc = re.sub(r'^[A-Z]{3}\s+', '', desc)
        desc = re.sub(r'\s+for\s+\w+-\d{4}$', '', desc)
        ws[f'E{row}'] = desc
        ws[f'F{row}'] = entry['currency']
        ws[f'G{row}'] = entry['amount']
        ws[f'H{row}'] = entry['rate']
        ws[f'I{row}'] = round_half_up(entry['amount_ils'])
        for col in cols:
            ws[f'{col}{row}'].border = border

    # Three summary rows at the bottom
    total_row = 6 + len(all_entries) + 1
    ws[f'B{total_row}']     = 'סה"כ עמלות'
    ws[f'G{total_row}']     = abs(sum(e['amount'] for e in all_fees))
    ws[f'I{total_row}']     = abs(round_half_up(sum(e['amount_ils'] for e in all_fees)))

    ws[f'B{total_row + 1}'] = 'סה"כ ריבית חובה (Debit)'
    ws[f'G{total_row + 1}'] = abs(sum(e['amount'] for e in all_debit))
    ws[f'I{total_row + 1}'] = abs(round_half_up(sum(e['amount_ils'] for e in all_debit)))

    ws[f'B{total_row + 2}'] = 'סה"כ ריבית זכות'
    ws[f'G{total_row + 2}'] = abs(sum(e['amount'] for e in all_credit))
    ws[f'I{total_row + 2}'] = abs(round_half_up(sum(e['amount_ils'] for e in all_credit)))

    for r in range(total_row, total_row + 3):
        for col in cols:
            ws[f'{col}{r}'].border = border
            ws[f'{col}{r}'].font = bold_font



def _write_cg_summary_table(sheet, h1_profit, h1_loss, h1_sell, h2_profit, h2_loss, h2_sell):
    """Write a compact summary table in columns Q-R rows 5-14."""
    TITLE_FILL = PatternFill(start_color='1F3E6B', end_color='1F3E6B', fill_type='solid')
    H1_FILL    = PatternFill(start_color='E8F4EA', end_color='E8F4EA', fill_type='solid')
    H2_FILL    = PatternFill(start_color='E8EEF8', end_color='E8EEF8', fill_type='solid')
    YEAR_FILL  = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    TITLE_FONT = Font(bold=True, color='FFFFFF', size=10)
    DATA_FONT  = Font(bold=True, size=10)
    CENTER     = Alignment(horizontal='center', vertical='center')
    LEFT       = Alignment(horizontal='left',   vertical='center')
    RIGHT      = Alignment(horizontal='right',  vertical='center')

    sheet.column_dimensions['Q'].width = 26
    sheet.column_dimensions['R'].width = 16

    rows = [
        (5,  'סיכום',                        None,                                  TITLE_FILL, TITLE_FONT),
        (6,  'רווח מחצית ראשונה',            round_half_up(h1_profit),              H1_FILL,    DATA_FONT),
        (7,  'הפסד מחצית ראשונה',            round_half_up(abs(h1_loss)),           H1_FILL,    DATA_FONT),
        (8,  'סכום מכירות מחצית ראשונה',     round_half_up(h1_sell),                H1_FILL,    DATA_FONT),
        (9,  'רווח מחצית שנייה',             round_half_up(h2_profit),              H2_FILL,    DATA_FONT),
        (10, 'הפסד מחצית שנייה',             round_half_up(abs(h2_loss)),           H2_FILL,    DATA_FONT),
        (11, 'סכום מכירות מחצית שנייה',      round_half_up(h2_sell),                H2_FILL,    DATA_FONT),
        (12, 'הפסד כולל שנתי',                    round_half_up(abs(h1_loss + h2_loss)), YEAR_FILL,  DATA_FONT),
        (13, 'מכירות שנתי',                  round_half_up(h1_sell + h2_sell),      YEAR_FILL,  DATA_FONT),
        (14, 'רווח שנתי לאחר קיזוז הפסדים',                    round_half_up((h1_profit + h2_profit) + (h1_loss + h2_loss)), YEAR_FILL, DATA_FONT),
    ]

    for row_num, label, value, fill, font in rows:
        sheet.row_dimensions[row_num].height = 18

        s = sheet.cell(row=row_num, column=17, value=label)
        s.fill = fill
        s.font = font
        s.alignment = CENTER if value is None else LEFT

        t = sheet.cell(row=row_num, column=18, value=value)
        t.fill = fill
        t.font = font
        t.alignment = RIGHT


def write_tax_form_files(file_dir, csv_file_name, closed_lots_list, inds_sorted_close_dates,
                         dividends_list, other_fees_data, unmatched_wht=None):
    template_file = os.path.dirname(os.path.abspath(__file__)) + '/tax_forms_template.xlsx'
    xfile = openpyxl.load_workbook(template_file)

    CG_COL_START,  CG_COL_END  = 1, 25
    DIV_COL_START, DIV_COL_END = 2, 10


    # Capital Gains sheets
    if closed_lots_list:
        sheet_names = ['Capital Gains (FOREX adjusted)']

        for sheet_name in sheet_names:
            sheet = xfile[sheet_name]
            sheet.freeze_panes = 'A5'  # freeze rows 1-4 (column headers)

            _style_header_row(sheet, 5, CG_COL_START, CG_COL_END, 'ינואר - יוני', _H1_FILL)
            h1_next, h1_net, h1_sell, h1_profit, h1_loss = _write_capital_gains_half(
                sheet, sheet_name, closed_lots_list, inds_sorted_close_dates, 'H1', 6)

            h2_header = h1_next + _TABLE_GAP
            _style_header_row(sheet, h2_header, CG_COL_START, CG_COL_END, 'יולי -דצמבר', _H2_FILL)
            h2_next, h2_net, h2_sell, h2_profit, h2_loss = _write_capital_gains_half(
                sheet, sheet_name, closed_lots_list, inds_sorted_close_dates, 'H2', h2_header + 1)

            # Summary table in column A
            _write_cg_summary_table(sheet, h1_profit, h1_loss, h1_sell, h2_profit, h2_loss, h2_sell)


    # Dividends sheet
    if dividends_list:
        sheet = xfile['Dividends']
        sheet.freeze_panes = 'A5'  # freeze rows 1-4 (column headers)

        _style_header_row(sheet, 5, DIV_COL_START, DIV_COL_END, 'ינואר - יוני', _H1_FILL)
        h1_next, h1_div, h1_tax = _write_dividends_half(sheet, dividends_list, 'H1', 6)
        sheet['B' + str(h1_next)] = 'סכום מחצית ראשונה'
        sheet['I' + str(h1_next)] = h1_div
        sheet['J' + str(h1_next)] = h1_tax
        _style_totals_row(sheet, h1_next, DIV_COL_START, DIV_COL_END, _H1_TOTALS_FILL)

        h2_header = h1_next + _TABLE_GAP
        _style_header_row(sheet, h2_header, DIV_COL_START, DIV_COL_END, 'יולי - דצמבר', _H2_FILL)
        h2_next, h2_div, h2_tax = _write_dividends_half(sheet, dividends_list, 'H2', h2_header + 1)
        sheet['B' + str(h2_next)] = 'סכום מחצית שנייה'
        sheet['I' + str(h2_next)] = h2_div
        sheet['J' + str(h2_next)] = h2_tax
        _style_totals_row(sheet, h2_next, DIV_COL_START, DIV_COL_END, _H2_TOTALS_FILL)

        # Year total
        year_row = h2_next + 2
        sheet['B' + str(year_row)] = 'סכום שנתי'
        sheet['I' + str(year_row)] = round_half_up(h1_div + h2_div)
        sheet['J' + str(year_row)] = round_half_up(h1_tax + h2_tax)
        _style_totals_row(sheet, year_row, DIV_COL_START, DIV_COL_END,
                          PatternFill(start_color='BFBFBF', end_color='BFBFBF', fill_type='solid'))

        # Prior-year withholding tax table (separate block below)
        if unmatched_wht:
            prior_start = year_row + 2
            _write_unmatched_wht_table(sheet, unmatched_wht, prior_start, DIV_COL_START, DIV_COL_END)

    # Other Fees & Interest sheet
    _write_other_fees_sheet(xfile, other_fees_data)

    file_path = f'{file_dir}/tax_forms_{csv_file_name}.xlsx'
    xfile.save(file_path)
    print(f'Written: {file_path}')


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def _extract_all(file_dir, csv_file_name, verbosity, date_slash_format):
    return (
        extract_trades_data_from_csv(
            file_dir, csv_file_name, verbosity=verbosity,
            date_slash_format=date_slash_format),
        extract_dividends_data_from_csv(
            file_dir, csv_file_name, verbosity=verbosity,
            date_slash_format=date_slash_format),
        extract_other_fees_data_from_csv(
            file_dir, csv_file_name, verbosity=verbosity,
            date_slash_format=date_slash_format),
    )


def generate_tax_forms(file_dir, csv_file_name, verbosity=0):
    try:
        trades, dividends_tuple, fees = _extract_all(
            file_dir, csv_file_name, verbosity, 'normal')
    except ValueError:
        print("*** failed to use date_slash_format='normal', attempting date_slash_format='USA'")
        trades, dividends_tuple, fees = _extract_all(
            file_dir, csv_file_name, verbosity, 'USA')

    closed_lots_list, inds_sorted_close_dates = trades
    dividends_list, unmatched_wht = dividends_tuple

    write_tax_form_files(file_dir, csv_file_name, closed_lots_list, inds_sorted_close_dates,
                         dividends_list, fees, unmatched_wht=unmatched_wht)
    print('Finished generating tax forms.')


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Run tax forms generator")
    parser.add_argument("-dir", "--dir", type=str, required=True, help="directory path of the csv file")
    parser.add_argument("-csv_file_name", "--csv_name", type=str, required=True, help="csv file name (without suffix)")
    parser.add_argument("-verbosity", "--verbosity", default=0, type=int, required=False,
                        help="verbosity of output during run")
    args = parser.parse_args()
    generate_tax_forms(args.dir, args.csv_name, args.verbosity)