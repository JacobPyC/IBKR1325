import argparse
import re
import csv
import datetime
import math
import os

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment

from aux_functions import get_date_format, get_trades_col_names, get_dividends_col_names
# from cpi_israel import get_israel_cpi_value  # CPI
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



# def add_cpi_fields(closed_lot_dict):  # CPI
#     closed_lot_dict['open_cpi'] = get_israel_cpi_value(closed_lot_dict['form_open_datetime'])  # CPI
#     closed_lot_dict['close_cpi'] = get_israel_cpi_value(closed_lot_dict['tax_event_datetime'])  # CPI
#     closed_lot_dict['cpi_ratio'] = closed_lot_dict['close_cpi'] / closed_lot_dict['open_cpi']  # CPI
#     closed_lot_dict['open_value_ILS_adjusted_cpi'] = (  # CPI
#         closed_lot_dict['open_value_ILS'] * closed_lot_dict['cpi_ratio']  # CPI
#     )  # CPI
#     closed_lot_dict['profit_ILS_cpi'] = calculate_taxable_profit(  # CPI
#         closed_lot_dict['profit'],  # CPI
#         closed_lot_dict['profit_trivial_ILS'],  # CPI
#         closed_lot_dict['close_value_ILS'] - closed_lot_dict['open_value_ILS_adjusted_cpi'],  # CPI
#     )  # CPI



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
    cpi_extraction_succeeded = True
    for closed_lot_trade, opening_resolution in zip(closed_lot_trades, opening_resolutions):
        allocated_close_fee = (
            closing_trade['fee'] * abs(closed_lot_trade['quantity']) / total_closed_quantity
            if total_closed_quantity != 0 else 0
        )
        closed_lot_dict = build_closed_lot_dict(
            closed_lot_trade, closing_trade, allocated_close_fee, col_names, rate_provider, opening_resolution)
        # try:  # CPI
        #     add_cpi_fields(closed_lot_dict)  # CPI
        # except Exception:  # CPI
        #     print("*** failed to extract CPI data, skipping it because it is used only for educational purpose.")  # CPI
        #     cpi_extraction_succeeded = False  # CPI
        closed_lots_list.append(closed_lot_dict)
        closed_lots_datetime_list.append(closed_lot_dict['tax_event_datetime'])
    return cpi_extraction_succeeded


# ---------------------------------------------------------------------------
# Extractors
# ---------------------------------------------------------------------------
def extract_trades_data_from_csv(file_dir, csv_file_name, verbosity=0, date_slash_format='normal'):
    csv_file = _csv_path(file_dir, csv_file_name)
    rate_provider = ExchangeRateProvider()
    col_names = get_trades_col_names(csv_file)

    cpi_extraction_succeeded = False
    closed_lots_list = []
    closed_lots_datetime_list = []
    inds_sorted_close_dates = []

    if col_names:
        cpi_extraction_succeeded = True

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
                # if cpi_extraction_succeeded:  # CPI
                #     msg += (  # CPI
                #         f", cpi open={d['open_cpi']}, close={d['close_cpi']}, ratio={d['cpi_ratio']}"  # CPI
                #     )  # CPI
                print(msg)

        inds_sorted_close_dates = sorted(
            range(len(closed_lots_datetime_list)), key=lambda i: closed_lots_datetime_list[i]
        )

    else:
        print('no trades exist in the file.')
    return closed_lots_list, inds_sorted_close_dates, cpi_extraction_succeeded


def _get_statement_year(csv_file):
    """Extract the statement year from the Period row in the CSV header."""
    with open(csv_file, 'r') as f:
        for row in csv.reader(f):
            if len(row) >= 4 and row[0] == 'Statement' and row[2] == 'Period':
                # e.g. 'January 1, 2025 - December 31, 2025'
                import re
                years = re.findall(r'\d{4}', row[3])
                if years:
                    return int(years[-1])  # use end year
    return None


def extract_dividends_data_from_csv(file_dir, csv_file_name, verbosity=0, date_slash_format='normal'):
    csv_file = _csv_path(file_dir, csv_file_name)
    rate_provider = ExchangeRateProvider()
    col_names = get_dividends_col_names(csv_file)
    statement_year = _get_statement_year(csv_file)

    if col_names:
        with open(csv_file, 'r') as read_obj:
            csv_reader = csv.reader(read_obj)
            dividends_list = []
            unmatched_withholding = []  # (event_dict,) for withholding with no dividend match
            for row in csv_reader:
                if verbosity == 1:
                    print(row)
                if row[col_names['main']] in ['Dividends', 'Withholding Tax'] and row[col_names['header']] == 'Data':
                    event_dict = {}
                    event_dict['currency'] = row[col_names['currency']]
                    if 'Total' not in event_dict['currency']:
                        datetime_string = row[col_names['datetime']]
                        date_format = get_date_format(datetime_string, date_slash_format=date_slash_format)
                        event_dict['datetime'] = datetime.datetime.strptime(datetime_string, date_format)
                        event_dict['date'] = event_dict['datetime'].strftime("%d/%m/%Y")
                        event_dict['ticker'] = row[col_names['ticker']].split('(')[0]
                        event_dict['amount'] = float(row[col_names['amount']])
                        if row[col_names['main']] == 'Dividends':
                            if (len(dividends_list) > 0
                                    and dividends_list[-1]['ticker'] == event_dict['ticker']
                                    and dividends_list[-1]['date'] == event_dict['date']):
                                dividends_list[-1]['amount'] += event_dict['amount']
                            else:
                                event_dict['withholding_tax'] = 0
                                dividends_list.append(event_dict)
                        elif row[col_names['main']] == 'Withholding Tax':
                            matched = False
                            for dividend_dict in dividends_list:
                                if (dividend_dict['ticker'] == event_dict['ticker']
                                        and dividend_dict['date'] == event_dict['date']):
                                    dividend_dict['withholding_tax'] += event_dict['amount']
                                    matched = True
                                    break
                            if not matched:
                                unmatched_withholding.append(event_dict)

            # Insert standalone entries for unmatched withholding rows
            for wh in unmatched_withholding:
                # Check again — maybe a dividend was added later in the loop (shouldn't happen but be safe)
                already_covered = any(
                    d['ticker'] == wh['ticker'] and d['date'] == wh['date']
                    for d in dividends_list
                )
                if already_covered:
                    for d in dividends_list:
                        if d['ticker'] == wh['ticker'] and d['date'] == wh['date']:
                            d['withholding_tax'] += wh['amount']
                else:
                    dividends_list.append({
                        'currency':       wh['currency'],
                        'datetime':       wh['datetime'],
                        'date':           wh['date'],
                        'ticker':         wh['ticker'],
                        'amount':         0,
                        'withholding_tax': wh['amount'],
                    })


        # Separate prior-year withholding entries (date year != statement year).
        # These are typically refunds (negative) or corrections for a different tax year.
        prior_year_withholding = []
        current_year_list = []
        for d in dividends_list:
            if statement_year and d["datetime"].year != statement_year:
                prior_year_withholding.append(d)
            else:
                current_year_list.append(d)
        dividends_list = current_year_list

        # Sort by date
        dividends_list.sort(key=lambda d: d["datetime"])
        prior_year_withholding.sort(key=lambda d: d["datetime"])

        for dividend_dict in dividends_list + prior_year_withholding:
            dividend_dict["currency_factor"] = rate_provider.get_rate(
                dividend_dict["currency"], dividend_dict["datetime"])
            dividend_dict["dividend"] = dividend_dict["amount"]
            dividend_dict["dividend_ILS"] = dividend_dict["dividend"] * dividend_dict["currency_factor"]
            dividend_dict["withholding_tax_ILS"] = dividend_dict["withholding_tax"] * dividend_dict["currency_factor"]

        return dividends_list, prior_year_withholding

    else:
        print("no dividends exist in the file.")
        return [], []


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

        # elif sheet_name == 'Capital Gains (CPI adjusted)':  # CPI
        #     sheet['K' + str(num_row)] = d.get('cpi_ratio', '')  # CPI
        #     sheet['L' + str(num_row)] = d.get('open_value_ILS_adjusted_cpi', '')  # CPI
        #     profit_ILS_name = 'profit_ILS_cpi'  # CPI
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
                         dividends_list, cpi_extraction_succeeded, other_fees_data,
                         prior_year_withholding=None):
    template_file = os.path.dirname(os.path.abspath(__file__)) + '/tax_forms_template.xlsx'
    xfile = openpyxl.load_workbook(template_file)

    CG_COL_START,  CG_COL_END  = 1, 25
    DIV_COL_START, DIV_COL_END = 2, 10


    # Capital Gains sheets
    if closed_lots_list:
        sheet_names = ['Capital Gains (FOREX adjusted)']
        # if cpi_extraction_succeeded:  # CPI
        #     sheet_names.append('Capital Gains (CPI adjusted)')  # CPI

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

        # Prior-year withholding refunds section
        if prior_year_withholding:
            _PRIOR_FILL = PatternFill(start_color='7B3F00', end_color='7B3F00', fill_type='solid')
            _PRIOR_TOTALS_FILL = PatternFill(start_color='FDF3E7', end_color='FDF3E7', fill_type='solid')
            pyw_header = year_row + 2
            _style_header_row(sheet, pyw_header, DIV_COL_START, DIV_COL_END,
                              'החזר ניכוי מס במקור משנים קודמות', _PRIOR_FILL)
            pyw_row = pyw_header + 1
            total_pyw_tax_ILS = 0
            for i, d in enumerate(prior_year_withholding):
                sheet['B' + str(pyw_row)] = i + 1
                sheet['C' + str(pyw_row)] = d['date']
                sheet['D' + str(pyw_row)] = d['ticker']
                sheet['E' + str(pyw_row)] = d['currency']
                sheet['F' + str(pyw_row)] = d['dividend']
                sheet['G' + str(pyw_row)] = d['withholding_tax']
                sheet['H' + str(pyw_row)] = d['currency_factor']
                sheet['I' + str(pyw_row)] = d['dividend_ILS']
                sheet['J' + str(pyw_row)] = d['withholding_tax_ILS']
                total_pyw_tax_ILS += d['withholding_tax_ILS']
                pyw_row += 1
            sheet['B' + str(pyw_row)] = 'סה"כ החזר ניכוי מס במקור משנים קודמות'
            sheet['J' + str(pyw_row)] = round_half_up(total_pyw_tax_ILS)
            _style_totals_row(sheet, pyw_row, DIV_COL_START, DIV_COL_END, _PRIOR_TOTALS_FILL)

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
        trades, dividends_result, fees = _extract_all(
            file_dir, csv_file_name, verbosity, 'normal')
    except ValueError:
        print("*** failed to use date_slash_format='normal', attempting date_slash_format='USA'")
        trades, dividends_result, fees = _extract_all(
            file_dir, csv_file_name, verbosity, 'USA')

    closed_lots_list, inds_sorted_close_dates, cpi_extraction_succeeded = trades
    dividends_list, prior_year_withholding = dividends_result
    write_tax_form_files(file_dir, csv_file_name, closed_lots_list, inds_sorted_close_dates,
                         dividends_list, cpi_extraction_succeeded, fees,
                         prior_year_withholding=prior_year_withholding)
    print('Finished generating tax forms.')


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Run tax forms generator")
    parser.add_argument("-dir", "--dir", type=str, required=True, help="directory path of the csv file")
    parser.add_argument("-csv_file_name", "--csv_name", type=str, required=True, help="csv file name (without suffix)")
    parser.add_argument("-verbosity", "--verbosity", default=0, type=int, required=False,
                        help="verbosity of output during run")
    args = parser.parse_args()
    generate_tax_forms(args.dir, args.csv_name, args.verbosity)