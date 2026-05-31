"""
USD/ILS exchange rate via Bank of Israel (BOI).
The Bank of Israel publishes the *Official Representative Exchange Rate*
(שער יציג) on every Israeli business day. It is sampled from interbank
trading at a specific time of day and is the rate used for accounting,
contracts and tax purposes in Israel.
API documented at:
  https://www.boi.org.il/media/tzxbuhhj/extracting-representative-exchange-rates-from-the-new-series-database.pdf
"""
import csv
import io
from datetime import date, timedelta
from urllib.parse import urlencode

import requests


BOI_BASE = (
    'https://edge.boi.org.il/FusionEdgeServer/sdmx/v2/data/dataflow/'
    'BOI.STATISTICS/EXR/1.0'
)

BOI_SERIES_BY_CURRENCY = {
    'USD': 'RER_USD_ILS',
    'EUR': 'RER_EUR_ILS',
    'GBP': 'RER_GBP_ILS',
    'JPY': 'RER_JPY_ILS',
    'AUD': 'RER_AUD_ILS',
    'CAD': 'RER_CAD_ILS',
    'CHF': 'RER_CHF_ILS',
    'DKK': 'RER_DKK_ILS',
    'NOK': 'RER_NOK_ILS',
    'SEK': 'RER_SEK_ILS',
    'ZAR': 'RER_ZAR_ILS',
    'JOD': 'RER_JOD_ILS',
    'LBP': 'RER_LBP_ILS',
    'EGP': 'RER_EGP_ILS',
}

BOI_UNITS_BY_CURRENCY = {
    'JPY': 100,
    'LBP': 10,
}


def _get_boi_rate_for_date(d: date, currency: str) -> float | None:
    """Fetch BOI representative rate for a single date. Returns None if no rate published (weekend/holiday). Retries on failure."""
    if currency not in BOI_SERIES_BY_CURRENCY:
        raise ValueError(f'Bank of Israel rate is not configured for {currency}')
    series = BOI_SERIES_BY_CURRENCY[currency]
    params = {
        'startperiod': d.isoformat(),
        'endperiod':   d.isoformat(),
        'format':      'csv',
    }
    url = f'{BOI_BASE}/{series}?{urlencode(params)}'
    last_error = None
    for _ in range(3):
        try:
            response = requests.get(url, timeout=(5, 15))
            if response.status_code == 404 or not response.text.strip():
                return None
            response.raise_for_status()
            reader = csv.DictReader(io.StringIO(response.text))
            rows = [row for row in reader if row.get('TIME_PERIOD') == d.isoformat()]
            if not rows:
                return None
            unit = BOI_UNITS_BY_CURRENCY.get(currency, 1)
            return float(rows[0]['OBS_VALUE']) / unit
        except Exception as e:
            last_error = e
    raise RuntimeError(f'BOI request failed after 3 attempts for {currency} on {d}: {last_error}')


class ExchangeRateProvider:
    def __init__(self):
        self.rate_cache = {}
        self.warning_cache = set()

    def get_rate(self, currency, dt) -> float:
        currency = currency.upper()
        if currency == 'ILS':
            return 1.0

        d = dt.date() if hasattr(dt, 'date') else dt
        cache_key = (currency, d)
        if cache_key in self.rate_cache:
            return self.rate_cache[cache_key]

        rate = self._get_boi_rate(currency, d)
        self.rate_cache[cache_key] = rate
        return rate

    def _get_boi_rate(self, currency: str, d: date) -> float:
        """Walk back day by day until a published rate is found (up to 10 days). Retries each day 3 times."""
        cur = d
        for _ in range(10):
            rate = _get_boi_rate_for_date(cur, currency)
            if rate is not None:
                if cur != d:
                    self._warn_once(
                        ('boi_previous_rate', currency, d),
                        f'*** using previous Bank of Israel rate for {currency}: {cur} for {d}'
                    )
                return rate
            cur -= timedelta(days=1)
        raise RuntimeError(f'No BOI rate found within 10 days of {d.isoformat()}')

    def _warn_once(self, key, message):
        if key not in self.warning_cache:
            print(message)
            self.warning_cache.add(key)