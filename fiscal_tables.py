"""Period-aware extraction from Fiscal.ai's rendered, accessible tables."""
import json
import re


def period_key(label):
    label = label.strip()
    if label in {'LTM', 'TTM'}:
        return (9999, 12)
    match = re.search(r"(?:'([0-9]{2})|\b(20[0-9]{2}))", label)
    if not match:
        return None
    year = 2000 + int(match[1]) if match[1] else int(match[2])
    months = 'Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec'.split()
    month = next((i + 1 for i, m in enumerate(months) if m in label), 12)
    return year, month


def table_rows(page):
    # Read cells together, preserving empty columns and using each table's own headers.
    return page.locator('table').evaluate_all("""tables => tables.map(table => ({
        headers: Array.from(table.querySelectorAll('thead tr:last-child th')).map(c => c.innerText.trim()),
        rows: Array.from(table.querySelectorAll('tbody tr')).map(r =>
            Array.from(r.querySelectorAll(':scope > td, :scope > th')).map(c => c.innerText.trim()))
    }))""")


def metric_points(page, metric, section=None):
    for table in table_rows(page):
        headers = table['headers']
        if section and (not headers or headers[0].casefold() != section.casefold()):
            continue
        for cells in table['rows']:
            if cells and cells[0].casefold() == metric.casefold() and len(cells) == len(headers):
                return list(zip(headers[1:], cells[1:]))
    return []


def select_latest(points):
    from auto import parse_metric_text
    eligible = [(period_key(label), value) for label, value in points
                if '(E)' not in label and label != 'NTM'
                and period_key(label) is not None and parse_metric_text(value) is not None]
    return max(eligible, key=lambda point: point[0])[1] if eligible else 'N/A'


def extract_metric(page, metric, click_first=None, pick='latest'):
    if click_first:
        page.get_by_text(click_first, exact=True).first.click(timeout=3000)
    points = metric_points(page, metric)
    if pick in {'latest', 'last'}:
        return select_latest(points)
    from auto import parse_metric_text
    return next((value for _, value in points if parse_metric_text(value) is not None), 'N/A')


def extract_metric_history(page, metric, section=None):
    from auto import parse_metric_text
    points = [(label, parse_metric_text(value)) for label, value in metric_points(page, metric, section)
              if label not in {'LTM', 'TTM', 'NTM'} and '(E)' not in label and period_key(label)]
    return json.dumps([{'label': label, 'value': value}
                       for label, value in sorted(points, key=lambda p: period_key(p[0])) if value is not None])


def extract_ratio_history(page, section_name, metric):
    return extract_metric_history(page, metric, section_name)


def select_first_estimate(points):
    from auto import parse_metric_text
    future = [(period_key(label), value) for label, value in points if '(E)' in label and period_key(label)]
    if not future:
        return 'N/A'
    # Do not substitute a later year or an actual for the missing first forecast year.
    value = min(future, key=lambda point: point[0])[1]
    return value if parse_metric_text(value) is not None else 'N/A'


def extract_estimate_mean_for_first_future_column(page, metric_tab_name):
    page.get_by_text('Annual', exact=True).first.click(timeout=5000)
    page.wait_for_timeout(1500)
    set_millions(page)
    return select_first_estimate(metric_points(page, 'Mean'))


def set_millions(page):
    control = page.get_by_text('M', exact=True)
    if control.count():
        control.first.click(timeout=5000)
        page.wait_for_timeout(500)


def overview_value(page, label):
    from auto import parse_metric_text
    lines = [line.strip() for line in page.locator('body').inner_text().splitlines() if line.strip()]
    for index, line in enumerate(lines[:-1]):
        if line.casefold() == label.casefold():
            value = lines[index + 1]
            return value if parse_metric_text(value) is not None else 'N/A'
    return 'N/A'


def extract_overview_growth_metric(page, label):
    value = overview_value(page, label)
    return value if value.endswith('%') else 'N/A'


def extract_ntm_valuation_metric(page, label):
    from auto import parse_metric_text
    lines = [line.strip() for line in page.locator('body').inner_text().splitlines() if line.strip()]
    try:
        start = lines.index('Valuation (NTM)')
    except ValueError:
        return 'N/A'
    for index in range(start + 1, len(lines) - 1, 2):
        if lines[index] == label:
            value = lines[index + 1]
            return value if parse_metric_text(value) is not None else 'N/A'
        if lines[index] in {'Financial Health', 'Growth (CAGR)', 'Dividends'}:
            break
    return 'N/A'


def current_price(page, stock_id):
    return overview_value(page, stock_id.replace('%20', ' '))
