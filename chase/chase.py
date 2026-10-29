import argparse
import csv
import io
import os
import re
import subprocess
from datetime import datetime as dt
from decimal import Decimal

try:
    import pdfplumber
    import pypdfium2 as pdfium
    from PIL import Image, ImageOps
except ImportError as exc:
    raise SystemExit(
        'Missing OCR dependencies. Install them with: '
        'py -m pip install pdfplumber pypdfium2 pillow'
    ) from exc

from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTAnno, LTChar, LTPage, LTTextBox, LTTextLine
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.pdfpage import PDFPage


CSV_DATE_FORMAT = '%m/%d/%Y'
CSV_HEADERS = [
    'Trans Date',
    'Check No',
    'Description',
    'Account',
    'Amount',
    'Out Amount',
    'Balance',
]
START_DATE = None
END_DATE = None
TESSERACT_CMD = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

MONEY_RE = re.compile(r'\$?-?\d[\d,]*\.\d{2}')
DATE_RE = re.compile(r'^\d{1,2}/\d{1,2}\b')
MONTH_RANGE_RE = re.compile(
    r'([A-Za-z]+ \d{1,2}, \d{4}) through ([A-Za-z]+ \d{1,2}, \d{4})'
)

WITHDRAWAL_SECTIONS = {
    'CHECKS PAID',
    'ATM & DEBIT CARD WITHDRAWALS',
    'ELECTRONIC WITHDRAWALS',
    'OTHER WITHDRAWALS',
    'FEES',
}
TARGET_SECTIONS = {'DEPOSITS AND ADDITIONS'} | WITHDRAWAL_SECTIONS

SECTION_MARKERS = [
    ('deposits and additions', 'DEPOSITS AND ADDITIONS'),
    ('checks paid', 'CHECKS PAID'),
    ('atm debit withdrawal', 'ATM & DEBIT CARD WITHDRAWALS'),
    ('electronic withdrawal', 'ELECTRONIC WITHDRAWALS'),
    ('other withdrawals', 'OTHER WITHDRAWALS'),
    ('fees', 'FEES'),
]

SECTION_HEADINGS = {
    'DEPOSITS AND ADDITIONS': 'DEPOSITS AND ADDITIONS',
    'DEPOSITS AND ADDITIONS (CONTINUED)': 'DEPOSITS AND ADDITIONS',
    'CHECKS PAID': 'CHECKS PAID',
    'ATM & DEBIT CARD WITHDRAWALS': 'ATM & DEBIT CARD WITHDRAWALS',
    'ELECTRONIC WITHDRAWALS': 'ELECTRONIC WITHDRAWALS',
    'ELECTRONIC WITHDRAWALS (CONTINUED)': 'ELECTRONIC WITHDRAWALS',
    'OTHER WITHDRAWALS': 'OTHER WITHDRAWALS',
    'FEES': 'FEES',
}

ACCOUNT_RULES = [
    (r'\b(venmo|online transfer|funds transfer|transfer from|transfer to)\b', 'Transfer'),
    (r'\b(deposit|remote online deposit)\b', '_Deposit'),
    (r'\b(ach credit|merchant services dep|stripe|square|shopify payments)\b', 'Sales / Revenue'),
    (r'\b(payroll|adp|gusto|paychex|quickbooks payroll|intuit payroll)\b', 'Payroll Expense'),
    (r'\b(irs|Dept of Labor)\b', 'Payroll Tax'),
    (r'\b(NYS Dol Ui|NYS Dtf Wt)\b', 'Corp Tax'),
    (r'\b(us treasury|eftps|state tax|tax payment|franchise tax|department of revenue)\b', 'Taxes'),
    (r'\b(rent|lease payment|property management)\b', 'Rent'),
    (r'\b(Nyc Dept of Fin)\b', 'Property tax'),
    (r'\b(comcast|xfinity|verizon|at&t|att|tmobile|t-mobile|spectrum|internet|phone)\b', 'Telephone & Internet'),
    (r'\b(home depot|ikea|lowes)\b', 'Repairs & Maintenance'),
    (r"\b(bjs|bj'?s|costco|walmart|wal-mart|dollar tree|dollartree|webstauran|target|Lidl|Supermarket)\b", 'Purchase'),
    (r'\b(pge|pg&e|duke energy|con edison|national grid|electric|gas bill|water|sewer|NGRID|Spire Bill Pay|MLGW|Sola-Fidlty|SD Gas & Elec|Nycwaterboard|Pseg)\b', 'Utilities'),
    (r'\b(insurance|geico|progressive|state farm|the hartford|workers comp|WESTGUARD|Uticafirst|Kaiser)\b', 'Insurance'),
    (r'\b(amazon|staples|office depot|officedepot|costco|walmart|target)\b', 'Office Expense'),
    (r'\b(microsoft|adobe|dropbox|zoom|slack|notion|aws|amazon web services|github|godaddy|intuit|quickbooks|Lilypad)\b', 'Software & Subscriptions'),
    (r'\b(Google|facebook|Facebk)\b', 'Ads'),
    (r'\b(ups|fedex|usps|shipping|postage)\b', 'Postage'),
    (r'\b(D\&T Tax)\b', 'Accounting'),
    (r'\b(restaurant|cafe|coffee|starbucks|doordash|ubereats|grubhub|mcdonald|chipotle|hot pot)\b', 'Meal & Entertainment'),
    (r'\b(shell|chevron|exxon|mobil|bp |sunoco|gas station|fuel)\b', 'Gasoline'),
    (r'\b(parking|toll)\b', 'Auto Expense'),
    (r'\b(Head Rush Technologie)\b', 'Equipments'),
    (r'\b(Security)\b', 'Security'),
    (r'\b(toll|uber|lyft|Mta)\b', 'Travel'),
    (r'\b(bank fee|monthly service fee|wire fee|overdraft|chargeback fee|merchant fee|service charge)\b', 'Bank Charges'),
    (r'\b(legal|law office|attorney|accounting|cpa|bookkeeping|consulting)\b', 'Professional Fees'),
    (r'\b(loan payment|loan pmt|credit card payment|cardmember services|amex epay|chase card)\b', 'Loan / Credit Card Payment'),
]


def load_account_rules(file_name):
    rules = []
    with open(file_name, newline='', encoding='utf-8-sig') as stream:
        reader = csv.DictReader(stream)
        if reader.fieldnames and {'Pattern', 'Account'}.issubset(reader.fieldnames):
            for row in reader:
                pattern = (row.get('Pattern') or '').strip()
                account = (row.get('Account') or '').strip()
                if pattern and account:
                    rules.append((pattern, account))
        else:
            stream.seek(0)
            for row in csv.reader(stream):
                if len(row) >= 2 and row[0].strip() and row[1].strip():
                    rules.append((row[0].strip(), row[1].strip()))
    return rules


def categorize_account(description, section='', check_no='', rules=None):
    if check_no or section == 'CHECKS PAID':
        return ''

    description = ' '.join(str(description).lower().split())
    for pattern, account in (rules or ACCOUNT_RULES):
        if re.search(pattern, description, flags=re.I):
            return account
    return ''


def clean_money(value):
    return Decimal(value.replace('$', '').replace(',', '').strip())


def fmt_decimal(value):
    return f'{value:.2f}'


def normalize_amount(amount_str, section):
    amount = clean_money(amount_str)
    if section in WITHDRAWAL_SECTIONS and amount > 0:
        amount = -amount
    return fmt_decimal(amount)


def sort_and_filter(data):
    filtered = []
    for row in data:
        date = row['Trans Date']
        if START_DATE and date < START_DATE:
            continue
        if END_DATE and date >= END_DATE:
            continue
        filtered.append(row)
    return sorted(
        filtered,
        key=lambda row: (
            row['Trans Date'],
            row.get('_file_order', 0),
            row.get('_order', 0),
        ),
    )


class LineConverter(PDFPageAggregator):
    def __init__(self, rsrcmgr, pageno=1, laparams=None):
        super().__init__(rsrcmgr, pageno=pageno, laparams=laparams)
        self.result = []

    def receive_layout(self, ltpage):
        lines = {}

        def render(item):
            if isinstance(item, (LTPage, LTTextBox)):
                for child in item:
                    render(child)
            elif isinstance(item, LTTextLine):
                child_str = ''
                for child in item:
                    if isinstance(child, (LTChar, LTAnno)):
                        child_str += child.get_text()
                child_str = ' '.join(child_str.split()).strip()
                if child_str:
                    lines.setdefault(
                        (self.pageno, round(item.bbox[1], 2)), []
                    ).append((item.bbox[0], child_str))
                for child in item:
                    render(child)

        render(ltpage)
        ordered = sorted(lines.items(), key=lambda item: (item[0][0], -item[0][1]))
        self.result = [
            [text for _, text in sorted(values, key=lambda value: value[0])]
            for _, values in ordered
        ]

    def get_result(self):
        return self.result


def pdf_to_lines(file_name):
    data = []
    with open(file_name, 'rb') as stream:
        rsrcmgr = PDFResourceManager()
        device = LineConverter(rsrcmgr, laparams=LAParams())
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        for page in PDFPage.get_pages(stream):
            interpreter.process_page(page)
            data.extend(device.get_result())
    return data


def normalize_check_no(value):
    digits = re.sub(r'\D', '', str(value or ''))
    return str(int(digits)) if digits else ''


def clean_ocr_payee(text):
    text = ' '.join(str(text).split()).strip()
    text = re.sub(
        r'^PAY\s+TO\s+THE(?:\s+ORDER\s+OF)?\s*',
        '',
        text,
        flags=re.I,
    )
    text = re.sub(r'^[^A-Za-z0-9]+', '', text)
    text = re.sub(r'\s+\$.*$', '', text)
    text = re.sub(r'\s+(?:DOLLARS?|DATE)\b.*$', '', text, flags=re.I)
    text = re.sub(r'\s+\d[\d,]*\.\d{2}\s*$', '', text)
    return text.strip(' |_;:-,.')


def image_to_png_bytes(image):
    image_buffer = io.BytesIO()
    image.save(image_buffer, format='PNG')
    return image_buffer.getvalue()


def remove_horizontal_form_lines(image):
    """Remove long check form rules while preserving handwriting strokes."""
    cleaned = image.copy()
    pixels = cleaned.load()
    width, height = cleaned.size
    rows_to_clear = []
    for y in range(height):
        dark_pixels = sum(1 for x in range(width) if pixels[x, y] < 135)
        if dark_pixels >= width * 0.48:
            rows_to_clear.append(y)
    for y in rows_to_clear:
        for clear_y in range(max(0, y - 1), min(height, y + 2)):
            for x in range(width):
                pixels[x, clear_y] = 255
    return cleaned


def prepare_ocr_image(image, threshold=False, scale=1, remove_lines=False):
    """Increase handwriting contrast without changing the source PDF."""
    prepared = ImageOps.autocontrast(ImageOps.grayscale(image), cutoff=1)
    if remove_lines:
        prepared = remove_horizontal_form_lines(prepared)
    if scale > 1:
        prepared = prepared.resize(
            (prepared.width * scale, prepared.height * scale),
            Image.Resampling.LANCZOS,
        )
    if threshold:
        # A fairly light threshold retains blue/gray pen strokes.
        prepared = prepared.point(lambda value: 0 if value < 205 else 255)
    return prepared


def run_tesseract_tsv(image, page_segmentation_mode):
    """Return OCR words with positions and Tesseract confidence values."""
    if not os.path.isfile(TESSERACT_CMD):
        raise RuntimeError(f'Tesseract was not found: {TESSERACT_CMD}')

    completed = subprocess.run(
        [
            TESSERACT_CMD,
            'stdin',
            'stdout',
            '--psm',
            str(page_segmentation_mode),
            '-l',
            'eng',
            'tsv',
        ],
        input=image_to_png_bytes(image),
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=True,
    )
    decoded = completed.stdout.decode('utf-8', errors='replace')
    words = []
    for row in csv.DictReader(io.StringIO(decoded), delimiter='\t'):
        word_text = (row.get('text') or '').strip()
        if not word_text:
            continue
        try:
            confidence = float(row.get('conf', -1))
            left = int(row.get('left', 0))
            top = int(row.get('top', 0))
            width = int(row.get('width', 0))
            height = int(row.get('height', 0))
        except (TypeError, ValueError):
            continue
        words.append(
            {
                'text': word_text,
                'confidence': confidence,
                'left': left,
                'top': top,
                'width': width,
                'height': height,
            }
        )
    return words


def words_to_text_and_confidence(words):
    text = ' '.join(word['text'] for word in words)
    valid_confidences = [
        word['confidence']
        for word in words
        if word['confidence'] >= 0
    ]
    confidence = (
        sum(valid_confidences) / len(valid_confidences)
        if valid_confidences
        else 0.0
    )
    return clean_ocr_payee(text), confidence


def find_pay_to_label(words, image_width, image_height):
    """Find the printed PAY TO THE ORDER OF label on a check."""
    label_tokens = {'PAY', 'TO', 'THE', 'ORDER', 'OF'}
    eligible = []
    for word in words:
        token = re.sub(r'[^A-Z]', '', word['text'].upper())
        center_y = word['top'] + word['height'] / 2
        if (
            token in label_tokens
            and word['left'] <= image_width * 0.32
            and image_height * 0.12 <= center_y <= image_height * 0.78
        ):
            item = dict(word)
            item['token'] = token
            eligible.append(item)

    anchors = sorted(
        eligible,
        key=lambda word: (
            {'ORDER': 0, 'OF': 1, 'PAY': 2, 'TO': 3, 'THE': 4}.get(
                word['token'], 5
            ),
            word['top'],
        ),
    )
    for anchor_word in anchors:
        anchor_center_y = anchor_word['top'] + anchor_word['height'] / 2
        nearby = [
            word
            for word in eligible
            if abs(
                (word['top'] + word['height'] / 2) - anchor_center_y
            ) <= image_height * 0.20
        ]
        nearby_tokens = {word['token'] for word in nearby}
        strong_pair = (
            {'ORDER', 'OF'} <= nearby_tokens
            or {'TO', 'THE'} <= nearby_tokens
            or (
                'PAY' in nearby_tokens
                and bool(nearby_tokens.intersection({'TO', 'THE', 'ORDER'}))
            )
        )
        # Some noisy scans recognize only the short printed word "OF".
        # In the left-side label area it remains a useful vertical anchor.
        if strong_pair or anchor_word['token'] == 'OF':
            return nearby
    return []


def locate_pay_to_label(check_image):
    """OCR an enlarged left strip so the tiny printed label is detectable."""
    width, height = check_image.size
    y_offset = int(height * 0.08)
    label_region = check_image.crop(
        (0, y_offset, int(width * 0.34), int(height * 0.80))
    )
    scale = 2
    enlarged = prepare_ocr_image(label_region, scale=scale)
    words = run_tesseract_tsv(enlarged, 11)

    translated_words = []
    for word in words:
        translated = dict(word)
        translated['left'] = int(word['left'] / scale)
        translated['top'] = int(word['top'] / scale) + y_offset
        translated['width'] = max(1, int(word['width'] / scale))
        translated['height'] = max(1, int(word['height'] / scale))
        translated_words.append(translated)
    return find_pay_to_label(translated_words, width, height)


def make_payee_crop_candidates(check_image, label_words):
    """Build a precise crop plus a layout-based fallback crop."""
    width, height = check_image.size
    crop_boxes = []

    if label_words:
        label_right = max(
            word['left'] + word['width'] for word in label_words
        )
        label_top = min(word['top'] for word in label_words)
        label_bottom = max(
            word['top'] + word['height'] for word in label_words
        )
        left = max(int(width * 0.10), label_right + int(width * 0.01))
        top = max(0, label_top - int(height * 0.08))
        bottom = min(height, label_bottom + int(height * 0.01))
        # A narrow crop removes the date/amount area and is much better for
        # short payees such as COMCAST. Keep a wider crop for long names.
        crop_boxes.append((left, top, int(width * 0.52), bottom))
        crop_boxes.append((left, top, int(width * 0.72), bottom))
    else:
        # Compact image pages use shorter, wider checks and place the payee
        # lower. Older statement pages use a taller check image.
        if width / max(height, 1) >= 2.55:
            fallback_box = (
                int(width * 0.11),
                int(height * 0.37),
                int(width * 0.72),
                int(height * 0.50),
            )
        else:
            fallback_box = (
                int(width * 0.11),
                int(height * 0.24),
                int(width * 0.72),
                int(height * 0.48),
            )
        crop_boxes.append(fallback_box)

    candidates = []
    seen = set()
    for box in crop_boxes:
        left, top, right, bottom = box
        if right - left < 40 or bottom - top < 15:
            continue
        rounded = tuple(int(value) for value in box)
        if rounded in seen:
            continue
        seen.add(rounded)
        candidates.append(check_image.crop(rounded))
    return candidates


def score_payee_candidate(text, confidence):
    if not text or len(re.findall(r'[A-Za-z]', text)) < 2:
        return -1000

    alpha_count = len(re.findall(r'[A-Za-z]', text))
    digit_count = len(re.findall(r'\d', text))
    word_count = len(re.findall(r"[A-Za-z][A-Za-z'.&-]*", text))
    score = confidence + min(alpha_count, 30) * 0.45

    if 1 <= word_count <= 8:
        score += 8
    else:
        score -= abs(word_count - 5) * 2
    if 3 <= len(text) <= 60:
        score += 6
    if digit_count > alpha_count * 0.35:
        score -= 15
    if re.search(r'\b(?:LLC|INC|CORP|LTD|COMPANY|CO)\b', text, flags=re.I):
        score += 5
    if re.search(
        r'\b(?:PAY|ORDER|DOLLARS?|DATE|VATE|CHASE|MEMO|AUTHORIZED)\b',
        text,
        flags=re.I,
    ):
        score -= 30
    return score


def recognize_payee_crop(payee_image, page_segmentation_mode, threshold=False):
    prepared = prepare_ocr_image(
        payee_image,
        threshold=threshold,
        scale=3,
        remove_lines=True,
    )
    words = run_tesseract_tsv(prepared, page_segmentation_mode)
    text, confidence = words_to_text_and_confidence(words)
    return {
        'text': text,
        'confidence': confidence,
        'score': score_payee_candidate(text, confidence),
    }


def ocr_check_payee(check_image):
    # Enlarge only the left label area before locating the payee line.
    label_words = locate_pay_to_label(check_image)
    payee_crops = make_payee_crop_candidates(check_image, label_words)

    candidates = []
    for crop_index, payee_crop in enumerate(payee_crops):
        candidates.append(recognize_payee_crop(payee_crop, 7))
        # Block mode is useful for joined-up handwriting. Run it only on
        # the most precise crop so a large statement does not become too slow.
        if crop_index == 0:
            candidates.append(recognize_payee_crop(payee_crop, 6))

    best = max(candidates, key=lambda item: item['score'])

    # When the first pass is uncertain, try a binarized version that often
    # recovers faint blue/gray pen strokes.
    if best['score'] < 62 and payee_crops:
        threshold_candidate = recognize_payee_crop(
            payee_crops[0],
            6,
            threshold=True,
        )
        if threshold_candidate['score'] > best['score']:
            best = threshold_candidate

    # If a damaged label produced an unusable crop, retry the conventional
    # payee band. This is conditional so normal checks do not pay the OCR cost.
    if not best['text']:
        width, height = check_image.size
        if width / max(height, 1) >= 2.55:
            fallback_box = (
                int(width * 0.11),
                int(height * 0.37),
                int(width * 0.72),
                int(height * 0.50),
            )
        else:
            fallback_box = (
                int(width * 0.11),
                int(height * 0.37),
                int(width * 0.72),
                int(height * 0.55),
            )
        fallback_crop = check_image.crop(fallback_box)
        fallback_candidates = [
            recognize_payee_crop(fallback_crop, 6),
            recognize_payee_crop(fallback_crop, 7),
        ]
        fallback_best = max(
            fallback_candidates,
            key=lambda item: item['score'],
        )
        if fallback_best['score'] > best['score']:
            best = fallback_best

    return best['text'], best['confidence']


def extract_check_payees(file_name):
    check_payees = {}
    render_scale = 6
    render_pdf = pdfium.PdfDocument(file_name)

    try:
        with pdfplumber.open(file_name) as text_pdf:
            for page_index, page in enumerate(text_pdf.pages):
                page_text = page.extract_text() or ''
                # Chase uses both "*start*images section" and
                # "*start*imagessection" in different statement formats.
                compact_page_text = re.sub(r'\s+', '', page_text.lower())
                if '*start*imagessection' not in compact_page_text:
                    continue

                check_images = [
                    image
                    for image in page.images
                    if image.get('width', 0) >= 150
                    and image.get('height', 0) >= 70
                ]
                if not check_images:
                    continue

                check_labels = []
                for word in page.extract_words():
                    match = re.fullmatch(r'#0*(\d+)', word['text'])
                    if match:
                        check_labels.append(
                            {
                                'check_no': normalize_check_no(match.group(1)),
                                'x': (word['x0'] + word['x1']) / 2,
                                'top': word['top'],
                            }
                        )

                rendered_page = render_pdf[page_index].render(
                    scale=render_scale
                ).to_pil()

                for image_info in sorted(
                    check_images,
                    key=lambda item: (item['top'], item['x0']),
                ):
                    image_center = (image_info['x0'] + image_info['x1']) / 2
                    candidates = [
                        label
                        for label in check_labels
                        if image_info['x0'] <= label['x'] <= image_info['x1']
                        and image_info['bottom'] - 3
                        <= label['top']
                        <= image_info['bottom'] + 35
                    ]
                    if not candidates:
                        continue

                    label = min(
                        candidates,
                        key=lambda item: (
                            abs(item['top'] - image_info['bottom']),
                            abs(item['x'] - image_center),
                        ),
                    )
                    check_no = label['check_no']

                    check_image = rendered_page.crop(
                        (
                            int(image_info['x0'] * render_scale),
                            int(image_info['top'] * render_scale),
                            int(image_info['x1'] * render_scale),
                            int(image_info['bottom'] * render_scale),
                        )
                    )
                    payee, confidence = ocr_check_payee(check_image)
                    if payee:
                        check_payees[check_no] = payee
                        review_note = (
                            ' [LOW CONFIDENCE - REVIEW]'
                            if confidence < 40
                            else ''
                        )
                        print(
                            f'OCR check #{check_no} payee: {payee} '
                            f'(confidence {confidence:.0f}){review_note}'
                        )
    finally:
        render_pdf.close()

    return check_payees


def infer_statement_years(lines):
    for row in lines:
        text = ' '.join(row)
        match = MONTH_RANGE_RE.search(text)
        if match:
            start = dt.strptime(match.group(1), '%B %d, %Y')
            end = dt.strptime(match.group(2), '%B %d, %Y')
            return start.year, end.year, start, end
    today = dt.today()
    return today.year, today.year, None, None


def infer_year(mmdd, start_year, end_year):
    month = int(mmdd.split('/')[0])
    if start_year == end_year:
        return start_year
    return end_year if month <= 6 else start_year


def detect_section_from_marker_or_heading(row):
    text = ' '.join(row).strip()
    compact = ' '.join(text.upper().replace('( CONTINUED )', '(CONTINUED)').split())
    lower = compact.lower()

    if (
        'daily ending balance' in lower
        or '*start*imagessection' in lower
        or '*start*images section' in lower
        or '*start*dreportraitdisclosure' in lower
    ):
        return None

    if '*start*' in lower:
        for marker, section in SECTION_MARKERS:
            if marker in lower:
                return section

    tokens = [token.strip().upper() for token in row if token and token.strip()]
    if tokens == ['ELECTRONIC WITHDRAWALS', '(CONTINUED)']:
        return 'ELECTRONIC WITHDRAWALS'
    if tokens == ['DEPOSITS AND ADDITIONS', '(CONTINUED)']:
        return 'DEPOSITS AND ADDITIONS'
    if len(tokens) == 1 and tokens[0] in SECTION_HEADINGS:
        return SECTION_HEADINGS[tokens[0]]
    return False


def is_header_or_noise(row):
    text = ' '.join(row).strip()
    compact = ' '.join(text.split()).strip()
    upper = compact.upper()

    if not compact:
        return True
    if '*START*' in upper or '*END*' in upper:
        return True
    if upper.startswith('ACCOUNT NUMBER'):
        return True
    if upper == 'PAGE OF' or re.fullmatch(r'\d+\s+\d+', upper):
        return True
    if re.fullmatch(r'\d{9,}', upper):
        return True
    if upper.startswith('PAGE ') and ' OF ' in upper:
        return True
    if upper in {
        'DATE DESCRIPTION AMOUNT',
        'DATE DESCRIPTION AMOUNT DATE DESCRIPTION AMOUNT',
        'DATE DESCRIPTION AMOUNT CHECK NO. DESCRIPTION DATE PAID AMOUNT',
        'CHECK NO. DESCRIPTION DATE PAID AMOUNT',
        'CHECK NO. DESCRIPTION PAID',
        'DESCRIPTION CHECK NO. PAID',
        'DATE',
        'DESCRIPTION',
        'AMOUNT',
        'DATE PAID',
        'CHECK NO.',
        'CHECK NO',
        'PAID',
    }:
        return True
    if MONTH_RANGE_RE.search(compact):
        return True
    if upper.startswith('TOTAL '):
        return True
    if 'DAILY ENDING BALANCE' in upper:
        return True
    if upper.startswith('ATM & DEBIT CARD SUMMARY') or upper.startswith('ATM & DEBIT CARD TOTALS'):
        return True
    if upper.startswith('TOTAL ATM WITHDRAWALS') or upper.startswith('TOTAL CARD PURCHASES') or upper.startswith('TOTAL CARD DEPOSITS'):
        return True
    if upper.endswith('CARD 0766') or upper.endswith('CARD 0840'):
        return True
    if upper.startswith('IF YOU SEE A DESCRIPTION') or upper.startswith('* ALL OF YOUR RECENT CHECKS'):
        return True
    if upper.startswith('^ AN IMAGE OF THIS CHECK'):
        return True
    if upper.startswith('THE FEES FOR THIS ACCOUNT ARE INCLUDED'):
        return True
    if upper.startswith('IN CASE OF ERRORS'):
        return True
    if upper.startswith('SEE BOTH FRONT AND BACK IMAGES') or upper.startswith('IMAGES ACCOUNT') or upper.startswith('IMAGES (CONTINUED)'):
        return True
    if re.fullmatch(r'\d', upper):
        return True
    return False


def strip_trailing_amount(text):
    return re.sub(r'\s*\$?-?\d[\d,]*\.\d{2}\s*$', '', text).strip()


def extract_trailing_amount(text):
    match = re.search(r'(\$?-?\d[\d,]*\.\d{2})\s*$', text)
    return match.group(1) if match else None


def row_to_text(row):
    text = ' '.join(str(value).strip() for value in row if str(value).strip()).strip()
    return re.sub(r'^\d+\s+(?=\d{1,2}/\d{1,2}\b)', '', text)


def parse_date_line(text):
    text = ' '.join(text.split()).strip()
    match = re.match(r'^(\d{1,2}/\d{1,2})\s+(.+)$', text)
    if not match:
        return None

    date = match.group(1)
    rest = match.group(2).strip()
    amount = extract_trailing_amount(rest)
    description = strip_trailing_amount(rest) if amount else rest
    description = re.sub(r'^\d{1,2}/\d{1,2}\s+', '', description).strip()
    return {'date': date, 'desc': description, 'amount': amount}


def normalize_description(description):
    description = ' '.join(str(description).split()).strip()

    match = re.search(
        r'\bB/O\s+Customer\s*=\s*(.*?)(?=\s+Bene\b)',
        description,
        flags=re.I,
    )
    if match:
        return match.group(1).strip()

    match = re.search(
        r'\bCO\s+Entry\s+Descr\s*:\s*(Pseg)\b',
        description,
        flags=re.I,
    )
    if match:
        return match.group(1).strip()

    match = re.search(
        r'\bOrig\s+CO\s+Name\s*:\s*(.*?)(?=\s+Orig\s+ID\s*:)',
        description,
        flags=re.I,
    )
    if match:
        return match.group(1).strip()

    match = re.match(
        r'^(ATM\s+(?:Cash\s+Deposit|Withdrawal))\b',
        description,
        flags=re.I,
    )
    if match:
        return match.group(1).strip()

    description = re.sub(
        r'\s+Transaction\s*#\s*:\s*.*$',
        '',
        description,
        flags=re.I,
    ).strip()

    match = re.match(
        r'^(?:Quickpay\s+With\s+)?Zelle\s+Payment\s+(?:To|From)\s+(.+)$',
        description,
        flags=re.I,
    )
    if match:
        core = match.group(1).strip()
        core = re.sub(
            r'\s+(?:\d+|Jpm[A-Za-z0-9]+|Bac[A-Za-z0-9]+)$',
            '',
            core,
            flags=re.I,
        )
        return core.strip()

    match = re.match(
        r'^(?:Recurring\s+)?Card\s+Purchase'
        r'(?:\s+(?:With\s+Pin|Return))?'
        r'\s+\d{1,2}/\d{1,2}\s+(.+)$',
        description,
        flags=re.I,
    )
    if match:
        merchant = match.group(1).strip()
        merchant = re.sub(
            r'\s+Card\s+\d+\b.*$', '', merchant, flags=re.I
        ).strip()
        if '*' in merchant:
            merchant = merchant.split('*', 1)[0].strip()
        merchant = re.sub(r'\s+#\w+.*$', '', merchant, flags=re.I).strip()

        known_merchants = [
            r'Amazon\.Com',
            r'Walgreens\s+Store',
            r'The\s+Home\s+Depot',
            r'Fei\s+Long\s+Market',
            r'Target',
            r'Exxon',
        ]
        for pattern in known_merchants:
            merchant_match = re.match(rf'^({pattern})\b', merchant, flags=re.I)
            if merchant_match:
                return merchant_match.group(1).strip()

        merchant_match = re.match(
            r'^(.+?\b(?:Inc|LLC|Ltd|Corp|Corporation)\b)',
            merchant,
            flags=re.I,
        )
        if merchant_match:
            return merchant_match.group(1).strip()
        return merchant.strip()

    return description.strip()


def translate_to_csv(lines, file_order=0, account_rules=None):
    start_year, end_year, _, _ = infer_statement_years(lines)
    csv_data = []
    current_section = None
    pending = None
    order = 0

    def append_row(section, trans_date, amount, desc='', check_no=''):
        nonlocal order
        year = infer_year(trans_date, start_year, end_year)
        trans_date_obj = dt.strptime(f'{trans_date}/{year}', '%m/%d/%Y')
        order += 1
        csv_data.append(
            {
                'Section': section,
                'Type': 'Deposit' if section == 'DEPOSITS AND ADDITIONS' else 'Withdrawal',
                'Trans Date': trans_date_obj,
                'Check No': check_no,
                'Description': normalize_description(desc),
                'Account': categorize_account(desc, section, check_no, account_rules),
                'Amount': normalize_amount(amount, section),
                'Balance': '',
                '_file_order': file_order,
                '_order': order,
            }
        )

    def finalize_pending():
        nonlocal pending
        if pending and pending.get('amount'):
            append_row(
                pending['section'],
                pending['date'],
                pending['amount'],
                ' '.join(pending.get('desc_parts', [])).strip(),
                pending.get('check_no', ''),
            )
        pending = None

    for row in lines:
        if not row:
            continue
        text = row_to_text(row)
        if not text:
            continue

        section_hit = detect_section_from_marker_or_heading(row)
        if section_hit is None:
            finalize_pending()
            current_section = None
            continue
        if section_hit:
            finalize_pending()
            current_section = section_hit
            continue
        if is_header_or_noise(row):
            continue
        if current_section not in TARGET_SECTIONS:
            continue

        if current_section == 'CHECKS PAID':
            finalize_pending()
            match = re.match(
                r'^\s*(\d+)(?:\s+\*?\s*\^)?\s+'
                r'(\d{1,2}/\d{1,2})(?:\s+\d{1,2}/\d{1,2})?\s+'
                r'(\$?\d[\d,]*\.\d{2})\s*$',
                text,
            )
            if not match:
                second_match = re.match(
                    r'^\s*(\d{1,2}/\d{1,2})\s+(\d+)'
                    r'(?:\s+\*?\s*\^)?\s+\d{1,2}/\d{1,2}\s+'
                    r'(\$?\d[\d,]*\.\d{2})\s*$',
                    text,
                )
                if second_match:
                    append_row(
                        'CHECKS PAID',
                        second_match.group(1),
                        second_match.group(3),
                        '',
                        second_match.group(2),
                    )
                continue
            append_row(
                'CHECKS PAID',
                match.group(2),
                match.group(3),
                '',
                match.group(1),
            )
            continue

        if DATE_RE.match(text):
            finalize_pending()
            parsed = parse_date_line(text)
            if not parsed:
                continue
            pending = {
                'section': current_section,
                'date': parsed['date'],
                'amount': parsed['amount'],
                'desc_parts': [parsed['desc']] if parsed['desc'] else [],
                'check_no': '',
            }
            continue

        if not pending or is_header_or_noise(row):
            continue

        continuation = text
        amount = extract_trailing_amount(continuation)
        if amount and pending.get('amount') is None:
            pending['amount'] = amount
            continuation_desc = strip_trailing_amount(continuation)
            if continuation_desc:
                pending['desc_parts'].append(continuation_desc)
        elif MONEY_RE.fullmatch(continuation):
            if pending.get('amount') is None:
                pending['amount'] = continuation
        else:
            pending['desc_parts'].append(continuation)

    finalize_pending()
    return sort_and_filter(csv_data)


def extract_statement_summary(lines):
    summary = {
        'Beginning Balance': None,
        'Ending Balance': None,
        'DEPOSITS AND ADDITIONS': {'count': None, 'amount': None},
        'CHECKS PAID': {'count': None, 'amount': None},
        'ATM & DEBIT CARD WITHDRAWALS': {'count': None, 'amount': None},
        'ELECTRONIC WITHDRAWALS': {'count': None, 'amount': None},
        'OTHER WITHDRAWALS': {'count': None, 'amount': None},
        'FEES': {'count': None, 'amount': None},
    }
    for row in lines:
        text = row_to_text(row)
        match = re.search(r'Beginning Balance\s+\$?([\d,]+\.\d{2})', text, re.I)
        if match:
            summary['Beginning Balance'] = clean_money(match.group(1))
        match = re.search(
            r'Ending Balance\s+(?:\d+\s+)?\$?([\d,]+\.\d{2})',
            text,
            re.I,
        )
        if match:
            summary['Ending Balance'] = clean_money(match.group(1))

        checks = [
            ('DEPOSITS AND ADDITIONS', r'Deposits and Additions\s+(\d+)\s+([\-$\d,\.]+)'),
            ('CHECKS PAID', r'Checks Paid\s+(\d+)\s+(-?[\-$\d,\.]+)'),
            ('ATM & DEBIT CARD WITHDRAWALS', r'ATM & Debit Card Withdrawals\s+(\d+)\s+(-?[\-$\d,\.]+)'),
            ('ELECTRONIC WITHDRAWALS', r'Electronic Withdrawals\s+(\d+)\s+(-?[\-$\d,\.]+)'),
            ('OTHER WITHDRAWALS', r'Other Withdrawals\s+(\d+)\s+(-?[\-$\d,\.]+)'),
            ('FEES', r'Fees\s+(\d+)\s+(-?[\-$\d,\.]+)'),
        ]
        for section, pattern in checks:
            summary_match = re.search(pattern, text, re.I)
            if summary_match:
                summary[section]['count'] = int(summary_match.group(1))
                summary[section]['amount'] = clean_money(summary_match.group(2))
    return summary


def validate_file(file_name, rows, lines):
    summary = extract_statement_summary(lines)
    output = [f'File: {os.path.basename(file_name)}']
    total = Decimal('0.00')

    for section in [
        'DEPOSITS AND ADDITIONS',
        'CHECKS PAID',
        'ATM & DEBIT CARD WITHDRAWALS',
        'ELECTRONIC WITHDRAWALS',
        'OTHER WITHDRAWALS',
        'FEES',
    ]:
        section_rows = [row for row in rows if row['Section'] == section]
        count = len(section_rows)
        amount = sum(Decimal(row['Amount']) for row in section_rows)
        total += amount
        expected = summary.get(section, {})
        if expected.get('count') is None and count == 0:
            continue

        status = 'OK'
        if expected.get('count') is not None and expected['count'] != count:
            status = 'COUNT MISMATCH'
        if expected.get('amount') is not None and expected['amount'] != amount:
            status = 'AMOUNT MISMATCH'

        output.append(
            f'  {section}: parsed count={count}, amount={amount}; '
            f'statement count={expected.get("count")}, '
            f'amount={expected.get("amount")} => {status}'
        )

    if summary['Beginning Balance'] is not None and summary['Ending Balance'] is not None:
        computed = summary['Beginning Balance'] + total
        status = 'OK' if computed == summary['Ending Balance'] else 'MISMATCH'
        output.append(
            f'  Balance check: beginning {summary["Beginning Balance"]} '
            f'+ net {total} = {computed}; ending {summary["Ending Balance"]} '
            f'=> {status}'
        )
    return '\n'.join(output)


def main():
    global CSV_DATE_FORMAT, START_DATE, END_DATE, TESSERACT_CMD

    parser = argparse.ArgumentParser(
        description='Convert Chase statement PDFs to CSV and OCR check payees.'
    )
    parser.add_argument('--start', default=None)
    parser.add_argument('--end', default=None)
    parser.add_argument('--date-format', default=CSV_DATE_FORMAT)
    parser.add_argument(
        '--output',
        default='Extracted_Chase_Activity_Stmt_{}.csv'.format(
            dt.today().strftime('%Y%m%d')
        ),
    )
    parser.add_argument('--dir', default='.', help='Directory containing PDF files')
    parser.add_argument('--validate-log', default=None)
    parser.add_argument('--account-rules', default=None)
    parser.add_argument(
        '--tesseract',
        default=TESSERACT_CMD,
        help='Path to tesseract.exe',
    )
    args = parser.parse_args()

    CSV_DATE_FORMAT = args.date_format
    START_DATE = dt.strptime(args.start, CSV_DATE_FORMAT) if args.start else None
    END_DATE = dt.strptime(args.end, CSV_DATE_FORMAT) if args.end else None
    TESSERACT_CMD = args.tesseract
    account_rules = (
        load_account_rules(args.account_rules)
        if args.account_rules
        else ACCOUNT_RULES
    )

    input_files = sorted(
        os.path.join(args.dir, file_name)
        for file_name in os.listdir(args.dir)
        if file_name.lower().endswith('.pdf')
    )
    if not input_files:
        raise SystemExit(f'No PDF files found in: {args.dir}')

    all_rows = []
    validation_blocks = []
    print('Started scanning Chase PDFs:')

    for file_order, file_name in enumerate(input_files):
        lines = pdf_to_lines(file_name)
        rows = translate_to_csv(
            lines,
            file_order=file_order,
            account_rules=account_rules,
        )

        check_payees = extract_check_payees(file_name)
        for row in rows:
            check_no = normalize_check_no(row.get('Check No', ''))
            payee = check_payees.get(check_no)
            if payee:
                row['Description'] = payee

        all_rows.extend(rows)
        validation_blocks.append(validate_file(file_name, rows, lines))
        print(f'Parsed {len(rows)} rows from file: {file_name}')

    all_rows = sort_and_filter(all_rows)
    with open(args.output, 'w', newline='', encoding='utf-8-sig') as stream:
        writer = csv.DictWriter(stream, fieldnames=CSV_HEADERS)
        writer.writeheader()
        for row in all_rows:
            output = {header: row.get(header, '') for header in CSV_HEADERS}
            output['Trans Date'] = dt.strftime(
                output['Trans Date'], CSV_DATE_FORMAT
            )

            amount = Decimal(row.get('Amount', '0.00'))
            if amount < 0:
                output['Amount'] = ''
                output['Out Amount'] = fmt_decimal(-amount)
            else:
                output['Amount'] = fmt_decimal(amount)
                output['Out Amount'] = ''
            writer.writerow(output)

    if args.validate_log:
        with open(args.validate_log, 'w', encoding='utf-8') as stream:
            stream.write('\n\n'.join(validation_blocks))

    print(f'CSV file generated with {len(all_rows)} rows: {args.output}')
    if args.validate_log:
        print(f'Validation log generated: {args.validate_log}')


if __name__ == '__main__':
    main()
