import openpyxl, io, sys, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
wb = openpyxl.load_workbook('📚 논술 독서 감상 기록(응답) (1).xlsx')
ws = wb.active

def clean(val):
    if val is None: return ''
    s = str(val).strip()
    return '' if s in ['.', 'None', 'nan'] else s

def normalize_title(title):
    return re.sub(r'[\s\-·,.:!?\u300c\u300d\u300e\u300f()\[\]]', '', title).lower()

def josa_reul(text):
    if not text: return '을'
    code = ord(text[-1]) - 0xAC00
    if 0 <= code < 11172:
        return '을' if (code % 28) != 0 else '를'
    return '을'

def josa_iran(text):
    if not text: return '이라는'
    code = ord(text[-1]) - 0xAC00
    if 0 <= code < 11172:
        return '이라는' if (code % 28) != 0 else '라는'
    return '이라는'

def fit_utf8(text, max_bytes):
    encoded = text.encode('utf-8')
    if len(encoded) <= max_bytes: return text
    trunc = encoded[:max_bytes]
    while True:
        try: return trunc.decode('utf-8')
        except: trunc = trunc[:-1]

def clean_inner_quotes(text):
    # 내부 따옴표 제거 또는 괄호로 치환
    return text.replace("'", '').replace('\u2018', '').replace('\u2019', '').replace('"', '')

def shorten_quote(text, max_bytes=90):
    text = re.sub(r'\s+', ' ', text).strip()
    text = clean_inner_quotes(text)
    # 첫 문장
    first = re.split(r'\.\s+|\n', text)[0].strip().rstrip('.')
    if len(first.encode('utf-8')) <= max_bytes:
        return re.sub(r'[,\s]+$', '', first).strip()
    # 쉼표 앞에서 자르기
    for m in re.finditer(r'[,，]', first):
        cand = first[:m.start()].strip()
        if len(cand.encode('utf-8')) <= max_bytes and len(cand) >= 8:
            return cand
    r = fit_utf8(first, max_bytes)
    return re.sub(r'[\s,.]+$', '', r).strip()

def split_to_sentences(text):
    text = re.sub(r'\s+', ' ', text).strip()
    # 명시적 마침표(다. 요. 음. 등) 뒤에서만 분리 — 따옴표 내부 오분리 방지
    parts = re.split(r'(?<=[다요음임함됨겠])\. (?=[가-힣A-Z"\'])', text)
    result = []
    for p in parts:
        p = p.strip().rstrip('.')
        if len(p) > 3:
            result.append(p)
    return result if result else [text.strip()]

def extract_sentences_with_period(text, max_bytes):
    sents = split_to_sentences(text)
    result_parts = []
    used_bytes = 0
    for s in sents:
        s = s.strip().rstrip('.')
        candidate_bytes = len((s + '. ').encode('utf-8'))
        if used_bytes + candidate_bytes <= max_bytes:
            result_parts.append(s)
            used_bytes += candidate_bytes
        else:
            if not result_parts:
                result_parts.append(fit_utf8(s, max_bytes).rstrip(',.'))
            break
    return '. '.join(result_parts)

# 느낌/반응 술어 — 이 술어를 포함한 문장을 찾는 용도 (선행 조사 불필요)
FEELING_VERBS = re.compile(
    r'느꼈|깨달았|깨닫게 되었|알게 되었|이해하게 되었|느낄 수 있었|알 수 있었|'
    r'인상 깊었|인상깊었|흥미로웠|신기했|놀라웠|감동적이었|뜻깊었|재미있었|'
    r'의문이 생겼|관심이 생겼|마음이 들었|생각이 들었|공감이 갔|'
    r'좋았다|배웠다|읽으면서 뿌듯|기억에 남았|감동받았'
)

def word_fit(text, max_bytes):
    """공백 경계 기준으로 max_bytes 이내에서 자르기"""
    if len(text.encode('utf-8')) <= max_bytes:
        return text
    raw = fit_utf8(text, max_bytes)
    idx = raw.rfind(' ')
    if idx > max(5, len(raw) // 2):
        raw = raw[:idx]
    return raw.rstrip('., ')

def phrase_from_sent(sent, max_bytes):
    """문장에서 max_bytes 이내의 핵심 구절 추출 (문장 앞부분 기준)"""
    sent = sent.strip().rstrip('.')
    sent = re.sub(r'^(그리고|또한|하지만|특히|또|그래서|이를 통해|이번에|오늘|저번에|이전에)\s*', '', sent).strip()
    sent = clean_inner_quotes(sent)
    if len(sent.encode('utf-8')) <= max_bytes:
        return sent
    # 쉼표 앞에서 자르기
    for mc in re.finditer(r'[,，]', sent):
        cand = sent[:mc.start()].strip()
        if 5 <= len(cand.encode('utf-8')) <= max_bytes:
            return cand
    return word_fit(sent, max_bytes)

def clause_with_feeling(sent, max_bytes):
    """느낌 동사를 포함한 절 추출 (쉼표/고 기준 분절)"""
    # 쉼표로 분절
    clauses = [c.strip() for c in re.split(r'[,，]', sent) if c.strip()]
    # 느낌 동사 포함 절 찾기
    for i, clause in enumerate(clauses):
        if FEELING_VERBS.search(clause):
            phrase = clean_inner_quotes(clause.rstrip('.'))
            phrase = re.sub(r'^(그리고|또한|하지만|특히|또|그래서)\s*', '', phrase).strip()
            if len(phrase.encode('utf-8')) <= max_bytes and len(phrase) >= 5:
                return phrase
            # 절이 너무 길면 '고' 기준으로 재분절
            sub_clauses = re.split(r'(?<=고) (?=[가-힣])', phrase)
            for sc in reversed(sub_clauses):
                sc = sc.strip()
                if FEELING_VERBS.search(sc) and len(sc.encode('utf-8')) <= max_bytes and len(sc) >= 5:
                    return sc
            # 그래도 길면 앞 절 반환
            if i > 0:
                prev = clean_inner_quotes(clauses[i-1].rstrip('.').strip())
                if len(prev.encode('utf-8')) <= max_bytes and len(prev) >= 5:
                    return prev
            return word_fit(phrase, max_bytes)
    return None

def extract_key_feeling(review, max_bytes=90):
    text = re.sub(r'\s+', ' ', review).strip()
    sents = split_to_sentences(text)

    # 1) 느낌 술어를 포함한 문장 → 그 절을 추출
    for sent in reversed(sents):
        if FEELING_VERBS.search(sent):
            phrase = clause_with_feeling(sent, max_bytes)
            if phrase and len(phrase) >= 5:
                return phrase
            # clause 추출 실패시 문장 앞부분
            phrase = phrase_from_sent(sent, max_bytes)
            if phrase and len(phrase) >= 5:
                return phrase

    # 2) 1인칭 문장
    for sent in reversed(sents):
        if re.match(r'^(나는|저는|나도|내가)', sent.strip()):
            phrase = phrase_from_sent(sent, max_bytes)
            if phrase and len(phrase) >= 5:
                return phrase

    # 3) 주관적 술어로 끝나는 문장
    subj_end = re.compile(r'(?:것 같다|싶다|보였다|느껴졌다|수 있었다|있을 것이다|되었다)$')
    for sent in reversed(sents):
        if subj_end.search(sent.strip()):
            phrase = phrase_from_sent(sent, max_bytes)
            if phrase and len(phrase) >= 5:
                return phrase

    # 4) 최후 수단
    for sent in reversed(sents):
        if len(sent.strip()) < 5: continue
        if re.match(r'^(오늘|이번|이런|그리고|또한|따라서|결론)', sent.strip()): continue
        phrase = phrase_from_sent(sent, max_bytes)
        if phrase and len(phrase) >= 5:
            return phrase

    return phrase_from_sent(sents[0], max_bytes)

CLOSING = '매주 책을 읽고 독서감상을 남기며 문해력이 상승.'
CLOSING_BYTES = len(CLOSING.encode('utf-8'))

def make_record(book, author, review, sentence, budget):
    header = f"'{book}({author})'을 읽고"

    # 기억문장
    sent_part = ''
    if sentence:
        short = shorten_quote(sentence, 90)
        if short:
            iran = josa_iran(short)
            sent_part = f" '{short}'{iran} 문장이 기억에 남는다고 하고"

    # 핵심감상
    key = extract_key_feeling(review, 85)
    p_key = josa_reul(key)
    feeling_part = f" '{key}'{p_key} 느낌."

    # 감상 본문에 쓸 수 있는 바이트
    fixed = (len(header.encode('utf-8')) + len(sent_part.encode('utf-8'))
             + len(feeling_part.encode('utf-8')) + 2)
    review_budget = max(60, budget - fixed)

    review_text = extract_sentences_with_period(review, review_budget)
    review_text = review_text.rstrip('.').rstrip(',').strip()

    return f"{header}{sent_part} {review_text}{feeling_part}"

# ── 데이터 수집 ──
students = {}

for row in ws.iter_rows(min_row=2, values_only=True):
    ban = clean(row[1])
    hakbun = clean(row[2])
    name = clean(row[3])
    book = clean(row[4])
    author = clean(row[5])
    review = clean(row[6])
    sentence = clean(row[7])

    if not book or not review:
        continue

    try: hnum = int(float(hakbun))
    except: hnum = str(hakbun)

    if hnum not in students:
        students[hnum] = {'ban': ban, 'name': name, 'books': {}}
    if ban and not students[hnum]['ban']:
        students[hnum]['ban'] = ban

    nk = normalize_title(book)
    if nk not in students[hnum]['books']:
        students[hnum]['books'][nk] = {'title': book, 'author': author,
                                        'review': review, 'sentence': sentence}
    else:
        ex = students[hnum]['books'][nk]
        if len(review) > len(ex['review']): ex['review'] = review
        if len(sentence) > len(ex['sentence']): ex['sentence'] = sentence
        if len(book) > len(ex['title']): ex['title'] = book; ex['author'] = author

ban_order = {'A반': 0, 'B반': 1, 'C반': 2, 'D반': 3}
sorted_students = sorted(students.items(),
    key=lambda x: (ban_order.get(x[1]['ban'], 99), str(x[0])))

MAX = 1500
output_lines = []
current_ban = None

for hnum, info in sorted_students:
    ban = info['ban'] or '미분류'
    if ban != current_ban:
        if current_ban is not None: output_lines.append('')
        output_lines.append(f'■ {ban}')
        output_lines.append('')
        current_ban = ban

    books = list(info['books'].values())
    usable = MAX - CLOSING_BYTES - 1
    per_budget = usable // len(books)

    parts = [make_record(b['title'], b['author'], b['review'], b['sentence'], per_budget)
             for b in books]
    combined = ' '.join(parts) + ' ' + CLOSING
    if len(combined.encode('utf-8')) > MAX:
        combined = fit_utf8(' '.join(parts), MAX - CLOSING_BYTES - 5) + '... ' + CLOSING

    output_lines.append(f'{hnum} {info["name"]}\t{combined}')

with open('논술_독서감상기록.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(output_lines))

over = [l for l in output_lines if '\t' in l and len(l.split('\t')[1].encode('utf-8')) > MAX]
print(f'완료! {len(students)}명 | 초과: {len(over)}개')
print()
shown = 0
for line in output_lines:
    if '\t' in line and shown < 8:
        hnum, text = line.split('\t', 1)
        print(f'[{hnum}] {len(text.encode("utf-8"))}B')
        print(text)
        print()
        shown += 1
