"""(주)프레시라인 고객사 재고 조회 페이지 데이터 생성기.

사용법:
    python tools/build_portal.py <GSFM_정기업무 폴더> <설정 파일(json)> <출력 폴더: docs/portal/data>

- 설정 파일(고객사 목록·아이디·비밀번호·엑셀 위치)은 사용자 컴퓨터에만 두고 저장소에 넣지 않는다.
  형식: [{"id": 아이디, "pw": 비밀번호, "name": 화면 이름, "kind": 양식, "start": 시작일, "file": 엑셀 경로, ...}, ...]
- 고객사마다 데이터를 비밀번호로 암호화(AES-GCM, PBKDF2-SHA256)해 <아이디 해시>.json 으로 저장한다.
- 재고 계산은 각 엑셀의 입출고현황 수식과 같은 방식: 기초재고(시작일 기준) + 시작일 이후 입고 − 출고 (+ 조정).
"""
import sys, os, json, base64, hashlib, secrets, datetime as dt
from zoneinfo import ZoneInfo
import openpyxl
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes

ITER = 310000
LOW_BOX = 5      # 품절 예정: 박스 환산 5박스 미만 (박스로 못 바꾸는 품목은 기본 단위 5 미만)
TH_WEEKS = 2     # 곧 소진: 최근 4주 평균 출고로 2주 안에 소진 예상
UNIT = {'BOX': '박스', 'PK': '팩', 'EA': '개'}


def id_hash(cid):
    return hashlib.sha256(('freshline-portal:' + cid.strip()).encode('utf-8')).hexdigest()[:24]

def d(v):
    if isinstance(v, dt.datetime): return v.date().isoformat()
    if isinstance(v, dt.date): return v.isoformat()
    return None

def num(v):
    try: return float(v or 0)
    except (TypeError, ValueError): return 0.0

def s(v):
    return str(v).strip() if v is not None and str(v).strip() else ''

def unit_of(v, default):
    return UNIT.get(s(v).upper(), default)


# ---------- 양식별 읽기 ----------
def load_standard(wb, cfg):
    """표준 양식: 입출고현황(B 품목, [C 단위], 기초재고 열) / 입고내역 A~E / 출고내역 A 일자, B 브랜드, C 점포, D 품목, E 규격, F 수량, G 비고"""
    st = wb['입출고현황']
    name_col, unit_col, open_col = 2, cfg.get('unit_col'), cfg['open_col']
    items = []
    for r in range(8, st.max_row + 1):
        n = s(st.cell(r, name_col).value)
        if not n: break
        if n in ('합계', '소계'): continue
        u = unit_of(st.cell(r, unit_col).value, None) if unit_col else cfg.get('unit', '박스')
        if unit_col and u is None: continue          # 단위 없는 부자재 등은 제외
        items.append({'품목': n, '단위': u, '입수': None, '기초': num(st.cell(r, open_col).value)})
    rows = []
    for r in wb['입고내역'].iter_rows(min_row=2, max_col=5, values_only=True):
        if not d(r[0]) or not s(r[1]): continue
        rows.append({'일자': d(r[0]), '구분': '입고', '점포': '', '품목': s(r[1]), '수량': num(r[3]), '단위': unit_of(r[2], cfg.get('unit', '박스')), '비고': s(r[4]) if len(r) > 4 else ''})
    for r in wb['출고내역'].iter_rows(min_row=2, max_col=7, values_only=True):
        if not d(r[0]) or not s(r[3]): continue
        rows.append({'일자': d(r[0]), '구분': '출고', '점포': s(r[2]), '품목': s(r[3]), '수량': num(r[5]), '단위': unit_of(r[4], cfg.get('unit', '박스')), '비고': s(r[6]) if len(r) > 6 else ''})
    return items, rows

def load_pack_info(wb, cfg):
    st, pi = wb['입출고현황'], wb['품목정보']
    per = {}
    for r in range(5, pi.max_row + 1):
        n = s(pi.cell(r, 3).value)
        if n: per[n] = (int(num(pi.cell(r, 7).value)) or None, s(pi.cell(r, 9).value) == '운영 중단')
    items = []
    for r in range(8, st.max_row + 1):
        n = s(st.cell(r, 2).value)
        if not n or n == '합계': break
        p, stopped = per.get(n, (None, False))
        if stopped: continue
        items.append({'품목': n, '단위': '팩', '입수': p, '기초': num(st.cell(r, 5).value)})
    rows = []
    for r in wb['입고내역'].iter_rows(min_row=2, max_col=5, values_only=True):
        if not d(r[0]) or not s(r[1]): continue
        rows.append({'일자': d(r[0]), '구분': '입고', '점포': '', '품목': s(r[1]), '수량': num(r[3]), '단위': '팩', '비고': s(r[4])})
    for r in wb['출고내역'].iter_rows(min_row=2, max_col=6, values_only=True):
        if not d(r[0]) or not s(r[2]): continue
        rows.append({'일자': d(r[0]), '구분': '출고', '점포': s(r[1]), '품목': s(r[2]), '수량': num(r[4]), '단위': '팩', '비고': s(r[5])})
    return items, rows

def load_brand_pack(wb, cfg):
    st, pi = wb['입출고현황'], wb['품목정보']
    per = {}
    for r in range(6, pi.max_row + 1):
        n = s(pi.cell(r, 4).value)
        if not n or n == '매장명': break
        per[n] = int(num(pi.cell(r, 6).value)) or None
    items = []
    for r in range(8, st.max_row + 1):
        b, n = s(st.cell(r, 2).value), s(st.cell(r, 3).value)
        if b == '합계' or (not b and not n): break
        if not n or n == '소계': continue
        items.append({'품목': n, '브랜드': b, '단위': '팩', '입수': per.get(n), '기초': num(st.cell(r, 6).value)})
    rows = []
    for r in wb['입고내역'].iter_rows(min_row=2, max_col=5, values_only=True):
        if not d(r[0]) or not s(r[1]): continue
        rows.append({'일자': d(r[0]), '구분': '입고', '점포': '', '품목': s(r[1]), '수량': num(r[3]), '단위': '팩', '비고': s(r[4])})
    for r in wb['출고내역'].iter_rows(min_row=2, max_col=6, values_only=True):
        if not d(r[0]) or not s(r[2]): continue
        rows.append({'일자': d(r[0]), '구분': '출고', '점포': s(r[1]), '품목': s(r[2]), '수량': num(r[4]), '단위': '팩', '비고': s(r[5])})
    return items, rows

def load_pack6(wb, cfg):
    st = wb['입출고현황']
    per = {}
    if '품목정보' in wb.sheetnames:                     # 박스입수: 품목정보 E열 ("6팩")
        pi = wb['품목정보']
        for r in range(5, pi.max_row + 1):
            n = s(pi.cell(r, 3).value)
            digits = ''.join(ch for ch in s(pi.cell(r, 5).value) if ch.isdigit())
            if n and digits: per[n] = int(digits)
    items = []
    for r in range(7, st.max_row + 1):
        n = s(st.cell(r, 2).value)
        if not n or n == '합계': break
        items.append({'품목': n, '단위': '팩', '입수': per.get(n, 6), '기초': num(st.cell(r, 8).value)})
    rows = []
    for r in wb['입고내역'].iter_rows(min_row=2, max_col=5, values_only=True):
        if not d(r[0]) or not s(r[1]): continue
        rows.append({'일자': d(r[0]), '구분': '입고', '점포': '', '품목': s(r[1]), '수량': num(r[3]), '단위': '팩', '비고': s(r[4])})
    for r in wb['출고내역'].iter_rows(min_row=2, max_col=7, values_only=True):
        if not d(r[0]) or not s(r[3]): continue
        rows.append({'일자': d(r[0]), '구분': '출고', '점포': s(r[2]), '품목': s(r[3]), '수량': num(r[5]), '단위': '팩', '비고': s(r[6])})
    return items, rows

def load_single_ledger(wb, cfg):
    st = wb['입출고현황']
    items = []
    for r in range(6, st.max_row + 1):
        n = s(st.cell(r, 3).value)
        if not n: continue
        if n == '합계': break
        p = st.cell(r, 4).value
        items.append({'품목': n, '단위': '개', '입수': int(p) if isinstance(p, (int, float)) and p else None, '기초': num(st.cell(r, 7).value)})
    rows = []
    for r in wb['수불내역'].iter_rows(min_row=2, max_col=6, values_only=True):
        if not d(r[0]) or not s(r[3]) or s(r[1]) not in ('입고', '출고', '조정'): continue
        rows.append({'일자': d(r[0]), '구분': s(r[1]), '점포': '', '품목': s(r[3]), '수량': num(r[4]), '단위': '개', '비고': s(r[5])})
    return items, rows

LOADERS = {'standard': load_standard, 'pack_info': load_pack_info, 'brand_pack': load_brand_pack, 'pack6': load_pack6, 'single_ledger': load_single_ledger}


# ---------- 계산 ----------
def build(cfg, base_dir):
    wb = openpyxl.load_workbook(os.path.join(base_dir, cfg['file']), data_only=True)
    items, rows = LOADERS[cfg['kind']](wb, cfg)
    nb = cfg.get('no_box')                     # 박스 환산 없이 낱개로만 표시: true(전체) 또는 품목 목록
    if nb:
        for it in items:
            if nb is True or it['품목'] in nb: it['입수'] = None
    start = cfg['start']
    rows.sort(key=lambda x: (x['일자'], {'입고': 0, '조정': 1, '출고': 2}[x['구분']]))
    asof = max(r['일자'] for r in rows) if rows else start
    end = dt.date.fromisoformat(asof)
    from4 = (end - dt.timedelta(days=27)).isoformat()
    waste = ('폐기', '파손')
    stock = []
    hide = set(cfg.get('hide_items') or [])   # 운영 중단 품목: 재고표에서만 빼고 입출고 내역은 그대로 둔다
    for it in items:
        if it['품목'] in hide: continue
        n = it['품목']
        mine = [r for r in rows if r['품목'] == n and r['일자'] >= start]
        q = it['기초'] + sum(r['수량'] for r in mine if r['구분'] == '입고') - sum(r['수량'] for r in mine if r['구분'] == '출고') + sum(r['수량'] for r in mine if r['구분'] == '조정')
        q = round(q, 2)
        wk = sum(r['수량'] for r in mine if r['구분'] == '출고' and from4 <= r['일자'] <= asof and not any(w in r['비고'] for w in waste)) / 4
        weeks = q / wk if wk > 0 else None
        boxes = q if it['단위'] == '박스' else (q // it['입수'] if it['입수'] else None)
        low_soon = cfg.get('low_soon') and boxes is not None and boxes < LOW_BOX   # 설정에 low_soon 이 있으면 5박스 미만도 '곧 소진'
        status = '품절' if q <= 0 else ('곧 소진' if low_soon or (weeks is not None and weeks < TH_WEEKS) else '여유')
        stock.append({'품목': n, '브랜드': it.get('브랜드', ''), '수량': q, '단위': it['단위'], '입수': it['입수'], '소진주': weeks, '상태': status})
    return asof, stock, rows

def encrypt(payload, password):
    salt, iv = secrets.token_bytes(16), secrets.token_bytes(12)
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITER).derive(password.encode('utf-8'))
    ct = AESGCM(key).encrypt(iv, json.dumps(payload, ensure_ascii=False).encode('utf-8'), None)
    b = lambda x: base64.b64encode(x).decode()
    return {'v': 2, 'iter': ITER, 'salt': b(salt), 'iv': b(iv), 'ct': b(ct)}

def main(base_dir, cfg_file, out_dir, only=None):
    clients = json.load(open(cfg_file, encoding='utf-8'))
    pws = {c['id']: c.get('pw') for c in clients if c.get('pw')}
    checked = dt.datetime.now(ZoneInfo('Asia/Seoul')).strftime('%Y-%m-%d %H:%M')
    os.makedirs(out_dir, exist_ok=True)
    results = []
    for cfg in clients:
        if only and cfg['id'] not in only: continue
        if cfg['id'] not in pws:
            results.append((cfg['id'], '비밀번호 없음 - 건너뜀')); continue
        try:
            asof, stock, rows = build(cfg, base_dir)
        except Exception as e:  # 한 고객사 파일이 깨져도 나머지는 계속
            results.append((cfg['id'], f'실패: {e}')); continue
        payload = {'name': cfg['name'], 'asof': asof, 'checked': checked, 'low': LOW_BOX, 'stock': stock, 'rows': rows}
        json.dump(encrypt(payload, pws[cfg['id']]), open(os.path.join(out_dir, id_hash(cfg['id']) + '.json'), 'w'))
        results.append((cfg['id'], asof, [(x['품목'], x['수량'], x['단위'], x['상태']) for x in stock]))
    for r in results: print(r)
    return results

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4:] or None)
