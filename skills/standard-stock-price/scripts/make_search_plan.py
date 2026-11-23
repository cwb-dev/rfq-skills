import argparse
import json
import re
from pathlib import Path
from urllib.parse import quote

EQUIVALENTS = {
    'DIN912': ['DIN912', 'ISO4762', 'GB70.1', 'GB70-76', 'GB70-85'],
    'DIN933': ['DIN933', 'GB5783', 'ISO4017', '特30'],
    'DIN934': ['DIN934', 'GB6170', 'ISO4032'],
}


def normalized_standard(value):
    key = re.sub(r'[^A-Z0-9]', '', value.upper())
    return key


def substitute_grade(item):
    text = str(item.get('material', '') + ' ' + item.get('grade', ''))
    if re.search(r'(?<!\d)5(?:\.8)?(?:级)?(?!\d)', text):
        return '4.8 Q235'
    if re.search(r'(?:6\.8|6级|8\.8|8级|10\.9|10级|12\.9|12级)', text):
        return '8.8' if '912' in normalized_standard(item['standard']) or '933' in normalized_standard(item['standard']) else '8级'
    return ''


def main():
    parser = argparse.ArgumentParser(description='Generate GPYH/Taodingding search URLs for standard fasteners.')
    parser.add_argument('data')
    parser.add_argument('output')
    parser.add_argument('--no-substitutes', action='store_true')
    args = parser.parse_args()
    data = json.loads(Path(args.data).read_text(encoding='utf-8'))
    queries = []
    for item in data['items']:
        standard = normalized_standard(item['standard'])
        size = item.get('spec') or item.get('size') or item['description']
        finish = item.get('finish', '')
        grade = item.get('grade', '')
        terms = [size]
        if grade:
            terms.append(grade)
        if finish and '未注明' not in finish:
            terms.append(finish)
        standards = EQUIVALENTS.get(standard, [item['standard']])
        variants = [(s, ' '.join(terms), '原标准/等效') for s in standards]
        if not args.no_substitutes:
            substitute = substitute_grade(item)
            if substitute:
                variants.append((standards[0], ' '.join(terms + [substitute]), '等级替代'))
                if standard == 'DIN933':
                    variants.append(('特30', ' '.join(terms + [substitute]), '等级替代'))
        for std, query, kind in variants:
            for platform in ('工品一号', '淘钉钉'):
                if platform == '工品一号':
                    url = 'https://www.gpyh.com/quickbuy/goodsSearchPage?isGoodsFilterFlg=no&searchKey=' + quote(query) + '&searchTab=jgj&isDirectory=false&provinceId=861&cityId=3524'
                else:
                    url = 'https://www.taodingding.com/searchList?search=' + quote(query)
                queries.append({'line': item['line'], 'platform': platform, 'standard': std, 'query': query, 'kind': kind, 'url': url})
    Path(args.output).write_text(json.dumps({'queries': queries}, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'{len(queries)} queries -> {args.output}')


if __name__ == '__main__':
    main()
