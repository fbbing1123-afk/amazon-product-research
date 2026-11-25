"""直接读取商品数据源生成集合链接，无需访问卖家精灵链接生成器。"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
import tempfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'modules'))
from shared.asin_collection_links import collection_from_workbook


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--workbook', type=Path, required=True)
    parser.add_argument('--site', help='例如US、JP、DE；省略时按数据源商品链接识别')
    parser.add_argument('--sheet', help='有多个商品源时显式指定页名')
    parser.add_argument('--output-dir', type=Path, required=True)
    args = parser.parse_args()
    result = collection_from_workbook(args.workbook, args.site, args.sheet)
    args.output_dir.mkdir(parents=True, exist_ok=True)
    directory = Path(tempfile.mkdtemp(prefix='asin-link-', dir=args.output_dir)).resolve()
    (directory / 'collection-link.txt').write_text(result['url'] + '\n', encoding='utf-8')
    (directory / 'collection-link.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'directory': str(directory), 'site': result['site'], 'asin_count': result['asin_count'],
                      'duplicate_rows': result['duplicate_rows'], 'url': result['url']}, ensure_ascii=False))


if __name__ == '__main__':
    main()
