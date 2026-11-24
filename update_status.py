#!/usr/bin/env python3
"""西多摩感染症サイネージ用JSON更新の雛形。
公式データ取得部は、東京都WEB感染症発生動向のダウンロード仕様を確定後に実装してください。
取得失敗時に既存 data/status.json を壊さない設計です。
"""
from pathlib import Path
import json, os, tempfile
OUT=Path('data/status.json')
def validate(d):
    for k in ('influenza','covid19'):
        if k not in d or not isinstance(d[k].get('value'),(int,float)):
            raise ValueError(f'invalid {k}')
def atomic_write(d):
    validate(d); OUT.parent.mkdir(parents=True,exist_ok=True)
    fd,tmp=tempfile.mkstemp(prefix='status-',suffix='.json',dir=OUT.parent)
    try:
        with os.fdopen(fd,'w',encoding='utf-8') as f: json.dump(d,f,ensure_ascii=False,indent=2)
        json.loads(Path(tmp).read_text(encoding='utf-8'))
        os.replace(tmp,OUT)
    finally:
        if os.path.exists(tmp): os.unlink(tmp)
if __name__=='__main__':
    raise SystemExit('公式データ取得URL・形式の確定後に取得処理を接続してください。既存JSONは変更していません。')
