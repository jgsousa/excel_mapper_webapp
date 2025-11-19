
import io
from pathlib import Path
from typing import List, Dict, Any

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, FileResponse

from core.logic import (
    load_env,
    read_target_headers,
    extract_keys,
    read_source_headers,
    run_mapping_for_pair,
    get_mapping_pair,
    set_mapping_pair,
    delete_mapping_pair,
)
from core.transfer import transfer_data

import json

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / 'app_data'
SOURCES_DIR = DATA_DIR / 'Sources'
TARGET_PATH = DATA_DIR / 'Target.xlsx'
TARGET_HEADERS_PATH = DATA_DIR / 'target_headers.json'
KEYS_PATH = DATA_DIR / 'keys.json'
HEADERS_PATH = DATA_DIR / 'headers.json'  # per-source headers cache
MAPPING_PATH = DATA_DIR / 'mapping.json'
CONFIG_PATH = DATA_DIR / 'config.json'

app = FastAPI(title='Excel Mapper API', version='1.1')
app.add_middleware(
    CORSMiddleware,
    allow_origins=['*'],
    allow_credentials=True,
    allow_methods=['*'],
    allow_headers=['*'],
)

load_env()
DATA_DIR.mkdir(parents=True, exist_ok=True)
SOURCES_DIR.mkdir(parents=True, exist_ok=True)


def _read_json(path: Path, default):
    if not path.exists():
        return default
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def _write_json(path: Path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


@app.post('/api/target/upload')
async def upload_target(file: UploadFile = File(...)):
    if not file.filename.lower().endswith('.xlsx'):
        raise HTTPException(status_code=400, detail='Please upload an .xlsx file')

    content = await file.read()
    TARGET_PATH.write_bytes(content)

    headers_by_sheet = read_target_headers(TARGET_PATH)
    keys_by_sheet = extract_keys(TARGET_PATH)

    _write_json(TARGET_HEADERS_PATH, headers_by_sheet)
    _write_json(KEYS_PATH, keys_by_sheet)

    return {
        'message': 'Target uploaded and analyzed',
        'sheets': list(headers_by_sheet.keys()),
        'headers': headers_by_sheet,
        'keys': keys_by_sheet,
    }


@app.get('/api/target/info')
async def target_info():
    if not TARGET_PATH.exists():
        raise HTTPException(status_code=404, detail='Target.xlsx not uploaded yet')
    return {
        'sheets': list(_read_json(TARGET_HEADERS_PATH, {}).keys()),
        'headers': _read_json(TARGET_HEADERS_PATH, {}),
        'keys': _read_json(KEYS_PATH, {}),
    }


@app.get('/api/target/download')
async def target_download():
    if not TARGET_PATH.exists():
        raise HTTPException(status_code=404, detail='Target.xlsx not available')
    return FileResponse(path=str(TARGET_PATH), filename='Target.xlsx', media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.post('/api/source/upload')
async def upload_source(file: UploadFile = File(...)):
    if not file.filename.lower().endswith('.xlsx'):
        raise HTTPException(status_code=400, detail='Please upload an .xlsx file')
    content = await file.read()
    dest = SOURCES_DIR / file.filename
    dest.write_bytes(content)

    try:
        hdrs = read_source_headers(dest)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

    cache = _read_json(HEADERS_PATH, {})
    cache[file.filename] = hdrs
    _write_json(HEADERS_PATH, cache)

    return {'message': 'Source uploaded', 'filename': file.filename, 'headers': hdrs}


@app.get('/api/sources')
async def list_sources():
    items = []
    cache = _read_json(HEADERS_PATH, {})
    for p in SOURCES_DIR.glob('*.xlsx'):
        items.append({'filename': p.name, 'headers': cache.get(p.name, [])})
    return {'sources': items}


@app.post('/api/map/run')
async def run_mapping(payload: Dict[str, Any]):
    filename = payload.get('filename')
    sheet = payload.get('sheet')
    if not filename or not sheet:
        raise HTTPException(status_code=400, detail='filename and sheet are required')

    headers_by_sheet = _read_json(TARGET_HEADERS_PATH, {})
    if not headers_by_sheet:
        raise HTTPException(status_code=400, detail='Target.xlsx not analyzed yet')

    try:
        mapped = run_mapping_for_pair(
            sources_dir=SOURCES_DIR,
            filename=filename,
            sheet=sheet,
            target_headers_by_sheet=headers_by_sheet,
            mapping_path=MAPPING_PATH,
            headers_cache_path=HEADERS_PATH,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    return {'key': f'{filename}::{sheet}', 'mapping': mapped}


@app.get('/api/map/{filename}/{sheet}')
async def get_mapping(filename: str, sheet: str):
    mapping = get_mapping_pair(MAPPING_PATH, filename, sheet)
    return {'key': f'{filename}::{sheet}', 'mapping': mapping}


@app.post('/api/map/{filename}/{sheet}')
async def set_mapping(filename: str, sheet: str, payload: Dict[str, Any]):
    rows = payload.get('mapping', [])
    if not isinstance(rows, list):
        raise HTTPException(status_code=400, detail='mapping must be a list')
    set_mapping_pair(MAPPING_PATH, filename, sheet, rows)
    return {'message': 'mapping updated'}


@app.delete('/api/map/{filename}/{sheet}')
async def delete_mapping(filename: str, sheet: str):
    delete_mapping_pair(MAPPING_PATH, filename, sheet)
    return {'message': 'mapping deleted'}


@app.post('/api/transfer')
async def api_transfer(payload: Dict[str, Any]):
    config = payload.get('config', [])
    if not isinstance(config, list) or not config:
        raise HTTPException(status_code=400, detail='config array is required')

    headers_by_sheet = _read_json(TARGET_HEADERS_PATH, {})
    keys_by_sheet = _read_json(KEYS_PATH, {})

    persisted = _read_json(MAPPING_PATH, {})
    mappings = {}
    for item in config:
        fname = item.get('filename')
        for t in item.get('target', []):
            sheet = t.get('sheet')
            key = f'{fname}::{sheet}'
            mappings[key] = persisted.get(key, [])

    report = transfer_data(
        config=config,
        sources_dir=SOURCES_DIR,
        target_path=TARGET_PATH,
        mappings=mappings,
        keys_by_sheet=keys_by_sheet,
        start_row=9,
    )

    return report
