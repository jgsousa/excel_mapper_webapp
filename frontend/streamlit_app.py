
import os
import requests
import streamlit as st

st.set_page_config(page_title='Migration Cockpit Mapper', layout='wide')

API_BASE = os.getenv('API_BASE_URL', 'http://localhost:8000')

if 'config' not in st.session_state:
    st.session_state.config = []

st.title('Migration Cockpit Mapper — Guided Steps')

with st.expander('Server Base URL', expanded=False):
    api_base_input = st.text_input('API base URL', value=API_BASE)
    if api_base_input:
        API_BASE = api_base_input

# ---------- Step 1: Upload Target ----------
st.header('Step 1: Upload Target.xlsx')
col1, col2 = st.columns([2,1])
with col1:
    target_file = st.file_uploader('Upload Target.xlsx', type=['xlsx'], key='tfile')
with col2:
    if st.button('Upload Target'):
        if not target_file:
            st.warning('Please choose a Target.xlsx file')
        else:
            files = {'file': (target_file.name, target_file.getvalue(), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
            r = requests.post(f'{API_BASE}/api/target/upload', files=files)
            if r.ok:
                st.success('Target uploaded and analyzed.')
                st.session_state.target_info = r.json()
            else:
                st.error(r.text)

# Show target info
if 'target_info' not in st.session_state:
    r = requests.get(f'{API_BASE}/api/target/info')
    if r.ok:
        st.session_state.target_info = r.json()

if 'target_info' in st.session_state:
    ti = st.session_state.target_info
    st.subheader('Target Sheets & Keys')
    st.json(ti)

# ---------- Step 2: Upload Sources & select sheets ----------
st.header('Step 2: Upload source files and select target sheets')
src_files = st.file_uploader('Upload one or more source files (.xlsx)', type=['xlsx'], accept_multiple_files=True, key='sfiles')
if st.button('Upload Sources'):
    if not src_files:
        st.warning('Please choose at least one source file')
    else:
        for f in src_files:
            files = {'file': (f.name, f.getvalue(), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
            r = requests.post(f'{API_BASE}/api/source/upload', files=files)
            if r.ok:
                st.success(f"Uploaded {f.name}")
            else:
                st.error(f"{f.name}: {r.text}")
        # refresh list
        r = requests.get(f'{API_BASE}/api/sources')
        if r.ok:
            st.session_state.sources = r.json()['sources']

# Load sources list if not present
if 'sources' not in st.session_state:
    r = requests.get(f'{API_BASE}/api/sources')
    if r.ok:
        st.session_state.sources = r.json()['sources']

if 'sources' in st.session_state and st.session_state.sources:
    sheets = st.session_state.get('target_info', {}).get('sheets', [])
    st.write('Select target sheets for each source:')
    new_config = []
    for s in st.session_state.sources:
        with st.container(border=True):
            st.write(f"**{s['filename']}**")
            sel = st.multiselect('Target sheets', options=sheets, key=f"sel_{s['filename']}")
            if sel:
                new_config.append({'filename': s['filename'], 'target': [{'sheet': x} for x in sel]})
    st.session_state.config = new_config

    if new_config:
        st.success('Current configuration:')
        st.json(new_config)

# Helpers to fetch headers for manual mapping

def _get_source_headers_for(fname: str):
    for s in st.session_state.get('sources', []):
        if s['filename'] == fname:
            return s.get('headers', [])
    return []

def _get_target_headers_for(sheet: str):
    ti = st.session_state.get('target_info', {})
    return ti.get('headers', {}).get(sheet, {})  # dict {header -> column}

# ---------- Step 3: Run mappings, edit & ADD new pairs ----------
st.header('Step 3: Run mappings for selected pairs, fix & add any missing')
if st.session_state.config:
    for item in st.session_state.config:
        fname = item['filename']
        targets = item['target']
        for t in targets:
            sheet = t['sheet']
            with st.container(border=True):
                st.subheader(f"Mapping for {fname} → {sheet}")
                c1, c2, c3, c4 = st.columns([1,1,1,2])
                with c1:
                    if st.button('Run mapping', key=f"runmap_{fname}_{sheet}"):
                        r = requests.post(f'{API_BASE}/api/map/run', json={'filename': fname, 'sheet': sheet})
                        if r.ok:
                            st.session_state[f'map_{fname}_{sheet}'] = r.json()['mapping']
                        else:
                            st.error(r.text)
                with c2:
                    if st.button('Load mapping', key=f"loadmap_{fname}_{sheet}"):
                        r = requests.get(f'{API_BASE}/api/map/{fname}/{sheet}')
                        if r.ok:
                            st.session_state[f'map_{fname}_{sheet}'] = r.json()['mapping']
                        else:
                            st.error(r.text)
                with c3:
                    if st.button('Delete mapping', key=f"delmap_{fname}_{sheet}"):
                        r = requests.delete(f'{API_BASE}/api/map/{fname}/{sheet}')
                        if r.ok:
                            st.session_state.pop(f'map_{fname}_{sheet}', None)
                            st.success('Mapping deleted')
                        else:
                            st.error(r.text)
                with c4:
                    st.info('Tip: You can add missing pairs below and then "Save filtered mapping".')

                mapping = st.session_state.get(f'map_{fname}_{sheet}', [])
                if mapping:
                    st.write('Review & uncheck incorrect rows:')
                    to_keep = []
                    for i, row in enumerate(mapping):
                        checked = st.checkbox(
                            label=f"{row['sourceDescription']} ({row['source']}) ⇒ {row['targetDescription']} ({row['target']})",
                            value=True,
                            key=f"chk_{fname}_{sheet}_{i}"
                        )
                        if checked:
                            to_keep.append(row)

                    st.markdown('---')
                    st.write('**Add a missing mapping**')
                    src_hdrs = _get_source_headers_for(fname)
                    tgt_hdrs_map = _get_target_headers_for(sheet)
                    src_names = [h['header'] for h in src_hdrs]
                    tgt_names = sorted(list(tgt_hdrs_map.keys()))

                    add_col1, add_col2, add_col3 = st.columns([2,2,1])
                    with add_col1:
                        sel_src = st.selectbox('Source header', options=['— choose —'] + src_names, key=f"addsrc_{fname}_{sheet}")
                    with add_col2:
                        sel_tgt = st.selectbox('Target header', options=['— choose —'] + tgt_names, key=f"addtgt_{fname}_{sheet}")
                    with add_col3:
                        if st.button('Add', key=f"addbtn_{fname}_{sheet}"):
                            if sel_src not in src_names or sel_tgt not in tgt_names:
                                st.warning('Select both a source and a target header')
                            else:
                                # Build the mapping row
                                src_col = next((h['column'] for h in src_hdrs if h['header'] == sel_src), None)
                                tgt_col = tgt_hdrs_map.get(sel_tgt)
                                if src_col and tgt_col:
                                    new_row = {
                                        'source': src_col,
                                        'sourceDescription': sel_src,
                                        'target': tgt_col,
                                        'targetDescription': sel_tgt,
                                    }
                                    # Prevent duplicates by target column
                                    existing_t_cols = {m['target'] for m in to_keep}
                                    if new_row['target'] in existing_t_cols:
                                        st.error(f"Target column {new_row['target']} already mapped; remove the old pair first.")
                                    else:
                                        to_keep.append(new_row)
                                        st.session_state[f'map_{fname}_{sheet}'] = to_keep
                                        st.success('Added mapping row (not yet saved to server).')
                                else:
                                    st.error('Could not resolve column letters; check headers.')

                    if st.button('Save filtered mapping', key=f"savemap_{fname}_{sheet}"):
                        r = requests.post(f'{API_BASE}/api/map/{fname}/{sheet}', json={'mapping': to_keep})
                        if r.ok:
                            st.success('Mapping updated')
                        else:
                            st.error(r.text)
                else:
                    st.info('No mapping loaded yet. Run or Load mapping above, then add pairs as needed.')

# ---------- Step 4: Run transfer ----------
st.header('Step 4: Transfer from source to target')
if st.button('Run transfer now'):
    if not st.session_state.config:
        st.warning('Please configure at least one source → sheet pair.')
    else:
        r = requests.post(f'{API_BASE}/api/transfer', json={'config': st.session_state.config})
        if r.ok:
            st.success('Transfer completed')
            st.json(r.json())
            st.download_button('Download updated Target.xlsx', data=requests.get(f'{API_BASE}/api/target/download').content, file_name='Target.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            st.error(r.text)
