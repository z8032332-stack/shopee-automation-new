"""
蝦皮自動化工具 - 控制台
啟動: py -3.12 -m streamlit run dashboard.py
"""
import streamlit as st
import subprocess, threading, queue, os, sys, time, re
from dotenv import load_dotenv

TOOL_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(TOOL_DIR, '.env')
load_dotenv(dotenv_path=ENV_PATH)

st.set_page_config(page_title='蝦皮控制台', page_icon='🛒', layout='wide')
st.title('🛒 蝦皮自動化控制台')

# ── session state ─────────────────────────────────────────────────────────────
if 'running'     not in st.session_state: st.session_state.running     = False
if 'proc'        not in st.session_state: st.session_state.proc        = None
if 'logs'        not in st.session_state: st.session_state.logs        = []
if 'current'     not in st.session_state: st.session_state.current     = ''
if 'log_q'       not in st.session_state: st.session_state.log_q       = None
if 'active_step' not in st.session_state: st.session_state.active_step = ''
if 'prog_cur'    not in st.session_state: st.session_state.prog_cur    = 0
if 'prog_total'  not in st.session_state: st.session_state.prog_total  = 0
if 'start_time'  not in st.session_state: st.session_state.start_time  = None
if 'start_cur'   not in st.session_state: st.session_state.start_cur   = 0

# ── helper ────────────────────────────────────────────────────────────────────
def update_env(updates: dict):
    lines = open(ENV_PATH, encoding='utf-8').readlines()
    new, found = [], set()
    for l in lines:
        key = l.split('=')[0].strip()
        if key in updates:
            new.append(f'{key}={updates[key]}\n'); found.add(key)
        else:
            new.append(l)
    for k, v in updates.items():
        if k not in found:
            new.append(f'{k}={v}\n')
    open(ENV_PATH, 'w', encoding='utf-8').writelines(new)

def _reader(proc, q):
    for line in iter(proc.stdout.readline, ''):
        q.put(line.rstrip())
    proc.stdout.close(); proc.wait()
    q.put(f'[系統] 程序結束 (exit={proc.returncode})')

def drain():
    q = st.session_state.log_q
    if not q: return
    while True:
        try:
            line = q.get_nowait()
            st.session_state.logs.append(line)
            m = re.search(r'\[(\s*\d+)/(\s*\d+)\]\s*(.+)', line)
            if m:
                cur   = int(m.group(1).strip())
                total = int(m.group(2).strip())
                name  = m.group(3).strip()[:50]
                st.session_state.current    = f"[{cur}/{total}] {name}"
                st.session_state.prog_cur   = cur
                st.session_state.prog_total = total
            if '[系統] 程序結束' in line:
                st.session_state.running = False
        except: break

def launch(cmd: list, step_name: str):
    st.session_state.logs        = []
    st.session_state.current     = ''
    st.session_state.running     = True
    st.session_state.active_step = step_name
    st.session_state.prog_cur    = 0
    st.session_state.prog_total  = 0
    st.session_state.start_time  = time.time()
    st.session_state.start_cur   = 0
    q = queue.Queue()
    st.session_state.log_q = q
    proc = subprocess.Popen(
        cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, encoding='utf-8', errors='replace', cwd=TOOL_DIR,
    )
    st.session_state.proc = proc
    threading.Thread(target=_reader, args=(proc, q), daemon=True).start()

def stop_proc():
    p = st.session_state.proc
    if p and p.poll() is None:
        p.terminate(); time.sleep(1)
        if p.poll() is None: p.kill()
    st.session_state.running = False
    st.session_state.logs.append('[系統] 使用者強制停止')

def status_bar(step_name):
    """顯示執行狀態列 + 進度條 + 停止鍵"""
    is_this = st.session_state.active_step == step_name
    c1, c2 = st.columns([6, 2])
    with c1:
        if st.session_state.running and is_this:
            cur        = st.session_state.prog_cur
            total      = st.session_state.prog_total
            t0         = st.session_state.start_time
            start_cur  = st.session_state.start_cur
            elapsed    = time.time() - t0 if t0 else 0
            done_items = max(cur - start_cur, 0)

            # 預估剩餘時間
            eta_str = ''
            if done_items > 0 and total > 0:
                secs_per = elapsed / done_items
                remain   = max(total - cur, 0)
                eta_secs = int(secs_per * remain)
                if eta_secs >= 3600:
                    eta_str = f'⏱ 預估剩餘 {eta_secs//3600}h{(eta_secs%3600)//60}m'
                elif eta_secs >= 60:
                    eta_str = f'⏱ 預估剩餘 {eta_secs//60}m{eta_secs%60}s'
                else:
                    eta_str = f'⏱ 預估剩餘 {eta_secs}s'

            st.info(f'🔄 {st.session_state.current}  {eta_str}')
            if total > 0:
                pct = min(cur / total, 1.0)
                elapsed_str = f'{int(elapsed//60)}m{int(elapsed%60)}s'
                st.progress(pct, text=f'{cur} / {total} 筆　已用 {elapsed_str}　{eta_str}')
        elif not st.session_state.running and is_this and st.session_state.logs:
            last = st.session_state.logs[-1]
            if '強制停止' in last:
                st.warning('⛔ 已停止')
            else:
                total = st.session_state.prog_total
                st.success(f'✅ 完成（共 {total} 筆）' if total else '✅ 完成')
    with c2:
        can_stop = st.session_state.running and is_this
        if st.button('🛑 強制停止', key=f'stop_{step_name}',
                     disabled=not can_stop, use_container_width=True):
            stop_proc(); st.rerun()

def log_panel(step_name):
    """顯示 log"""
    drain()
    is_this = st.session_state.active_step == step_name
    if is_this and st.session_state.logs:
        st.code('\n'.join(st.session_state.logs[-300:]), language=None)
        if st.button('🗑️ 清除記錄', key=f'clear_{step_name}'):
            st.session_state.logs = []; st.session_state.current = ''; st.rerun()
    elif not is_this and st.session_state.active_step:
        st.caption(f'目前執行中：{st.session_state.active_step}')
    else:
        st.caption('尚未開始...')

# ── Excel 路徑（共用）────────────────────────────────────────────────────────
with st.sidebar:
    st.header('⚙️ 共用設定')
    excel_path = st.text_input(
        'Excel 路徑',
        value=os.getenv('EXCEL_PATH',
              r'C:\Users\user\Desktop\蝦皮素材\蝦皮選品_2026年5月new.xlsx')
    )
    if os.path.exists(excel_path):
        st.success('✅ Excel 找到')
    else:
        st.error('❌ Excel 不存在')
    st.divider()
    st.caption(f'工具目錄:\n{TOOL_DIR}')

# ── 分頁 ──────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    '🔍 Step1 選品',
    '🎬 Step2 抓影片',
    '✍️ Step3 文案+標題',
    '🎞️ Step4 後製',
    '📤 Step5 上傳',
])

# ═══════════════════════════ Step 1 選品 ══════════════════════════════════════
with tab1:
    st.subheader('Step 1 — 選品')
    st.info('⚠️ 執行前請先用 CDP 模式開 Chrome 並登入 affiliate.shopee.tw')
    st.code(
        r'"C:\Program Files\Google\Chrome\Application\chrome.exe" '
        r'--remote-debugging-port=9222 --user-data-dir="C:\Temp\ChromeCDP2"',
        language=None
    )

    status_bar('選品')

    disabled = st.session_state.running
    if st.button('▶️ 開始選品', key='start_step1',
                 disabled=disabled, type='primary', use_container_width=True):
        update_env({'EXCEL_PATH': excel_path})
        launch([sys.executable, os.path.join(TOOL_DIR, 'shopee_keyword_scraper_home.py')], '選品')
        st.rerun()

    log_panel('選品')

# ═══════════════════════════ Step 2 抓影片 ════════════════════════════════════
with tab2:
    st.subheader('Step 2 — 抓影片素材')
    st.info('⚠️ Chrome 彈出後請登入蝦皮，完成後按 Enter')

    c1, c2 = st.columns(2)
    with c1:
        s2_start = st.number_input('從第幾筆（編號）', min_value=1, value=13, step=1, key='s2s')
    with c2:
        s2_end   = st.number_input('到第幾筆（編號）', min_value=1, value=9999, step=1, key='s2e')

    status_bar('抓影片')

    disabled = st.session_state.running
    if st.button('▶️ 開始抓影片素材', key='start_step2',
                 disabled=disabled, type='primary', use_container_width=True):
        if not os.path.exists(excel_path):
            st.error('找不到 Excel！'); st.stop()
        update_env({'EXCEL_PATH': excel_path,
                    'START_ROW': str(int(s2_start)),
                    'END_ROW':   str(int(s2_end))})
        launch([sys.executable, os.path.join(TOOL_DIR, 'shopee_video_maker_home3.py')], '抓影片')
        st.rerun()

    log_panel('抓影片')

# ═══════════════════════════ Step 3 文案+標題 ══════════════════════════════════
with tab3:
    st.subheader('Step 3 — 文案 + 標題')

    c1, c2 = st.columns(2)
    with c1:
        s3_start = st.number_input('從第幾筆（編號）', min_value=1, value=1, step=1, key='s3s')
    with c2:
        s3_end   = st.number_input('到第幾筆（編號）', min_value=1, value=9999, step=1, key='s3e')

    status_bar('文案')

    disabled = st.session_state.running
    if st.button('▶️ 開始生成文案+標題', key='start_step3',
                 disabled=disabled, type='primary', use_container_width=True):
        if not os.path.exists(excel_path):
            st.error('找不到 Excel！'); st.stop()
        update_env({'EXCEL_PATH': excel_path,
                    'START_ROW': str(int(s3_start)),
                    'END_ROW':   str(int(s3_end))})
        launch([sys.executable, os.path.join(TOOL_DIR, 'gen_copy.py')], '文案')
        st.rerun()

    log_panel('文案')

# ═══════════════════════════ Step 4 後製 ══════════════════════════════════════
with tab4:
    st.subheader('Step 4 — 後製影片')

    c1, c2 = st.columns(2)
    with c1:
        s4_start = st.number_input('從第幾筆（編號）', min_value=1, value=1, step=1, key='s4s')
    with c2:
        s4_end   = st.number_input('到第幾筆（編號）', min_value=1, value=9999, step=1, key='s4e')

    status_bar('後製')

    disabled = st.session_state.running
    if st.button('▶️ 開始後製', key='start_step4',
                 disabled=disabled, type='primary', use_container_width=True):
        if not os.path.exists(excel_path):
            st.error('找不到 Excel！'); st.stop()
        update_env({'EXCEL_PATH': excel_path,
                    'START_ROW': str(int(s4_start)),
                    'END_ROW':   str(int(s4_end))})
        launch([sys.executable, os.path.join(TOOL_DIR, 'shopee_video_producer.py')], '後製')
        st.rerun()

    log_panel('後製')

# ═══════════════════════════ Step 5 上傳 ══════════════════════════════════════
with tab5:
    st.subheader('Step 5 — 上傳到蝦皮')

    c1, c2, c3 = st.columns(3)
    with c1:
        s5_start = st.number_input('從第幾筆（編號）', min_value=1, value=662, step=1, key='s5s')
    with c2:
        s5_count = st.number_input('上傳幾部', min_value=1, value=50, step=1, key='s5c')
    with c3:
        s5_gap   = st.number_input('每部間隔（秒）', min_value=10, value=60, step=5, key='s5g')

    phone_ip = st.text_input('手機 IP:PORT',
                             value=os.getenv('PHONE_IP', '192.168.0.12:5555'),
                             help='手機重開機後 IP 可能會變，每次上傳前確認')

    status_bar('上傳')

    disabled = st.session_state.running
    if st.button('▶️ 開始上傳', key='start_step5',
                 disabled=disabled, type='primary', use_container_width=True):
        if not os.path.exists(excel_path):
            st.error('找不到 Excel！'); st.stop()
        update_env({'EXCEL_PATH': excel_path, 'PHONE_IP': phone_ip})
        launch([
            sys.executable,
            os.path.join(TOOL_DIR, 'shopee_upload_a52s.py'),
            '--phone', 'a52s',
            '--start', str(int(s5_start)),
            '--count', str(int(s5_count)),
            '--gap',   str(int(s5_gap)),
        ], '上傳')
        st.rerun()

    log_panel('上傳')

# ── 自動刷新（執行中每 2 秒刷新）────────────────────────────────────────────
if st.session_state.running:
    time.sleep(2)
    st.rerun()
