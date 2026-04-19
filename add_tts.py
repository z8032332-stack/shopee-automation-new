"""
add_tts.py — 為已完成但缺旁白的影片補上 TTS + BGM
用法：python add_tts.py           （預設處理 ROW_START ~ ROW_END）
      python add_tts.py 8 13     （指定行號範圍）

原理：
  1. 從 Excel 讀文案 → edge-tts 生成旁白 MP3
  2. BGM 混音
  3. FFmpeg -c:v copy（影片不重新渲染，只換音軌）
  速度：每支約 1-2 分鐘
"""
import sys, io, os, re, glob, time, random, logging, asyncio, subprocess, tempfile
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

from dotenv import load_dotenv
load_dotenv(dotenv_path=os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env'))

import openpyxl
import edge_tts
import imageio_ffmpeg

logging.basicConfig(level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s', datefmt='%H:%M:%S')
log = logging.getLogger(__name__)

# ── 路徑設定（從 .env 讀）────────────────────────────────────────────────────
EXCEL_PATH = os.getenv('EXCEL_PATH',  r'C:\Users\user\Desktop\蝦皮素材\蝦皮選品_2026年 (1).xlsx')
FINAL_DIR  = os.getenv('FINAL_DIR',   r'C:\Users\user\Desktop\蝦皮素材\影片完成')
BGM_DIR    = os.getenv('BGM_DIR',     r'C:\Users\user\Desktop\蝦皮素材\BGM音樂')
FFMPEG_EXE = os.getenv('FFMPEG_PATH', '') or imageio_ffmpeg.get_ffmpeg_exe()
TTS_VOICE  = os.getenv('TTS_VOICE',   'zh-TW-HsiaoChenNeural')

COL_NAME   = int(os.getenv('COL_NAME',   '2'))
COL_COPY   = int(os.getenv('COL_COPY',   '8'))
COL_STATUS = int(os.getenv('COL_STATUS', '10'))

# 預設補的行號範圍（Excel 資料行，不含標題）
ROW_START = 8
ROW_END   = 13

# ══════════════════════════════════════════════════════════════════════════════
# 斷句
# ══════════════════════════════════════════════════════════════════════════════
def split_sentences(text):
    text = str(text).replace('\\n', '\n').replace('\n', '。')
    parts = re.split(r'(?<=[。！？])', text)
    return [s.strip() for s in parts
            if s.strip() and re.search(r'[\u4e00-\u9fff]', s)]

# ══════════════════════════════════════════════════════════════════════════════
# TTS 生成（單一 event loop，避免 Windows WinError 2）
# ══════════════════════════════════════════════════════════════════════════════
async def _generate_all_tts(sentences, tmpdir):
    """在同一個 event loop 內依序生成所有句子的 TTS"""
    segments = []
    for i, sent in enumerate(sentences):
        mp3 = os.path.join(tmpdir, f'tts_{i:03d}.mp3')
        success = False
        for attempt in range(3):
            try:
                await asyncio.sleep(0.8 + attempt * 0.5)
                communicate = edge_tts.Communicate(sent, TTS_VOICE)
                await communicate.save(mp3)
                success = True
                break
            except Exception as e:
                log.warning('  TTS [%d] 第%d次失敗: %s', i, attempt + 1, e)
        if success and os.path.exists(mp3):
            segments.append({'mp3': mp3, 'text': sent})
        else:
            segments.append({'mp3': None, 'text': sent})
    return segments

def speed_to_target(src_mp3, dst_mp3, target_dur):
    """用 FFmpeg atempo 把音檔壓縮/拉伸到 target_dur 秒"""
    from moviepy import AudioFileClip
    try:
        clip = AudioFileClip(src_mp3)
        natural = clip.duration
        clip.close()
    except Exception:
        natural = target_dur

    speed = natural / target_dur  # >1 加速, <1 減速
    speed = max(0.5, min(speed, 8.0))  # 限制範圍

    # atempo 每次只能 0.5~2.0，超過要串聯
    filters = []
    s = speed
    while s > 2.0:
        filters.append('atempo=2.0')
        s /= 2.0
    while s < 0.5:
        filters.append('atempo=0.5')
        s /= 0.5
    filters.append(f'atempo={s:.6f}')
    af = ','.join(filters)

    cmd = [FFMPEG_EXE, '-y', '-i', src_mp3,
           '-filter:a', af, dst_mp3]
    subprocess.run(cmd, capture_output=True, timeout=30)

def build_tts_mp3(sentences, tmpdir, target_per_sent=2.5):
    """生成每句 TTS，壓縮到 target_per_sent 秒，對齊字幕"""
    raw = asyncio.run(_generate_all_tts(sentences, tmpdir))

    segments = []
    for i, seg in enumerate(raw):
        if seg['mp3'] and os.path.exists(seg['mp3']):
            # 壓縮到 target_per_sent 秒
            timed_mp3 = os.path.join(tmpdir, f'tts_timed_{i:03d}.mp3')
            speed_to_target(seg['mp3'], timed_mp3, target_per_sent)
            mp3_final = timed_mp3 if os.path.exists(timed_mp3) else seg['mp3']
            segments.append({'mp3': mp3_final, 'duration': target_per_sent, 'text': seg['text']})
            log.info('  TTS [%d] → %.1fs（對齊字幕）%s', i, target_per_sent, seg['text'][:25])
        else:
            segments.append({'mp3': None, 'duration': target_per_sent, 'text': seg['text']})
    return segments

# ══════════════════════════════════════════════════════════════════════════════
# BGM
# ══════════════════════════════════════════════════════════════════════════════
def pick_bgm():
    files = []
    for ext in ('*.mp3', '*.wav', '*.m4a', '*.aac'):
        files += glob.glob(os.path.join(BGM_DIR, ext))
    return random.choice(files) if files else None

# ══════════════════════════════════════════════════════════════════════════════
# 核心：用 FFmpeg 把 TTS+BGM 混入影片（-c:v copy，不重新渲染）
# ══════════════════════════════════════════════════════════════════════════════
def replace_audio(video_path, segments, tmpdir):
    valid_segs = [s for s in segments if s['mp3'] and os.path.exists(s['mp3'])]
    if not valid_segs:
        log.warning('  沒有有效 TTS，跳過')
        return False

    total_tts = sum(s['duration'] for s in segments)

    # Step A：把所有 TTS mp3 串接成一個 tts_full.mp3
    concat_list = os.path.join(tmpdir, 'concat.txt')
    with open(concat_list, 'w', encoding='utf-8') as f:
        for s in valid_segs:
            f.write(f"file '{s['mp3'].replace(chr(92), '/')}'\n")

    tts_full = os.path.join(tmpdir, 'tts_full.mp3')
    cmd = [FFMPEG_EXE, '-y', '-f', 'concat', '-safe', '0',
           '-i', concat_list, '-c', 'copy', tts_full]
    subprocess.run(cmd, capture_output=True, timeout=60)

    # Step C：混音 TTS(100%) + BGM(25%)
    bgm_file = pick_bgm()
    mixed_audio = os.path.join(tmpdir, 'mixed.aac')

    if bgm_file:
        # BGM 先loop到夠長，再 amix
        cmd = [
            FFMPEG_EXE, '-y',
            '-i', tts_full,
            '-stream_loop', '-1', '-i', bgm_file,
            '-filter_complex',
            f'[0:a]volume=1.0[tts];[1:a]volume=0.25[bgm];[tts][bgm]amix=inputs=2:duration=first[out]',
            '-map', '[out]',
            '-c:a', 'aac', '-b:a', '128k',
            '-t', str(total_tts),
            mixed_audio
        ]
        log.info('  TTS + BGM 混音')
    else:
        cmd = [FFMPEG_EXE, '-y', '-i', tts_full,
               '-c:a', 'aac', '-b:a', '128k', mixed_audio]
        log.info('  只有 TTS（無 BGM）')

    subprocess.run(cmd, capture_output=True, timeout=120)

    if not os.path.exists(mixed_audio):
        log.error('  混音失敗')
        return False

    # Step D：-c:v copy 換音軌（不重新渲染影片）
    tmp_out = video_path + '.newaudio.mp4'
    cmd = [
        FFMPEG_EXE, '-y',
        '-i', video_path,
        '-i', mixed_audio,
        '-c:v', 'copy',          # 影片直接複製，不重新渲染
        '-c:a', 'aac',
        '-map', '0:v:0',
        '-map', '1:a:0',
        '-shortest',
        '-movflags', '+faststart',
        tmp_out
    ]
    result = subprocess.run(cmd, capture_output=True, timeout=120)
    if result.returncode != 0 or not os.path.exists(tmp_out):
        log.error('  FFmpeg 失敗: %s', result.stderr.decode(errors='replace')[-200:])
        return False

    os.replace(tmp_out, video_path)
    log.info('  ✓ 音軌替換完成')
    return True

# ══════════════════════════════════════════════════════════════════════════════
# 主程式
# ══════════════════════════════════════════════════════════════════════════════
def main():
    # 從命令列取得行號範圍（可選）
    args = sys.argv[1:]
    row_start = int(args[0]) if len(args) >= 1 else ROW_START
    row_end   = int(args[1]) if len(args) >= 2 else ROW_END

    print('=' * 55)
    print(f'補旁白工具 — 行 {row_start} ~ {row_end}')
    print(f'影片目錄: {FINAL_DIR}')
    print(f'TTS 聲音: {TTS_VOICE}')
    print('=' * 55)

    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active
    done, fail, skip = 0, 0, 0

    for row in ws.iter_rows(min_row=2):
        row_idx = row[0].row - 1  # 資料行號（從1開始）
        if row_idx < row_start or row_idx > row_end:
            continue

        name      = row[COL_NAME   - 1].value
        copy_text = row[COL_COPY   - 1].value

        if not name:
            continue
        if not copy_text or str(copy_text).strip() in ('', 'None', '文案'):
            log.warning('[%d] 無文案，跳過', row_idx)
            skip += 1
            continue

        # 找對應影片
        pattern = os.path.join(FINAL_DIR, f'{row_idx:03d}_*.mp4')
        matches = glob.glob(pattern)
        if not matches:
            log.warning('[%d] 找不到影片（%s）', row_idx, pattern)
            skip += 1
            continue

        video_path = matches[0]
        print(f'\n[{row_idx}] {str(name)[:45]}')
        print(f'  影片: {os.path.basename(video_path)}')

        sentences = split_sentences(str(copy_text))
        log.info('  斷句 %d 句', len(sentences))

        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmpdir:
            segments = build_tts_mp3(sentences, tmpdir)
            total_tts = sum(s['duration'] for s in segments)
            log.info('  TTS 總時長 %.1fs', total_tts)

            if replace_audio(video_path, segments, tmpdir):
                done += 1
                log.info('完成 → %s', os.path.basename(video_path))
            else:
                fail += 1

    print(f'\n{"=" * 55}')
    print(f'完成 {done} 個 / 失敗 {fail} 個 / 跳過 {skip} 個')

if __name__ == '__main__':
    main()
