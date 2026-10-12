/**
 * CasePhotos — PM 아코디언 '시공 완료 사진' 칸 (2026-10-12, 관리자 전용).
 *
 * 사업자등록증 칸처럼 프로젝트 아코디언에서 바로 완료 사진·작업지시서를 올리고
 * [시공사례 글 요청]을 누르면 평일 아침 Claude 예약 작업이 초안을 쓴다.
 * 초안 확인·수정·복사는 /case-posts 페이지(#job=...)에서. 사진은 화면에 띄우지 않고
 * 사진 수를 드라이브 '시공사례' 폴더 링크로 보여 준다(문서 폴더에서 확인).
 *
 * 백엔드: /api/case-posts/... (blueprints/case_posts.py). 접수 1건 = job.
 *  - 진행 중 job(사진 보관/접수)이 있으면 거기에 파일을 더하고, 없거나 이미 초안이 나왔으면 새 job.
 * 사진은 긴 변 2000px JPEG 로 줄여 한 장씩 보낸다(프로덕션 요청 8MB 제한).
 */
import logger from '../utils/logger.js';

const API = '/api/case-posts';
const MAX_EDGE = 2000;
const OPEN_STATUSES = ['uploading', 'queued'];
const STATUS_STYLE = {
  uploading: 'background:#fff3bf;color:#946200;',
  queued: 'background:#e7f5ff;color:#1864ab;',
  drafted: 'background:#d3f9d8;color:#2b8a3e;',
  posted: 'background:#e9ecef;color:#495057;',
};

function esc(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function attr(v) {
  return String(v == null ? '' : v).replace(/["\\]/g, '\\$&');
}

async function api(path, opts) {
  const res = await fetch(API + path, Object.assign({ credentials: 'same-origin' }, opts || {}));
  let data = {};
  try { data = await res.json(); } catch (_) { /* noop */ }
  if (!res.ok || data.success === false) throw new Error(data.error || `요청 실패 (${res.status})`);
  return data;
}

async function shrink(file) {
  if (!file.type.startsWith('image/') || file.type === 'image/gif') return { blob: file, name: file.name };
  try {
    const bmp = await createImageBitmap(file, { imageOrientation: 'from-image' });
    const scale = Math.min(1, MAX_EDGE / Math.max(bmp.width, bmp.height));
    if (scale === 1 && file.size < 1.5 * 1024 * 1024) { bmp.close(); return { blob: file, name: file.name }; }
    const c = document.createElement('canvas');
    c.width = Math.round(bmp.width * scale); c.height = Math.round(bmp.height * scale);
    c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height); bmp.close();
    const blob = await new Promise(r => c.toBlob(r, 'image/jpeg', 0.85));
    if (!blob) return { blob: file, name: file.name };
    return { blob, name: file.name.replace(/\.[^.]+$/, '') + '.jpg' };
  } catch (_) { return { blob: file, name: file.name }; }
}

export default class CasePhotos {
  constructor() {
    this._jobs = {};   // code → job | null
    this._busy = {};   // code → true (업로드 중)
    window.casePhotos = this;
  }

  static isAdmin() {
    return window.userRole === 'admin' || window.userRole === 'super_admin';
  }

  _card(code) {
    return document.querySelector(`[data-cp-card="${attr(code)}"]`);
  }

  _toast(message, type = 'info') {
    import('./Toast.js').then(({ default: Toast }) => new Toast().show(message, type))
      .catch(() => logger.debug(`[알림] ${message}`));
  }

  /** 아코디언 렌더 직후(아직 DOM 삽입 전) 호출 — fetch 완료 시점엔 삽입돼 있다. */
  onAccordionRendered(code) {
    const c = String(code || '').trim();
    if (!c || !CasePhotos.isAdmin()) return;
    this.refresh(c);
  }

  async refresh(code) {
    try {
      const data = await api(`/by-project/${encodeURIComponent(code)}`);
      this._jobs[code] = data.job;
    } catch (err) {
      logger.warn('[CASE_PHOTOS] 조회 실패:', err);
      this._jobs[code] = undefined;
    }
    this._render(code);
  }

  _render(code, progress) {
    const card = this._card(code);
    if (!card) return;
    const body = card.querySelector('[data-cp-body]');
    if (!body) return;
    if (!card.dataset.cpBound) this._bindDrop(card, code);
    const job = this._jobs[code];
    if (job === undefined) { body.innerHTML = '<span style="color:#adb5bd;">불러오지 못했습니다</span>'; return; }
    if (progress) { body.innerHTML = `<span style="color:#1c7ed6;font-size:.85rem;">${esc(progress)}</span>`; return; }

    const files = (job && job.files) || [];
    const photos = files.filter(f => f.kind === 'photo').length;
    const orders = files.filter(f => f.kind === 'order').length;
    const busy = !!this._busy[code];
    const pill = (text, style, title = '') =>
      `<span title="${esc(title)}" style="display:inline-flex;align-items:center;padding:3px 10px;border-radius:12px;${style}">${text}</span>`;

    // 사진 수 — 드라이브 '시공사례' 폴더가 있으면 그 폴더로 링크(사진은 문서 폴더에서 확인)
    let countPill;
    if (!job) {
      countPill = pill('없음', 'background:#eef0f2;color:#6b7280;', '사진을 올리거나 카톡 PC에서 이 칸으로 끌어다 놓으세요');
    } else {
      const text = `사진 ${photos}${job.has_order_text || job.order_text ? ' · 지시 글' : ''}${orders ? ` · 도면 ${orders}` : ''}`;
      countPill = job.drive_folder_id
        ? `<a href="https://drive.google.com/drive/folders/${encodeURIComponent(job.drive_folder_id)}" target="_blank" rel="noopener"
              title="드라이브 문서 폴더 / 시공사례 열기" style="display:inline-flex;align-items:center;padding:3px 10px;border-radius:12px;background:#e7f5ff;color:#1864ab;text-decoration:none;">${text}</a>`
        : pill(text, 'background:#eef0f2;color:#495057;', (job.drive_errors || []).join(' / ') || '드라이브 백업 없음');
    }
    const statusPill = !job || job.status === 'uploading' ? ''
      : pill(esc(job.status_label), STATUS_STYLE[job.status] || '',
        job.status === 'queued' ? '다음 평일 아침 7시 반쯤 초안이 만들어집니다' : '');
    const warn = job && (job.drive_errors || []).length
      ? `<i class="fas fa-exclamation-triangle" style="color:#e8590c;" title="드라이브 백업: ${esc(job.drive_errors.join(' / '))}"></i>` : '';

    const q = `'${attr(code)}'`;
    const btn = (label, icon, onclick, title, cls = 'btn-outline-secondary', disabled = false) =>
      `<button type="button" class="btn btn-sm ${cls}" style="padding:1px 8px;" title="${esc(title)}" ${disabled ? 'disabled' : ''}
               onclick="${onclick}"><i class="${icon} me-1"></i>${label}</button>`;
    const actions = [
      btn('사진', 'fas fa-images', `window.casePhotos.pick(${q},'photo')`, '완료 사진 올리기 (여러 장)', 'btn-outline-primary', busy),
      btn('지시서', 'fas fa-file-alt', `window.casePhotos.openOrder(${q})`,
        '카톡 작업지시 글 붙여넣기', job && (job.has_order_text || job.order_text) ? 'btn-outline-success' : 'btn-outline-secondary', busy),
    ];
    if (job && job.status === 'uploading') {
      actions.push(btn('글 요청', 'fas fa-pen-nib', `window.casePhotos.requestDraft(${q})`,
        '시공사례 글로 만들기 — 다음 평일 아침에 초안 작성', 'btn-primary', busy || !photos));
    }
    if (job && job.status !== 'uploading') {
      actions.push(`<a class="btn btn-sm btn-outline-success" style="padding:1px 8px;" href="/case-posts#job=${encodeURIComponent(job.id)}"
                       target="_blank" rel="noopener" title="시공사례 페이지에서 보기"><i class="fas fa-external-link-alt me-1"></i>${job.status === 'queued' ? '접수' : '초안'}</a>`);
    }
    body.innerHTML = `${countPill}${statusPill}${warn}${actions.join('')}`;
  }

  _bindDrop(card, code) {
    card.dataset.cpBound = '1';
    card.addEventListener('dragover', e => { e.preventDefault(); card.style.outline = '2px dashed #339af0'; });
    card.addEventListener('dragleave', () => { card.style.outline = ''; });
    card.addEventListener('drop', e => {
      e.preventDefault(); card.style.outline = '';
      // 끌어다 놓기는 사진만 (작업지시는 [지시서] 창에 글로)
      const imgs = Array.from(e.dataTransfer.files || []).filter(f => !/\.pdf$/i.test(f.name) && f.type !== 'application/pdf');
      this.upload(code, imgs.map(f => ['photo', f]));
    });
  }

  pick(code, kind) {
    const input = document.createElement('input');
    input.type = 'file';
    input.multiple = true;
    input.accept = kind === 'order' ? 'image/*,application/pdf' : 'image/*';
    input.onchange = () => this.upload(code, Array.from(input.files || []).map(f => [kind, f]));
    input.click();
  }

  async _ensureJob(code) {
    const job = this._jobs[code];
    if (job && OPEN_STATUSES.includes(job.status)) return job.id;
    const fd = new FormData(); fd.append('code', code);
    const data = await api('/jobs', { method: 'POST', body: fd });
    return data.job.id;
  }

  /** 작업지시 창 — 카톡 단체방 작업지시 글 붙여넣기(텍스트만). */
  openOrder(code) {
    const job = this._jobs[code];
    const editable = !job || ['uploading', 'queued'].includes(job.status);
    document.getElementById('cpOrderOverlay')?.remove();
    const ov = document.createElement('div');
    ov.id = 'cpOrderOverlay';
    ov.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:2000;display:flex;align-items:center;justify-content:center;padding:16px;';
    ov.innerHTML = `
      <div style="background:#fff;border-radius:10px;width:100%;max-width:560px;padding:16px;box-shadow:0 10px 30px rgba(0,0,0,.2);">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
          <b><i class="fas fa-file-alt me-2" style="color:#1c7ed6;"></i>${esc(code)} 작업지시</b>
          <button type="button" class="btn btn-sm btn-light" data-cp-close>✕</button>
        </div>
        <div style="font-size:.85rem;color:#868e96;margin-bottom:6px;">
          카톡 단체방의 작업지시 글을 길게 눌러(PC는 드래그) 복사한 뒤 붙여넣으세요. 전화번호·호수는 글에 쓰이지 않습니다.
        </div>
        <textarea data-cp-order-text rows="12" style="width:100%;border:1px solid #ced4da;border-radius:6px;padding:8px;font-size:.9rem;line-height:1.5;"
                  placeholder="예)
날짜 : 10월 13일 오전 9시
삼성 5마력 냉난방기 4대 설치 공사
15배관 50M이내 ...
1. 실외기실은 같은층 발코니 입니다." ${editable ? '' : 'disabled'}></textarea>
        <div style="display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-top:10px;">
          <span style="flex:1"></span>
          <button type="button" class="btn btn-sm btn-light" data-cp-close>닫기</button>
          <button type="button" class="btn btn-sm btn-primary" data-cp-order-save ${editable ? '' : 'disabled'}>저장</button>
        </div>
        ${editable ? '' : '<div style="font-size:.8rem;color:#e8590c;margin-top:6px;">초안이 이미 나온 건이라 바꿀 수 없습니다.</div>'}
      </div>`;
    document.body.appendChild(ov);
    const ta = ov.querySelector('[data-cp-order-text]');
    ta.value = (job && job.order_text) || '';
    const close = () => ov.remove();
    ov.querySelectorAll('[data-cp-close]').forEach(b => b.addEventListener('click', close));
    ov.addEventListener('click', e => { if (e.target === ov) close(); });
    ov.querySelector('[data-cp-order-save]').addEventListener('click', async (e) => {
      e.currentTarget.disabled = true;
      try {
        const jobId = await this._ensureJob(code);
        const fd = new FormData(); fd.append('text', ta.value);
        await api(`/jobs/${encodeURIComponent(jobId)}/order-text`, { method: 'POST', body: fd });
        close();
        this._toast('작업지시를 저장했습니다', 'success');
        await this.refresh(code);
      } catch (err) {
        e.currentTarget.disabled = false;
        this._toast(err.message, 'error');
      }
    });
    if (editable) ta.focus();
  }

  async upload(code, items) {
    items = items.filter(([, f]) => f.type.startsWith('image/') || f.type === 'application/pdf'
      || /\.(heic|jpe?g|png|webp|pdf)$/i.test(f.name));
    if (!items.length || this._busy[code]) return;
    this._busy[code] = true;
    const failed = [];
    const driveErrs = new Set();
    try {
      const jobId = await this._ensureJob(code);
      for (let i = 0; i < items.length; i++) {
        const [kind, file] = items[i];
        this._render(code, `올리는 중 ${i + 1} / ${items.length} (${file.name})`);
        let ok = false; let lastErr = '';
        for (let attempt = 0; attempt < 3 && !ok; attempt++) {
          try {
            const s = await shrink(file);
            if (s.blob.size > 7.5 * 1024 * 1024) throw new Error('파일이 너무 큽니다(7.5MB 초과)');
            const fd = new FormData(); fd.append('kind', kind); fd.append('file', s.blob, s.name);
            const r = await api(`/jobs/${encodeURIComponent(jobId)}/files`, { method: 'POST', body: fd });
            if (r.drive_error) driveErrs.add(r.drive_error);
            ok = true;
          } catch (err) {
            lastErr = err.message;
            if (/너무 큽니다|올릴 수 없습니다|더할 수 없습니다/.test(lastErr)) break;
          }
        }
        if (!ok) failed.push(`${file.name}: ${lastErr}`);
      }
    } catch (err) {
      failed.push(err.message);
    } finally {
      this._busy[code] = false;
    }
    await this.refresh(code);
    if (failed.length) this._toast(`못 올린 파일 ${failed.length}개 — ${failed[0]}`, 'error');
    else this._toast(`${items.length}개 올렸습니다`, 'success');
    if (driveErrs.size) logger.warn('[CASE_PHOTOS] 드라이브 백업:', [...driveErrs]);
  }

  async requestDraft(code) {
    const job = this._jobs[code];
    if (!job) return;
    if (!(job.has_order_text || job.order_text) && !(job.files || []).some(f => f.kind === 'order')
        && !confirm('작업지시가 없습니다. 사진만으로 쓰면 빈칸(○○)이 많아집니다. 그래도 요청할까요?')) return;
    if (!confirm('이 현장을 시공사례 글로 만들까요?\n다음 평일 아침에 홈페이지·블로그 초안이 만들어집니다.')) return;
    try {
      await api(`/jobs/${encodeURIComponent(job.id)}/submit`, { method: 'POST' });
      this._toast('시공사례 글 요청을 접수했습니다', 'success');
    } catch (err) {
      this._toast(err.message, 'error');
    }
    await this.refresh(code);
  }
}
