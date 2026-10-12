/**
 * CasePhotos — PM 아코디언 '시공 완료 사진' 칸 (2026-10-12, 관리자 전용).
 *
 * 사업자등록증 칸처럼 프로젝트 아코디언에서 바로 완료 사진·작업지시서를 올리고
 * [시공사례 글 요청]을 누르면 평일 아침 Claude 예약 작업이 초안을 쓴다.
 * 초안 확인·수정·복사는 /case-posts 페이지(#job=...)에서.
 *
 * 백엔드: /api/case-posts/... (blueprints/case_posts.py). 접수 1건 = job.
 *  - 진행 중 job(사진 보관/접수)이 있으면 거기에 파일을 더하고, 없거나 이미 초안이 나왔으면 새 job.
 * 사진은 긴 변 2000px JPEG 로 줄여 한 장씩 보낸다(프로덕션 요청 8MB 제한).
 * ⚠️ CSP img-src 에 blob: 이 없어 미리보기는 서버 URL 만 쓴다.
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

    const files = (job && job.files) || [];
    const photos = files.filter(f => f.kind === 'photo');
    const orders = files.filter(f => f.kind === 'order');
    const busy = !!this._busy[code];
    const status = job
      ? `<span style="padding:2px 9px;border-radius:10px;font-size:.78rem;font-weight:600;${STATUS_STYLE[job.status] || ''}">${esc(job.status_label)}</span>`
      : '<span style="color:#adb5bd;font-size:.85rem;">아직 없음</span>';
    const thumbs = files.slice(0, 10).map(f => {
      const url = `${API}/jobs/${encodeURIComponent(job.id)}/files/${encodeURIComponent(f.name)}`;
      const inner = /\.pdf$/i.test(f.name)
        ? '<i class="fas fa-file-pdf" style="font-size:1.4rem;color:#e03131;"></i>'
        : `<img src="${url}" loading="lazy" alt="" style="width:100%;height:100%;object-fit:cover;">`;
      return `<a href="${url}" target="_blank" rel="noopener" title="${esc(f.kind === 'order' ? '작업지시서' : '사진')} ${esc(f.original || f.name)}"
                 style="width:52px;height:52px;border-radius:6px;overflow:hidden;border:1px solid ${f.kind === 'order' ? '#ffc9c9' : '#e3e6ea'};
                        display:inline-flex;align-items:center;justify-content:center;background:#f8f9fa;">${inner}</a>`;
    }).join('');
    const more = files.length > 10 ? `<span style="color:#868e96;font-size:.8rem;align-self:center;">+${files.length - 10}</span>` : '';

    const btn = (label, icon, onclick, cls = 'btn-outline-secondary', disabled = false) =>
      `<button type="button" class="btn btn-sm ${cls}" ${disabled ? 'disabled' : ''} onclick="${onclick}">
         <i class="${icon} me-1"></i>${label}</button>`;
    const q = `'${attr(code)}'`;
    const actions = [
      btn('사진 올리기', 'fas fa-images', `window.casePhotos.pick(${q},'photo')`, 'btn-outline-primary', busy),
      btn('작업지시서', 'fas fa-file-alt', `window.casePhotos.pick(${q},'order')`, 'btn-outline-secondary', busy),
    ];
    if (job && job.status === 'uploading') {
      actions.push(btn('시공사례 글 요청', 'fas fa-pen-nib', `window.casePhotos.requestDraft(${q})`, 'btn-primary', busy || !photos.length));
    }
    if (job && job.status !== 'uploading') {
      actions.push(`<a class="btn btn-sm btn-outline-success" href="/case-posts#job=${encodeURIComponent(job.id)}" target="_blank" rel="noopener">
                      <i class="fas fa-external-link-alt me-1"></i>${job.status === 'queued' ? '접수 보기' : '초안 보기'}</a>`);
    }
    const hint = !job ? '사진을 올려 두고, 글로 만들 현장이면 [시공사례 글 요청]을 누르세요. 카톡 PC에서 끌어다 놓아도 됩니다.'
      : job.status === 'queued' ? '다음 평일 아침 7시 반쯤 초안이 만들어집니다.'
      : job.status === 'drafted' ? '초안이 나왔습니다. 사진을 더 올리면 새 접수로 보관됩니다.'
      : '';
    const warn = ((job && job.drive_errors) || []).map(m => `<div style="color:#e8590c;font-size:.8rem;">드라이브 백업: ${esc(m)}</div>`).join('');

    body.innerHTML = `
      <div style="display:flex;flex-wrap:wrap;gap:8px;align-items:center;">
        ${status}
        ${job ? `<span style="color:#495057;font-size:.85rem;">사진 ${photos.length}장${orders.length ? ` · 작업지시서 ${orders.length}` : ''}</span>` : ''}
        <span style="flex:1"></span>
        ${actions.join('')}
      </div>
      ${files.length ? `<div style="display:flex;flex-wrap:wrap;gap:6px;margin-top:8px;">${thumbs}${more}</div>` : ''}
      ${progress ? `<div style="margin-top:6px;font-size:.85rem;color:#1c7ed6;">${esc(progress)}</div>` : ''}
      ${hint ? `<div style="margin-top:6px;font-size:.8rem;color:#868e96;">${esc(hint)}</div>` : ''}
      ${warn}`;
  }

  _bindDrop(card, code) {
    card.dataset.cpBound = '1';
    card.addEventListener('dragover', e => { e.preventDefault(); card.style.outline = '2px dashed #339af0'; });
    card.addEventListener('dragleave', () => { card.style.outline = ''; });
    card.addEventListener('drop', e => {
      e.preventDefault(); card.style.outline = '';
      const all = Array.from(e.dataTransfer.files || []);
      const pdfs = all.filter(f => f.type === 'application/pdf' || /\.pdf$/i.test(f.name));
      const imgs = all.filter(f => !pdfs.includes(f));
      this.upload(code, [...pdfs.map(f => ['order', f]), ...imgs.map(f => ['photo', f])]);
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
