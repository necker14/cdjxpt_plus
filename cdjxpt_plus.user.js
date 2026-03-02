// ==UserScript==
// @name         cdjxpt_plus
// @namespace    cdjxpt-auto
// @version      0.3.9
// @description  报表助手：搬运图片、清空建设节点、删除所有照片、导出全部图片（统一命名）
// @match        https://www.cdjxpt.cn/gyjjddpt/qsmzq-web/*
// @grant        GM_xmlhttpRequest
// @connect      cdnjs.cloudflare.com
// @connect      cdn.jsdelivr.net
// @connect      unpkg.com
// @connect      cdn.bootcdn.net
// @connect      cdn.staticfile.org
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  const CFG_KEY = 'cdjxpt_auto_cfg_v2';
  const SCRIPT_VERSION = '0.3.9';
  const DEFAULT_SAVE_API = 'https://www.cdjxpt.cn/iis/situation/saveSituation.json';
  const DEFAULT_DETAIL_API = 'https://www.cdjxpt.cn/iis/situation/editSituation.json';
  const SITUATION_LIST_KEYS = [
    'investConstructSituations',
    'situationList',
    'situations',
    'projectSituations',
    'items',
  ];
  const NODE_ARRAY_KEY_HINTS = [
    'nodeinfos',
    'nodeinfolist',
    'nodeinfo',
    'monthnodeinfos',
    'currentmonthnodeinfos',
    'completednodeinfos',
    'finishnodeinfos',
    'buildnodeinfos',
    'constructnodeinfos',
    'nodeitems',
    'nodelist',
    'nodes',
  ];
  const NODE_ARRAY_KEY_EXCLUDES = [
    'adjunct',
    'attach',
    'file',
    'image',
    'img',
    'photo',
    'picture',
  ];
  const IMAGE_EXTS = new Set(['jpg', 'jpeg', 'png', 'bmp', 'gif', 'tif', 'tiff', 'webp']);
  const JSZIP_CDN_URLS = [
    'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js',
    'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js',
    'https://unpkg.com/jszip@3.10.1/dist/jszip.min.js',
    'https://cdn.bootcdn.net/ajax/libs/jszip/3.10.1/jszip.min.js',
    'https://cdn.staticfile.org/jszip/3.10.1/jszip.min.js',
  ];

  if (!location.hash.includes('/formFill/')) return;

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const clone = (v) => JSON.parse(JSON.stringify(v));

  let running = false;
  let jsZipLoader = null;

  function parseHashQuery() {
    const h = location.hash || '';
    const i = h.indexOf('?');
    if (i < 0) return {};
    const q = new URLSearchParams(h.slice(i + 1));
    const o = {};
    q.forEach((v, k) => { o[k] = v; });
    return o;
  }

  function parseSearchQuery() {
    const q = new URLSearchParams(location.search || '');
    const o = {};
    q.forEach((v, k) => { o[k] = v; });
    return o;
  }

  function getRouteKey() {
    const h = location.hash || '';
    const path = h.split('?')[0];
    const m = path.match(/#\/formFill\/([^/]+)/);
    return m ? m[1] : '';
  }

  function getContext() {
    const q = { ...parseSearchQuery(), ...parseHashQuery() };
    return {
      type: q.type || '',
      dataId: q.dataId || '',
      reportDate: q.reportDate || '',
      token: q.token ? decodeURIComponent(q.token) : '',
      routeKey: getRouteKey(),
    };
  }

  function loadCfg() {
    try {
      return JSON.parse(localStorage.getItem(CFG_KEY) || '{}');
    } catch {
      return {};
    }
  }

  function saveCfg(cfg) {
    localStorage.setItem(CFG_KEY, JSON.stringify(cfg));
  }

  function setStatus(msg, isErr = false) {
    const el = document.getElementById('cdjxpt-auto-status');
    if (!el) return;
    el.textContent = msg;
    el.style.color = isErr ? '#c0392b' : '#0f5132';
  }

  function setButtonsDisabled(disabled) {
    const ids = [
      'cdjxpt-fill-current',
      'cdjxpt-save-cfg',
      'cdjxpt-copy-photo',
      'cdjxpt-clear-node',
      'cdjxpt-delete-photos',
      'cdjxpt-download-images',
    ];
    ids.forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.disabled = disabled;
    });
  }

  async function runLocked(fn) {
    if (running) {
      setStatus('已有任务在执行，请稍候...', true);
      return;
    }
    running = true;
    setButtonsDisabled(true);
    try {
      await fn();
    } finally {
      running = false;
      setButtonsDisabled(false);
    }
  }

  function isElementVisible(el) {
    if (!el || !(el instanceof Element)) return false;
    const style = window.getComputedStyle(el);
    if (style.display === 'none' || style.visibility === 'hidden') return false;
    const rect = el.getBoundingClientRect();
    return rect.width > 0 && rect.height > 0;
  }

  function findNearestClickable(el) {
    let cur = el;
    for (let i = 0; i < 6 && cur; i++) {
      if (cur.tagName === 'BUTTON' || cur.tagName === 'A') return cur;
      if (cur.getAttribute && cur.getAttribute('role') === 'button') return cur;
      if (typeof cur.onclick === 'function') return cur;
      const cls = String(cur.className || '').toLowerCase();
      if (/(btn|button|icon|delete|trash|remove)/.test(cls)) return cur;
      cur = cur.parentElement;
    }
    return el;
  }

  function looksDeleteElement(el) {
    if (!el || !(el instanceof Element)) return false;
    const cls = String(el.className || '').toLowerCase();
    const title = String(el.getAttribute?.('title') || '');
    const aria = String(el.getAttribute?.('aria-label') || '');
    const text = String(el.textContent || '').trim();
    if (/(delete|trash|remove|el-icon-delete|icon-delete|shanchu|garbage)/i.test(cls)) return true;
    if (/(删除|清除|移除)/.test(`${title} ${aria} ${text}`)) return true;
    if (text === '' || text === '🗑' || text === '🗑️') return true;
    return false;
  }

  function getOwnText(el) {
    if (!el || !(el instanceof Element)) return '';
    return Array.from(el.childNodes || [])
      .filter((n) => n && n.nodeType === Node.TEXT_NODE)
      .map((n) => String(n.textContent || ''))
      .join('');
  }

  function findNodeSectionRoots() {
    const labels = Array.from(document.querySelectorAll('*')).filter((el) => {
      const own = normalizeText(getOwnText(el));
      return own.includes('本月建设节点完成情况') || own.includes('添加本月已完成节点');
    });
    const roots = [];
    const seen = new Set();
    labels.forEach((label) => {
      let cur = label;
      for (let i = 0; i < 10 && cur; i++) {
        const t = normalizeText(cur.textContent || '');
        const iconCount = Array.from(cur.querySelectorAll('*')).filter((el) => (el.textContent || '').trim() === '').length;
        if (
          t.includes('本月建设节点完成情况')
          && t.includes('建设节点')
          && (t.includes('完成时间') || t.includes('添加本月已完成节点'))
          && iconCount > 0
        ) {
          if (!seen.has(cur)) {
            seen.add(cur);
            roots.push(cur);
          }
          return;
        }
        cur = cur.parentElement;
      }
    });
    return roots.filter((root, idx) => !roots.some((other, j) => j !== idx && root.contains(other)));
  }

  function collectNodeDeleteButtons() {
    const roots = findNodeSectionRoots();
    if (!roots.length) return [];
    const seen = new Set();
    const out = [];

    roots.forEach((root) => {
      let raw = Array.from(root.querySelectorAll('i.jxfont')).filter((el) => (el.textContent || '').trim() === '');
      if (!raw.length) {
        raw = Array.from(root.querySelectorAll('*')).filter((el) => {
          if ((el.textContent || '').trim() !== '') return false;
          const cls = String(el.className || '').toLowerCase();
          return cls.includes('delete') || cls.includes('trash') || cls.includes('jxfont');
        });
      }
      raw.forEach((el) => {
        const clickEl = findNearestClickable(el);
        if (!clickEl || seen.has(clickEl)) return;
        if ((clickEl).disabled) return;
        if (!isElementVisible(clickEl)) return;
        if (!root.contains(clickEl)) return;
        seen.add(clickEl);
        out.push(clickEl);
      });
    });
    return out;
  }

  async function clickDialogConfirmIfNeeded() {
    const roots = Array.from(document.querySelectorAll('.el-message-box__wrapper,.el-message-box,.el-popconfirm,.el-popover,[role="dialog"]'))
      .filter(isElementVisible);
    if (!roots.length) return false;
    const buttons = roots.flatMap((r) => Array.from(r.querySelectorAll('button'))).filter(isElementVisible);
    const ok = buttons.find((b) => /(确定|确认|是|删除)/.test(String(b.textContent || '').trim()));
    if (!ok) return false;
    ok.click();
    await sleep(120);
    return true;
  }

  async function clearNodesByUIClick() {
    let totalClicks = 0;
    let sameCountRounds = 0;
    let lastCount = -1;

    for (let round = 0; round < 20; round++) {
      const buttons = collectNodeDeleteButtons();
      if (!buttons.length) break;

      if (buttons.length === lastCount) sameCountRounds += 1;
      else sameCountRounds = 0;
      lastCount = buttons.length;
      if (sameCountRounds >= 3) break;

      let clickedThisRound = 0;
      for (const btn of buttons) {
        if (!isElementVisible(btn)) continue;
        try {
          btn.scrollIntoView({ block: 'center', inline: 'center' });
          await sleep(8);
          btn.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
          btn.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
          btn.click();
          clickedThisRound += 1;
          totalClicks += 1;
          await sleep(8);
          const confirmed = await clickDialogConfirmIfNeeded();
          if (confirmed) await sleep(40);
        } catch (_) {}
      }
      await sleep(80);
      if (!clickedThisRound) break;
    }
    return totalClicks;
  }

  function collectPhotoUploadItems() {
    return Array.from(document.querySelectorAll('.ant-upload-list-item,.el-upload-list__item')).filter((item) => {
      if (!isElementVisible(item)) return false;
      const cls = String(item.className || '').toLowerCase();
      if (cls.includes('upload-select') || cls.includes('upload-btn')) return false;
      return !!(
        item.querySelector('.ant-upload-list-item-thumbnail')
        || item.querySelector('.el-upload-list__item-thumbnail')
        || item.querySelector('img')
      );
    });
  }

  function findPhotoDeleteButtonInItem(item) {
    if (!item) return null;

    const candidates = Array.from(item.querySelectorAll('button,a,span,i')).filter((el) => {
      const cls = String(el.className || '').toLowerCase();
      const title = String(el.getAttribute?.('title') || '');
      const aria = String(el.getAttribute?.('aria-label') || '');
      const text = String(el.textContent || '').trim();
      if (/(删除|delete|移除)/i.test(`${title} ${aria} ${text}`)) return true;
      if (/(anticon-delete|el-icon-delete|icon-delete)/i.test(cls)) return true;
      return text === '' || text === '🗑' || text === '🗑️';
    });

    if (candidates.length) {
      const explicit = candidates.find((el) => {
        const title = String(el.getAttribute?.('title') || '');
        const aria = String(el.getAttribute?.('aria-label') || '');
        const text = String(el.textContent || '').trim();
        return /(删除|delete|移除)/i.test(`${title} ${aria} ${text}`);
      });
      return findNearestClickable(explicit || candidates[candidates.length - 1]);
    }

    const actionBtns = Array.from(item.querySelectorAll('.ant-upload-list-item-card-actions-btn, .el-icon-delete'));
    if (!actionBtns.length) return null;
    return findNearestClickable(actionBtns[actionBtns.length - 1]);
  }

  async function clearPhotosByUIClick() {
    let totalClicks = 0;
    let sameCountRounds = 0;
    let lastCount = -1;

    for (let round = 0; round < 40; round++) {
      const items = collectPhotoUploadItems();
      const count = items.length;
      if (!count) break;

      if (count === lastCount) sameCountRounds += 1;
      else sameCountRounds = 0;
      lastCount = count;
      if (sameCountRounds >= 3) break;

      let clickedThisRound = 0;
      for (const item of items) {
        if (!document.contains(item)) continue;
        try {
          item.scrollIntoView({ block: 'center', inline: 'center' });
          item.dispatchEvent(new MouseEvent('mouseenter', { bubbles: true }));
          item.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
          await sleep(20);

          const btn = findPhotoDeleteButtonInItem(item);
          if (!btn) continue;
          if ((btn).disabled) continue;

          btn.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
          btn.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
          btn.click();

          clickedThisRound += 1;
          totalClicks += 1;
          await sleep(60);

          const confirmed = await clickDialogConfirmIfNeeded();
          if (confirmed) await sleep(80);
        } catch (_) {}
      }

      await sleep(120);
      if (!clickedThisRound) break;
    }

    return totalClicks;
  }

  function absUrl(url) {
    if (!url) return '';
    try {
      return new URL(url, location.origin).toString();
    } catch {
      return '';
    }
  }

  function pathOf(url) {
    try {
      const u = new URL(url, location.origin);
      return `${u.origin}${u.pathname}`;
    } catch {
      return '';
    }
  }

  function sanitizeFilePart(v) {
    return String(v || '')
      .replace(/[\\/:*?"<>|]/g, '_')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function loadScriptByUrl(url) {
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = url;
      s.async = true;
      s.onload = () => resolve();
      s.onerror = () => reject(new Error(`script load failed: ${url}`));
      document.head.appendChild(s);
    });
  }

  function gmRequestText(url) {
    return new Promise((resolve, reject) => {
      if (typeof GM_xmlhttpRequest !== 'function') {
        reject(new Error('GM_xmlhttpRequest unavailable'));
        return;
      }
      GM_xmlhttpRequest({
        method: 'GET',
        url,
        timeout: 20000,
        onload: (resp) => {
          const status = Number(resp?.status || 0);
          if (status >= 200 && status < 300 && resp.responseText) {
            resolve(resp.responseText);
          } else {
            reject(new Error(`HTTP ${status || 'ERR'}`));
          }
        },
        onerror: () => reject(new Error('network error')),
        ontimeout: () => reject(new Error('request timeout')),
      });
    });
  }

  function loadJSZipByCode(code) {
    const fn = new Function(`
${code}
return (typeof JSZip !== 'undefined'
  ? JSZip
  : (typeof globalThis !== 'undefined' ? globalThis.JSZip : undefined));
`);
    const lib = fn();
    if (!lib) throw new Error('JSZip not found after eval');
    try { window.JSZip = lib; } catch (_) {}
    return lib;
  }

  async function ensureJSZip() {
    if (window.JSZip) return window.JSZip;
    if (!jsZipLoader) {
      jsZipLoader = (async () => {
        let lastErr = null;
        for (const url of JSZIP_CDN_URLS) {
          try {
            await loadScriptByUrl(url);
            if (window.JSZip) return window.JSZip;
          } catch (e) {
            lastErr = e;
          }
        }

        for (const url of JSZIP_CDN_URLS) {
          try {
            const code = await gmRequestText(url);
            const lib = loadJSZipByCode(code);
            if (lib) return lib;
          } catch (e) {
            lastErr = e;
          }
        }

        throw lastErr || new Error('压缩组件加载失败，请稍后重试');
      })();
    }
    return jsZipLoader;
  }

  function normalizeText(v) {
    return String(v || '').replace(/\s+/g, '').toLowerCase();
  }

  function normalizeReportNameForCompare(v) {
    return normalizeText(v)
      .replace(/[（(]?\d{4}年\d{1,2}月[)）]?/g, '')
      .replace(/[（(]?\d{4}[-/]\d{1,2}[)）]?/g, '')
      .replace(/[（(]?\d{4}\.\d{1,2}[)）]?/g, '');
  }

  function formatReportPeriod(reportDate) {
    const m = String(reportDate || '').match(/^(\d{4})-(\d{2})/);
    if (!m) return String(reportDate || '未知报表期');
    return `${m[1]}年${m[2]}月`;
  }

  function collectSituationListRefs(obj) {
    if (!obj || typeof obj !== 'object') return [];
    const refs = [];
    const queue = [{ node: obj, path: 'obj' }];
    const visited = new Set();
    while (queue.length) {
      const { node, path } = queue.shift();
      if (!node || typeof node !== 'object' || visited.has(node)) continue;
      visited.add(node);

      for (const key of SITUATION_LIST_KEYS) {
        if (Array.isArray(node[key])) {
          refs.push({ parent: node, key, list: node[key], path: `${path}.${key}` });
        }
      }

      Object.entries(node).forEach(([k, v]) => {
        if (!v || typeof v !== 'object') return;
        if (Array.isArray(v)) {
          v.forEach((it, idx) => {
            if (it && typeof it === 'object') queue.push({ node: it, path: `${path}.${k}[${idx}]` });
          });
        } else {
          queue.push({ node: v, path: `${path}.${k}` });
        }
      });
    }
    return refs;
  }

  function pickBestSituationListRef(refs, expectedLength = 0) {
    if (!refs.length) return null;
    const keyRank = new Map(SITUATION_LIST_KEYS.map((k, i) => [k, i]));
    let best = refs[0];
    let bestScore = -Infinity;
    refs.forEach((ref) => {
      const rank = keyRank.has(ref.key) ? keyRank.get(ref.key) : 99;
      const depth = String(ref.path || '').split('.').length;
      let score = 1000 - rank * 200 - depth * 60;
      if (expectedLength > 0) {
        const diff = Math.abs((ref.list || []).length - expectedLength);
        score += diff === 0 ? 1200 : Math.max(0, 500 - diff * 5);
      } else {
        score += Math.min((ref.list || []).length, 100);
      }
      if (score > bestScore) {
        bestScore = score;
        best = ref;
      }
    });
    return best;
  }

  function extractSituationList(obj, expectedLength = 0) {
    const refs = collectSituationListRefs(obj);
    const picked = pickBestSituationListRef(refs, expectedLength);
    return picked ? picked.list : [];
  }

  function pickString(obj, keys) {
    for (const k of keys) {
      const v = obj?.[k];
      if (typeof v === 'string' && v.trim()) return v.trim();
    }
    return '';
  }

  function getItemKey(item) {
    return pickString(item, ['baseInfoId', 'baseInfoID', 'projectId', 'id']);
  }

  function getItemSecondaryMatchKeys(item) {
    const keys = [
      pickString(item, ['mainMonitorCode', 'mainCode', 'monitorCode']),
      pickString(item, ['mainMonitorCodeName']),
      pickString(item, ['projectCode', 'investCode']),
    ].filter(Boolean).map((v) => normalizeText(v));
    return Array.from(new Set(keys));
  }

  function getProjectName(item, fallbackIndex) {
    const v = pickString(item, [
      'projectName',
      'projectCnName',
      'investConstructName',
      'name',
      'mainMonitorCodeName',
      'projectTitle',
    ]);
    return v || `项目${fallbackIndex + 1}`;
  }

  function getCompanyNameRaw(item) {
    const v = pickString(item, [
      'companyName',
      'enterpriseName',
      'enterpriseCnName',
      'investUnit',
      'buildUnit',
      'projectUnit',
      'ownerName',
      'orgName',
      'company',
    ]);
    return v || '';
  }

  function getCompanyName(item) {
    const v = getCompanyNameRaw(item);
    return v || '未知公司';
  }

  function getProjectCompanyKey(item, fallbackIndex) {
    const projectKey = normalizeText(getProjectName(item, fallbackIndex));
    const companyKey = normalizeText(getCompanyNameRaw(item));
    if (!projectKey || !companyKey) return '';
    return `${projectKey}@@${companyKey}`;
  }

  function getAdjuncts(item) {
    return Array.isArray(item?.adjuncts) ? item.adjuncts : [];
  }

  function isImageAdjunct(adj) {
    if (!adj || typeof adj !== 'object') return false;
    const suffix = String(adj?.suffix || '').replace(/^\./, '').toLowerCase();
    if (suffix && IMAGE_EXTS.has(suffix)) return true;
    const name = String(adj?.originalName || adj?.newName || adj?.name || adj?.fileName || '').toLowerCase();
    if (/\.(jpg|jpeg|png|bmp|gif|tif|tiff|webp)(?:\?|$)/.test(name)) return true;
    const url = String(getAdjunctUrl(adj) || '').toLowerCase();
    if (/\.(jpg|jpeg|png|bmp|gif|tif|tiff|webp)(?:\?|$)/.test(url)) return true;
    const mime = String(adj?.contentType || adj?.mimeType || '').toLowerCase();
    if (mime.startsWith('image/')) return true;
    return false;
  }

  function getPhotoAdjuncts(item) {
    return getAdjuncts(item).filter(isImageAdjunct);
  }

  function splitAdjunctsByImage(adjuncts) {
    const photo = [];
    const other = [];
    (Array.isArray(adjuncts) ? adjuncts : []).forEach((adj) => {
      if (isImageAdjunct(adj)) photo.push(adj);
      else other.push(adj);
    });
    return { photo, other };
  }

  function normalizeKey(v) {
    return String(v || '').replace(/[\s_-]/g, '').toLowerCase();
  }

  function keyLooksNodeArray(key) {
    const k = normalizeKey(key);
    if (!k) return false;
    if (NODE_ARRAY_KEY_EXCLUDES.some((x) => k.includes(x))) return false;
    return NODE_ARRAY_KEY_HINTS.some((x) => k.includes(x));
  }

  function arrayLooksNodeLike(arr) {
    if (!Array.isArray(arr) || arr.length === 0) return false;
    const sample = arr.find((v) => v && typeof v === 'object');
    if (!sample) return false;
    const keys = Object.keys(sample).map((k) => normalizeKey(k));
    if (!keys.length) return false;
    const hints = ['node', 'nodename', 'complete', 'finish', 'month', 'time', 'date', 'construct', 'build', 'progress'];
    let hit = 0;
    keys.forEach((k) => {
      if (hints.some((h) => k.includes(h))) hit += 1;
    });
    return hit >= 2 || keys.some((k) => k === 'node' || k === 'nodename');
  }

  function shouldHandleNodeArray(key, arr) {
    if (!Array.isArray(arr)) return false;
    if (arr.length === 0) return keyLooksNodeArray(key);
    return keyLooksNodeArray(key) || arrayLooksNodeLike(arr);
  }

  function walkNodeArrays(root, clearMode) {
    const visited = new Set();
    let count = 0;
    let fields = 0;

    function walk(node) {
      if (!node || typeof node !== 'object' || visited.has(node)) return;
      visited.add(node);

      if (Array.isArray(node)) {
        node.forEach((it) => {
          if (it && typeof it === 'object') walk(it);
        });
        return;
      }

      Object.entries(node).forEach(([k, v]) => {
        if (Array.isArray(v)) {
          if (shouldHandleNodeArray(k, v)) {
            if (v.length > 0) fields += 1;
            count += v.length;
            if (clearMode) node[k] = [];
            return;
          }
          v.forEach((it) => {
            if (it && typeof it === 'object') walk(it);
          });
          return;
        }
        if (v && typeof v === 'object') walk(v);
      });
    }

    walk(root);
    return { count, fields };
  }

  function clearNodeArraysInItem(item) {
    return walkNodeArrays(item, true);
  }

  function countNodeArraysInItem(item) {
    return walkNodeArrays(item, false);
  }

  function getAdjunctUrl(adj) {
    return adj?.url || adj?.relativePath || adj?.path || adj?.filePath || '';
  }

  function getAdjunctExt(adj, url) {
    const s = String(adj?.suffix || '').trim().toLowerCase();
    if (s) return s.startsWith('.') ? s.slice(1) : s;
    const m = String(url || '').match(/\.([a-zA-Z0-9]+)(?:\?|$)/);
    if (m) return m[1].toLowerCase();
    return 'jpg';
  }

  function addUniqueAdjunctMapEntry(countMap, valueMap, key, adjuncts) {
    if (!key) return;
    countMap.set(key, (countMap.get(key) || 0) + 1);
    if (!valueMap.has(key)) valueMap.set(key, clone(adjuncts));
  }

  function finalizeUniqueAdjunctMap(countMap, valueMap) {
    const out = new Map();
    let collisionCount = 0;
    countMap.forEach((count, key) => {
      if (count === 1) out.set(key, valueMap.get(key));
      else collisionCount += 1;
    });
    return { map: out, collisionCount };
  }

  function buildSourceMaps(sourceObj) {
    const list = extractSituationList(sourceObj);
    const allProjectKeys = new Set();
    const keyCountMap = new Map();
    const keyValueMap = new Map();
    const secondaryCountMap = new Map();
    const secondaryValueMap = new Map();
    const nameCountMap = new Map();
    const nameValueMap = new Map();
    const projectCompanyCountMap = new Map();
    const projectCompanyValueMap = new Map();
    let photoProjectCount = 0;

    list.forEach((it, idx) => {
      const key = getItemKey(it);
      if (key) allProjectKeys.add(key);

      const adjuncts = getPhotoAdjuncts(it);
      if (!adjuncts.length) return;
      photoProjectCount += 1;
      addUniqueAdjunctMapEntry(keyCountMap, keyValueMap, key, adjuncts);
      getItemSecondaryMatchKeys(it).forEach((k) => addUniqueAdjunctMapEntry(secondaryCountMap, secondaryValueMap, k, adjuncts));
      addUniqueAdjunctMapEntry(nameCountMap, nameValueMap, normalizeText(getProjectName(it, idx)), adjuncts);
      addUniqueAdjunctMapEntry(projectCompanyCountMap, projectCompanyValueMap, getProjectCompanyKey(it, idx), adjuncts);
    });

    const { map: byKey, collisionCount: keyCollisionCount } = finalizeUniqueAdjunctMap(keyCountMap, keyValueMap);
    const { map: bySecondaryKey, collisionCount: secondaryCollisionCount } = finalizeUniqueAdjunctMap(secondaryCountMap, secondaryValueMap);
    const { map: byName, collisionCount: nameCollisionCount } = finalizeUniqueAdjunctMap(nameCountMap, nameValueMap);
    const { map: byProjectCompanyKey, collisionCount: projectCompanyCollisionCount } = finalizeUniqueAdjunctMap(projectCompanyCountMap, projectCompanyValueMap);

    return {
      allProjectKeys,
      byKey,
      bySecondaryKey,
      byName,
      byProjectCompanyKey,
      photoProjectCount,
      collisions: {
        key: keyCollisionCount,
        secondary: secondaryCollisionCount,
        name: nameCollisionCount,
        projectCompany: projectCompanyCollisionCount,
      },
    };
  }

  function buildPhotoMapsFromList(list) {
    const byKey = new Map();
    const bySecondaryKey = new Map();
    const byName = new Map();
    const byProjectCompanyKey = new Map();

    (Array.isArray(list) ? list : []).forEach((it, idx) => {
      const photos = getPhotoAdjuncts(it);
      if (!photos.length) return;
      const photoClone = clone(photos);

      const key = getItemKey(it);
      if (key && !byKey.has(key)) byKey.set(key, photoClone);

      const projectCompanyKey = getProjectCompanyKey(it, idx);
      if (projectCompanyKey && !byProjectCompanyKey.has(projectCompanyKey)) byProjectCompanyKey.set(projectCompanyKey, photoClone);

      const nameKey = normalizeText(getProjectName(it, idx));
      if (nameKey && !byName.has(nameKey)) byName.set(nameKey, photoClone);

      getItemSecondaryMatchKeys(it).forEach((k) => {
        if (k && !bySecondaryKey.has(k)) bySecondaryKey.set(k, photoClone);
      });
    });

    return {
      byKey,
      bySecondaryKey,
      byName,
      byProjectCompanyKey,
    };
  }

  function pickPhotosFromMaps(maps, item, idx) {
    const key = getItemKey(item);
    const byKey = key ? maps.byKey.get(key) : null;
    if (byKey && byKey.length) return byKey;

    const byProjectCompany = maps.byProjectCompanyKey.get(getProjectCompanyKey(item, idx));
    if (byProjectCompany && byProjectCompany.length) return byProjectCompany;

    const byName = maps.byName.get(normalizeText(getProjectName(item, idx)));
    if (byName && byName.length) return byName;

    const bySecondary = getItemSecondaryMatchKeys(item).map((k) => maps.bySecondaryKey.get(k)).find(Boolean) || null;
    if (bySecondary && bySecondary.length) return bySecondary;

    return null;
  }

  function syncPhotosToAllPayloadLists(payload, primaryList) {
    const refs = collectSituationListRefs(payload);
    if (!refs.length || !Array.isArray(primaryList)) {
      return { touchedLists: 0, syncedItems: 0 };
    }

    const maps = buildPhotoMapsFromList(primaryList);
    let touchedLists = 0;
    let syncedItems = 0;

    refs.forEach((ref) => {
      const list = ref?.list;
      if (!Array.isArray(list) || list === primaryList) return;
      touchedLists += 1;

      for (let i = 0; i < list.length; i++) {
        const item = list[i];
        const picked = pickPhotosFromMaps(maps, item, i);
        if (!picked || !picked.length) continue;
        const { other } = splitAdjunctsByImage(getAdjuncts(item));
        item.adjuncts = other.concat(clone(picked));
        syncedItems += 1;
      }
    });

    return { touchedLists, syncedItems };
  }

  function extractReportIdentity(detail, ctx) {
    const obj = detail?.obj || {};
    const extra = detail?.extraData || {};

    const reportCode = pickString(obj, ['reportCode', 'tableCode', 'formCode'])
      || pickString(extra, ['reportCode', 'tableCode', 'formCode']);
    const reportId = pickString(obj, ['reportId', 'tableId', 'formId'])
      || pickString(extra, ['reportId', 'tableId', 'formId']);
    const reportName = pickString(obj, ['reportName', 'reportTitle', 'title', 'name'])
      || pickString(extra, ['reportName', 'reportTitle', 'title', 'name'])
      || `route:${ctx.routeKey || 'unknown'}`;

    const list = extractSituationList(obj);
    const firstKeys = list[0] ? Object.keys(list[0]).sort().join('|') : 'none';

    return {
      routeKey: ctx.routeKey || '',
      reportCode,
      reportId,
      reportName,
      shape: `${list.length}|${firstKeys}`,
    };
  }

  function assertCompatible(targetId, sourceId, cfgRouteKey) {
    if (cfgRouteKey && targetId.routeKey && cfgRouteKey !== targetId.routeKey) {
      throw new Error(`源路由(${cfgRouteKey}) 与当前路由(${targetId.routeKey})不一致，已拦截，防止串表`);
    }

    if (targetId.routeKey && sourceId.routeKey && targetId.routeKey !== sourceId.routeKey) {
      throw new Error(`源/目标路由不一致(${sourceId.routeKey} -> ${targetId.routeKey})，已拦截`);
    }

    if (targetId.reportCode && sourceId.reportCode && targetId.reportCode !== sourceId.reportCode) {
      throw new Error(`报表代码不一致(${sourceId.reportCode} -> ${targetId.reportCode})，已拦截`);
    }

    if (targetId.reportId && sourceId.reportId && targetId.reportId !== sourceId.reportId) {
      throw new Error(`报表ID不一致(${sourceId.reportId} -> ${targetId.reportId})，已拦截`);
    }

    if (
      targetId.reportName
      && sourceId.reportName
      && normalizeReportNameForCompare(targetId.reportName) !== normalizeReportNameForCompare(sourceId.reportName)
    ) {
      throw new Error(`报表名称不一致(${sourceId.reportName} -> ${targetId.reportName})，已拦截`);
    }
  }

  function findDetailApiFromPerformance(ctx) {
    try {
      const entries = performance.getEntriesByType('resource') || [];
      const found = entries
        .map((e) => e.name)
        .filter((name) => name.includes('dataId=') && name.includes('reportDate='))
        .filter((name) => name.includes(`dataId=${ctx.dataId}`))
        .filter((name) => name.includes(`reportDate=${ctx.reportDate}`))
        .filter((name) => name.includes('type=edit'))
        .filter((name) => /\.json(\?|$)/.test(name));
      if (!found.length) return '';
      return pathOf(found[found.length - 1]);
    } catch {
      return '';
    }
  }

  function hookSaveCaptureOnceFor(win) {
    if (!win || !win.XMLHttpRequest) throw new Error('页面上下文不可用，无法挂载保存请求捕获器');
    if (win.__cdjxptSaveHooked) return;
    win.__cdjxptSaveHooked = true;
    win.__cdjxptCapturedSaves = [];
    win.__cdjxptSaveResponses = [];
    win.__cdjxptPendingSaveBody = '';

    const XHR = win.XMLHttpRequest;
    const origOpen = XHR.prototype.open;
    const origSend = XHR.prototype.send;

    XHR.prototype.open = function (method, url, ...rest) {
      this.__cdjxptUrl = url;
      this.__cdjxptMethod = method;
      return origOpen.call(this, method, url, ...rest);
    };

    XHR.prototype.send = function (body) {
      let bodyToSend = body;
      try {
        if (
          typeof body === 'string'
          && (this.__cdjxptUrl || '').includes('save')
          && typeof win.__cdjxptPendingSaveBody === 'string'
          && win.__cdjxptPendingSaveBody
        ) {
          bodyToSend = win.__cdjxptPendingSaveBody;
          win.__cdjxptPendingSaveBody = '';
        }

        if (typeof bodyToSend === 'string' && (this.__cdjxptUrl || '').includes('save')) {
          win.__cdjxptCapturedSaves.push({
            method: this.__cdjxptMethod || '',
            url: String(this.__cdjxptUrl || ''),
            body: bodyToSend,
            at: Date.now(),
          });
          this.addEventListener('loadend', () => {
            try {
              win.__cdjxptSaveResponses.push({
                via: 'xhr',
                status: Number(this.status || 0),
                url: String(this.__cdjxptUrl || ''),
                responseText: String(this.responseText || ''),
                at: Date.now(),
              });
            } catch (_) {}
          }, { once: true });
        }
      } catch (_) {}
      return origSend.call(this, bodyToSend ?? body);
    };

    if (typeof win.fetch === 'function' && !win.__cdjxptFetchHooked) {
      win.__cdjxptFetchHooked = true;
      const origFetch = win.fetch.bind(win);
      win.fetch = async function (...args) {
        let callArgs = args;
        try {
          const input = args[0];
          const init = args[1] || {};
          const url = typeof input === 'string' ? input : String(input?.url || '');
          const method = String(init?.method || input?.method || 'GET').toUpperCase();
          let bodyRaw = init?.body;
          if (
            url.includes('save')
            && method !== 'GET'
            && typeof win.__cdjxptPendingSaveBody === 'string'
            && win.__cdjxptPendingSaveBody
          ) {
            bodyRaw = win.__cdjxptPendingSaveBody;
            win.__cdjxptPendingSaveBody = '';
            callArgs = [input, { ...init, body: bodyRaw }];
          }
          const body = typeof bodyRaw === 'string'
            ? bodyRaw
            : (bodyRaw && typeof bodyRaw.toString === 'function' ? bodyRaw.toString() : '');
          if (url.includes('save') && method !== 'GET' && body) {
            win.__cdjxptCapturedSaves.push({
              method,
              url: String(url),
              body,
              at: Date.now(),
            });
          }
        } catch (_) {}
        const resp = await origFetch(...callArgs);
        try {
          const input = callArgs[0];
          const init = callArgs[1] || {};
          const url = typeof input === 'string' ? input : String(input?.url || '');
          const method = String(init?.method || input?.method || 'GET').toUpperCase();
          if (url.includes('save') && method !== 'GET') {
            const cloneResp = resp.clone();
            const txt = await cloneResp.text();
            win.__cdjxptSaveResponses.push({
              via: 'fetch',
              status: Number(resp.status || 0),
              url: String(url),
              responseText: txt,
              at: Date.now(),
            });
          }
        } catch (_) {}
        return resp;
      };
    }
  }

  function findSaveDraftButtonInDocument(doc) {
    if (!doc) return null;
    return Array.from(doc.querySelectorAll('button')).find((b) => (b.textContent || '').includes('存为草稿')) || null;
  }

  function findSaveDraftButton() {
    return findSaveDraftButtonInDocument(document);
  }

  async function ensurePayloadTemplateInWindow(win, label = '当前页面', options = {}) {
    const requireSaveSuccess = options.requireSaveSuccess !== false;
    hookSaveCaptureOnceFor(win);
    const doc = win.document;
    const btn = findSaveDraftButtonInDocument(doc);
    if (!btn) throw new Error(`${label}未找到“存为草稿”按钮`);

    const before = (win.__cdjxptCapturedSaves || []).length;
    const beforeResp = (win.__cdjxptSaveResponses || []).length;
    btn.click();

    for (let i = 0; i < 40; i++) {
      await sleep(200);
      const arr = win.__cdjxptCapturedSaves || [];
      if (arr.length > before) {
        const rec = arr[arr.length - 1];
        let parsed = null;
        try {
          parsed = JSON.parse(rec.body);
        } catch {
          throw new Error(`${label}捕获到保存请求但 JSON 解析失败`);
        }
        if (requireSaveSuccess) {
          for (let j = 0; j < 35; j++) {
            await sleep(150);
            const respArr = win.__cdjxptSaveResponses || [];
            if (respArr.length <= beforeResp) continue;
            const lastResp = respArr[respArr.length - 1];
            let respJson = null;
            try {
              respJson = JSON.parse(String(lastResp?.responseText || '{}'));
            } catch {
              respJson = null;
            }
            if (respJson && respJson.success === false) {
              const code = respJson?.rCode ? `, rCode=${respJson.rCode}` : '';
              if (String(respJson?.msg || '').includes('其他地方已登录') || String(respJson?.rCode || '') === '69') {
                throw new Error(`会话失效：${respJson.msg}${code}。请先退出其它登录并重新登录后再搬运`);
              }
              throw new Error(`${label}存草稿失败：${respJson?.msg || 'unknown'}${code}`);
            }
            break;
          }
        }
        const saveApi = pathOf(rec.url) || DEFAULT_SAVE_API;
        return { payload: parsed, saveApi };
      }
    }

    throw new Error(`${label}未捕获到保存请求，请先确认可正常“存为草稿”`);
  }

  async function ensurePayloadTemplate() {
    return ensurePayloadTemplateInWindow(window, '当前页面', { requireSaveSuccess: true });
  }

  async function captureSourcePayloadViaIframe({ routeKey, dataId, reportDate, token }) {
    if (!routeKey) throw new Error('缺少源路由，无法打开源报表页面');
    const src = `${location.origin}/gyjjddpt/qsmzq-web/#/formFill/${routeKey}?type=edit&dataId=${encodeURIComponent(dataId)}&reportDate=${encodeURIComponent(reportDate)}&token=${encodeURIComponent(token)}`;
    const iframe = document.createElement('iframe');
    iframe.style.cssText = 'position:fixed;left:-99999px;top:-99999px;width:1px;height:1px;opacity:0;pointer-events:none;';
    iframe.src = src;
    document.body.appendChild(iframe);

    try {
      let win = null;
      for (let i = 0; i < 120; i++) {
        await sleep(200);
        win = iframe.contentWindow;
        const doc = iframe.contentDocument;
        if (!win || !doc) continue;
        if (doc.readyState !== 'complete' && doc.readyState !== 'interactive') continue;
        const btn = findSaveDraftButtonInDocument(doc);
        if (btn) break;
      }
      if (!win || !iframe.contentDocument) throw new Error('源报表页面加载失败');
      return await ensurePayloadTemplateInWindow(win, '源报表', { requireSaveSuccess: false });
    } finally {
      iframe.remove();
    }
  }

  async function apiGetDetail({ detailApi, dataId, reportDate, token, type }) {
    const qs = new URLSearchParams({ type, dataId, reportDate, token });
    const url = `${detailApi}?${qs.toString()}`;
    const resp = await fetch(url, { credentials: 'include' });
    const text = await resp.text();
    if (!resp.ok) throw new Error(`GET detail failed: ${resp.status} ${text.slice(0, 120)}`);
    let json = null;
    try {
      json = JSON.parse(text);
    } catch {
      throw new Error(`GET detail non-json response: ${text.slice(0, 120)}`);
    }
    if (!json || json.success !== true) {
      const code = json?.rCode ? `, rCode=${json.rCode}` : '';
      throw new Error(`GET detail error: ${json?.msg || 'unknown'}${code}`);
    }
    return json;
  }

  async function apiSave({ saveApi, payload }) {
    const body = JSON.stringify(payload);
    try {
      const resp = await fetch(saveApi, {
        method: 'POST',
        credentials: 'include',
        headers: { 'content-type': 'application/json;charset=UTF-8' },
        body,
      });

      const text = await resp.text();
      if (!resp.ok) throw new Error(`SAVE failed: ${resp.status} ${text.slice(0, 200)}`);

      let json = null;
      try {
        json = JSON.parse(text);
      } catch {
        throw new Error(`SAVE non-json response: ${text.slice(0, 200)}`);
      }

      if (!json.success) {
        const code = json?.rCode ? `, rCode=${json.rCode}` : '';
        throw new Error(`SAVE error: ${json.msg || 'unknown'}${code}`);
      }
      return json;
    } catch (e) {
      const msg = String(e?.message || e);
      if (!/其他地方已登录|rCode=69|SAVE error/i.test(msg)) throw e;
      return saveByUiInjectedPayload(payload);
    }
  }

  async function saveByUiInjectedPayload(payload) {
    hookSaveCaptureOnceFor(window);
    const btn = findSaveDraftButton();
    if (!btn) throw new Error('未找到“存为草稿”按钮，无法使用页面通道保存');

    const beforeResp = (window.__cdjxptSaveResponses || []).length;
    window.__cdjxptPendingSaveBody = JSON.stringify(payload);
    btn.click();

    for (let i = 0; i < 50; i++) {
      await sleep(200);
      const arr = window.__cdjxptSaveResponses || [];
      if (arr.length > beforeResp) {
        const rec = arr[arr.length - 1];
        let json = null;
        try {
          json = JSON.parse(rec.responseText || '{}');
        } catch {
          throw new Error(`页面通道保存返回非JSON: ${(rec.responseText || '').slice(0, 120)}`);
        }
        if (!json.success) {
          const code = json?.rCode ? `, rCode=${json.rCode}` : '';
          throw new Error(`页面通道保存失败: ${json?.msg || 'unknown'}${code}`);
        }
        return json;
      }
    }

    throw new Error('页面通道保存超时，未收到保存结果');
  }

  function resolveApis(ctx, saveApiFromCapture) {
    const detailFromPerf = findDetailApiFromPerformance(ctx);
    const detailApi = detailFromPerf || DEFAULT_DETAIL_API;
    const saveApi = saveApiFromCapture || DEFAULT_SAVE_API;
    return { detailApi, saveApi };
  }

  function buildPayloadListRef(payload, expectedLength = 0) {
    for (const key of SITUATION_LIST_KEYS) {
      if (Array.isArray(payload?.[key])) return payload[key];
    }
    const refs = collectSituationListRefs(payload);
    const picked = pickBestSituationListRef(refs, expectedLength);
    if (picked) return picked.list;
    if (!Array.isArray(payload.situationList)) payload.situationList = [];
    return payload.situationList;
  }

  async function runAutomation({ doCopyPhotos, doClearNodes }) {
    const ctx = getContext();
    const cfg = loadCfg();

    if (ctx.type !== 'edit') throw new Error('请先进入“当月填报(edit)”页面再运行');
    if (!ctx.dataId || !ctx.reportDate || !ctx.token) throw new Error('当前页面缺少 dataId/reportDate/token');
    if (!doCopyPhotos && !doClearNodes) throw new Error('未选择任何执行动作');

    // 统一规则：清空建设节点始终只走 UI 点击垃圾桶，不走接口改包。
    if (doClearNodes && !doCopyPhotos) {
      const beforeCount = collectNodeDeleteButtons().length;
      setStatus('正在清空建设节点');
      const uiClickCount = await clearNodesByUIClick();
      await sleep(300);
      const remainingUiCount = collectNodeDeleteButtons().length;
      const parts = [
        `清空建设节点：初始 ${beforeCount} 个，模拟点击 ${uiClickCount} 次`,
        `校验：当前剩余节点 ${remainingUiCount} 个`,
        '请手动点击“存为草稿”完成保存',
      ];
      if (remainingUiCount > 0) {
        const failMsg = `${parts.join(' | ')} | 清空校验失败：仍有 ${remainingUiCount} 个节点待删`;
        setStatus(failMsg, true);
        alert(failMsg.replace(/\s\|\s/g, '\n'));
        return;
      }
      setStatus(parts.join(' | '));
      alert(parts.join('\n'));
      return;
    }

    if (!cfg.sourceDataId || !cfg.sourceReportDate) {
      throw new Error('照片迁移需要填写“源 dataId / 源报表期”');
    }

    setStatus('正在准备保存模板...');
    const { payload, saveApi } = await ensurePayloadTemplate();
    const { detailApi } = resolveApis(ctx, saveApi);

    const sourceCtx = { ...ctx, routeKey: cfg.sourceRouteKey || ctx.routeKey };
    let expectedLen = 0;
    let sourceMaps = {
      allProjectKeys: new Set(),
      byKey: new Map(),
      bySecondaryKey: new Map(),
      byName: new Map(),
      byProjectCompanyKey: new Map(),
      photoProjectCount: 0,
      collisions: { key: 0, secondary: 0, name: 0, projectCompany: 0 },
    };
    let sourceFetchFallback = false;
    let sourceFetchNote = '';

    try {
      setStatus('正在抓取当前/源报表详情...');
      const targetDetail = await apiGetDetail({
        detailApi,
        dataId: ctx.dataId,
        reportDate: ctx.reportDate,
        token: ctx.token,
        type: 'edit',
      });
      const sourceDetail = await apiGetDetail({
        detailApi,
        dataId: cfg.sourceDataId,
        reportDate: cfg.sourceReportDate,
        token: ctx.token,
        type: 'read',
      });
      const targetId = extractReportIdentity(targetDetail, ctx);
      const sourceId = extractReportIdentity(sourceDetail, sourceCtx);
      assertCompatible(targetId, sourceId, cfg.sourceRouteKey || '');
      expectedLen = extractSituationList(targetDetail.obj).length || 0;
      sourceMaps = buildSourceMaps(sourceDetail.obj);
    } catch (e) {
      sourceFetchFallback = true;
      sourceFetchNote = String(e?.message || e);
      setStatus('详情接口失败，尝试页面抓取源报表数据...');
      if (cfg.sourceRouteKey && ctx.routeKey && cfg.sourceRouteKey !== ctx.routeKey) {
        throw new Error(`源路由(${cfg.sourceRouteKey}) 与当前路由(${ctx.routeKey})不一致，已拦截`);
      }
      const captured = await captureSourcePayloadViaIframe({
        routeKey: sourceCtx.routeKey || ctx.routeKey,
        dataId: cfg.sourceDataId,
        reportDate: cfg.sourceReportDate,
        token: ctx.token,
      });
      sourceMaps = buildSourceMaps(captured.payload);
    }

    const list = buildPayloadListRef(payload, expectedLen);

    let copied = 0;
    let photoSkipped = 0;
    let fallbackCopied = 0;

    for (let i = 0; i < list.length; i++) {
      const item = list[i];

      const key = getItemKey(item);
      const byKey = key ? sourceMaps.byKey.get(key) : null;
      const keyKnownInSource = !!(key && sourceMaps.allProjectKeys.has(key));
      let picked = byKey || null;

      // 主键存在且在源表可识别时，严格按主键处理；源无图则不降级，避免错配照片。
      if (!picked && !keyKnownInSource) {
        const byProjectCompany = sourceMaps.byProjectCompanyKey.get(getProjectCompanyKey(item, i));
        const byName = sourceMaps.byName.get(normalizeText(getProjectName(item, i)));
        const bySecondary = getItemSecondaryMatchKeys(item).map((k) => sourceMaps.bySecondaryKey.get(k)).find(Boolean) || null;
        picked = byProjectCompany || byName || bySecondary || null;
        if (picked) fallbackCopied += 1;
      }

      if (picked && picked.length > 0) {
        const { other } = splitAdjunctsByImage(getAdjuncts(item));
        item.adjuncts = other.concat(clone(picked));
        copied += 1;
      } else {
        // 未匹配到源图片时，不改当前项目图片，避免误清空
        photoSkipped += 1;
      }
    }

    if (copied === 0 && sourceMaps.photoProjectCount > 0) {
      throw new Error(`未匹配到可搬运图片。源有图项目 ${sourceMaps.photoProjectCount}，目标项目 ${list.length}。请确认源 dataId/报表期是否正确`);
    }

    const syncInfo = syncPhotosToAllPayloadLists(payload, list);

    setStatus('正在保存照片...');
    await apiSave({ saveApi, payload });

    let uiClickCount = 0;
    if (doClearNodes) {
      setStatus('正在清空建设节点');
      uiClickCount = await clearNodesByUIClick();
      await sleep(300);
    }

    setStatus('正在回读校验...');
    let finalVerifyPhotoCount = -1;
    let verifyNote = '';
    try {
      const verify = await apiGetDetail({
        detailApi,
        dataId: ctx.dataId,
        reportDate: ctx.reportDate,
        token: ctx.token,
        type: 'edit',
      });
      const vList = extractSituationList(verify.obj);
      finalVerifyPhotoCount = vList.filter((x) => getPhotoAdjuncts(x).length > 0).length;
    } catch (e) {
      verifyNote = String(e?.message || e);
      finalVerifyPhotoCount = list.filter((x) => getPhotoAdjuncts(x).length > 0).length;
    }
    const remainingUiCount = doClearNodes ? collectNodeDeleteButtons().length : 0;

    const statusParts = [];
    statusParts.push(`照片迁移完成：已搬运 ${copied} 项，未匹配 ${photoSkipped} 项，源有图 ${sourceMaps.photoProjectCount} 项`);
    statusParts.push(`匹配回退 ${fallbackCopied} 项`);
    if (syncInfo.touchedLists > 0) statusParts.push(`多列表同步：${syncInfo.touchedLists} 列表，${syncInfo.syncedItems} 项`);
    if (sourceFetchFallback) statusParts.push('源数据读取：备用模式');
    if (verifyNote) statusParts.push(`校验失败：${verifyNote}`);
    else statusParts.push(`校验：当前有照片项目 ${finalVerifyPhotoCount} 项`);
    if (doClearNodes) statusParts.push(`节点清空：点击 ${uiClickCount} 次，剩余 ${remainingUiCount} 个`);

    const userLines = [];
    userLines.push('照片搬运已完成');
    userLines.push(`已搬运：${copied} 个项目`);
    userLines.push(`未匹配：${photoSkipped} 个项目（保持原样）`);
    if (!verifyNote) {
      userLines.push(`当前有图片的项目：${finalVerifyPhotoCount} 个`);
    } else {
      userLines.push('系统自动校验失败，请刷新页面后检查图片是否显示');
    }
    if (syncInfo.touchedLists > 0) {
      userLines.push('已自动处理不同报表结构的数据同步');
    }
    if (sourceFetchFallback) {
      userLines.push('已自动使用备用方式读取源报表');
    }

    if (doClearNodes) {
      userLines.push(`建设节点删除：已点击 ${uiClickCount} 次，剩余 ${remainingUiCount} 个`);
      userLines.push('建设节点不会自动保存，请手动点击“存为草稿”');
    } else {
      userLines.push('页面即将自动刷新，刷新后可直接看到最新图片');
      userLines.push('如果还没显示，请手动再刷新一次页面');
    }

    if (doClearNodes && remainingUiCount > 0) {
      userLines.push('仍有建设节点未删完，可再点一次“清空建设节点”');
      setStatus(statusParts.join(' | '), true);
      alert(userLines.join('\n'));
      return;
    }

    setStatus(statusParts.join(' | '));
    alert(userLines.join('\n'));
    if (!doClearNodes) {
      setStatus('正在刷新页面显示最新图片...');
      location.reload();
    }
  }

  async function downloadAllImages() {
    const ctx = getContext();
    if (!ctx.dataId || !ctx.reportDate || !ctx.token) throw new Error('当前页面缺少 dataId/reportDate/token');

    const { saveApi } = await ensurePayloadTemplate();
    const { detailApi } = resolveApis(ctx, saveApi);

    setStatus('正在获取当前报表数据...');
    const detail = await apiGetDetail({
      detailApi,
      dataId: ctx.dataId,
      reportDate: ctx.reportDate,
      token: ctx.token,
      type: ctx.type === 'read' ? 'read' : 'edit',
    });

    const identity = extractReportIdentity(detail, ctx);
    const reportName = sanitizeFilePart(identity.reportName || `报表_${ctx.routeKey || 'unknown'}`);
    const reportPeriod = sanitizeFilePart(formatReportPeriod(ctx.reportDate));

    const list = extractSituationList(detail.obj);
    const tasks = [];

    list.forEach((item, idx) => {
      const project = sanitizeFilePart(getProjectName(item, idx));
      const company = sanitizeFilePart(getCompanyName(item));
      const adj = getPhotoAdjuncts(item);
      adj.forEach((a, k) => {
        const rel = getAdjunctUrl(a);
        if (!rel) return;
        const abs = absUrl(rel);
        if (!abs) return;
        const ext = getAdjunctExt(a, rel);
        const suffix = adj.length > 1 ? `-${String(k + 1).padStart(2, '0')}` : '';
        const filename = `${project}-${company}-${reportName}-${reportPeriod}${suffix}.${ext}`;
        tasks.push({ url: abs, filename });
      });
    });

    if (!tasks.length) {
      setStatus('当前报表没有可导出图片');
      alert('当前报表没有可导出图片');
      return;
    }

    setStatus(`正在准备压缩包，共 ${tasks.length} 张图片...`);
    const JSZip = await ensureJSZip();
    const zip = new JSZip();

    const usedNames = new Map();
    let ok = 0;
    let fail = 0;

    for (let i = 0; i < tasks.length; i++) {
      const t = tasks[i];
      let filename = t.filename;

      const cnt = (usedNames.get(filename) || 0) + 1;
      usedNames.set(filename, cnt);
      if (cnt > 1) {
        const dot = filename.lastIndexOf('.');
        if (dot > 0) {
          filename = `${filename.slice(0, dot)}-${String(cnt).padStart(2, '0')}${filename.slice(dot)}`;
        } else {
          filename = `${filename}-${String(cnt).padStart(2, '0')}`;
        }
      }

      try {
        const resp = await fetch(t.url, { credentials: 'include' });
        if (!resp.ok) throw new Error(String(resp.status));
        const blob = await resp.blob();
        zip.file(filename, blob);
        ok += 1;
      } catch (e) {
        fail += 1;
      }

      setStatus(`正在打包 ${i + 1}/${tasks.length}...`);
      await sleep(120);
    }

    if (!ok) {
      const msg = `导出失败：${tasks.length} 张图片都下载失败`;
      setStatus(msg, true);
      alert(msg);
      return;
    }

    setStatus('正在生成压缩包...');
    const zipBlob = await zip.generateAsync({
      type: 'blob',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });
    const zipName = `${reportName}-${reportPeriod}-图片.zip`;
    const a = document.createElement('a');
    const blobUrl = URL.createObjectURL(zipBlob);
    a.href = blobUrl;
    a.download = zipName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(blobUrl);

    const msg = `导出完成：成功 ${ok} 张，失败 ${fail} 张，已下载压缩包`;
    setStatus(msg, fail > 0);
    alert(msg);
  }

  async function deleteAllPhotos() {
    const ctx = getContext();
    if (ctx.type !== 'edit') throw new Error('请先进入“当月填报(edit)”页面再运行');

    const beforeCount = collectPhotoUploadItems().length;
    if (!beforeCount) {
      setStatus('当前页面没有可删除照片');
      alert('当前页面没有可删除照片');
      return;
    }

    setStatus('正在删除所有照片...');
    const clickCount = await clearPhotosByUIClick();
    await sleep(300);
    const remainCount = collectPhotoUploadItems().length;

    const parts = [
      `删除所有照片：初始 ${beforeCount} 张`,
      `模拟点击删除 ${clickCount} 次`,
      `校验：当前剩余 ${remainCount} 张`,
      '请手动点击“存为草稿”完成保存',
    ];

    if (remainCount > 0) {
      const msg = `${parts.join(' | ')} | 仍有照片未删`;
      setStatus(msg, true);
      alert(msg.replace(/\s\|\s/g, '\n'));
      return;
    }

    const msg = parts.join(' | ');
    setStatus(msg);
    alert(msg.replace(/\s\|\s/g, '\n'));
  }

  function mountPanel() {
    if (document.getElementById('cdjxpt-auto-panel')) return;

    const cfg = loadCfg();
    const ctx = getContext();

    const panel = document.createElement('div');
    panel.id = 'cdjxpt-auto-panel';
    panel.style.cssText = [
      'position:fixed',
      'right:14px',
      'bottom:14px',
      'z-index:999999',
      'width:360px',
      'background:#ffffff',
      'border:1px solid #d0d7de',
      'border-radius:10px',
      'box-shadow:0 8px 24px rgba(0,0,0,.16)',
      'padding:10px',
      'font-size:12px',
      'line-height:1.4',
      'color:#1f2328',
    ].join(';');

    panel.innerHTML = `
      <div style="font-weight:700;margin-bottom:8px;">cdjxpt_plus v${SCRIPT_VERSION}</div>
      <div style="margin-bottom:6px;">当前：route=<b>${ctx.routeKey || '-'}</b>，type=<b>${ctx.type || '-'}</b>，reportDate=<b>${ctx.reportDate || '-'}</b></div>
      <label style="display:block;margin:4px 0 2px;">源 dataId（例如 12745117）</label>
      <input id="cdjxpt-source-id" type="text" value="${cfg.sourceDataId || ''}" style="width:100%;box-sizing:border-box;padding:4px 6px;" />
      <label style="display:block;margin:6px 0 2px;">源报表期（YYYY-MM-01）</label>
      <input id="cdjxpt-source-date" type="text" value="${cfg.sourceReportDate || ''}" style="width:100%;box-sizing:border-box;padding:4px 6px;" />
      <div style="display:flex;gap:6px;flex-wrap:wrap;margin-top:8px;">
        <button id="cdjxpt-fill-current" style="padding:4px 8px;">读取当前报表</button>
        <button id="cdjxpt-save-cfg" style="padding:4px 8px;">保存源配置</button>
        <button id="cdjxpt-copy-photo" style="padding:4px 8px;">搬运照片</button>
        <button id="cdjxpt-clear-node" style="padding:4px 8px;">清空建设节点</button>
        <button id="cdjxpt-delete-photos" style="padding:4px 8px;">删除所有照片</button>
        <button id="cdjxpt-download-images" style="padding:4px 8px;">导出所有图片</button>
      </div>
      <div id="cdjxpt-auto-status" style="margin-top:8px;color:#0f5132;word-break:break-word;">待命</div>
      <div style="margin-top:6px;color:#6e7781;">说明：会校验路由/报表身份，避免不同报表串写。</div>
    `;

    document.body.appendChild(panel);

    const sourceIdEl = document.getElementById('cdjxpt-source-id');
    const sourceDateEl = document.getElementById('cdjxpt-source-date');

    document.getElementById('cdjxpt-fill-current').addEventListener('click', () => {
      const now = getContext();
      if (!now.dataId || !now.reportDate) {
        setStatus('当前页面未识别到 dataId/reportDate，请先进入具体报表填报页', true);
        alert('当前页面未识别到 dataId/reportDate，请先进入具体报表填报页。');
        return;
      }

      sourceIdEl.value = now.dataId;
      sourceDateEl.value = now.reportDate;
      saveCfg({
        ...loadCfg(),
        sourceDataId: now.dataId,
        sourceReportDate: now.reportDate,
        sourceRouteKey: now.routeKey || '',
      });
      setStatus(`已读取当前报表：dataId=${now.dataId}，报表期=${now.reportDate}，route=${now.routeKey || '-'}`);
    });

    document.getElementById('cdjxpt-save-cfg').addEventListener('click', () => {
      const now = getContext();
      saveCfg({
        ...loadCfg(),
        sourceDataId: sourceIdEl.value.trim(),
        sourceReportDate: sourceDateEl.value.trim(),
        sourceRouteKey: now.routeKey || '',
      });
      setStatus('源配置已保存');
    });

    document.getElementById('cdjxpt-copy-photo').addEventListener('click', async () => {
      await runLocked(async () => {
        try {
          saveCfg({ ...loadCfg(), sourceDataId: sourceIdEl.value.trim(), sourceReportDate: sourceDateEl.value.trim() });
          await runAutomation({ doCopyPhotos: true, doClearNodes: false });
        } catch (e) {
          console.error(e);
          setStatus(String(e.message || e), true);
          alert(`执行失败：${e.message || e}`);
        }
      });
    });

    document.getElementById('cdjxpt-clear-node').addEventListener('click', async () => {
      await runLocked(async () => {
        try {
          saveCfg({ ...loadCfg(), sourceDataId: sourceIdEl.value.trim(), sourceReportDate: sourceDateEl.value.trim() });
          await runAutomation({ doCopyPhotos: false, doClearNodes: true });
        } catch (e) {
          console.error(e);
          setStatus(String(e.message || e), true);
          alert(`执行失败：${e.message || e}`);
        }
      });
    });

    document.getElementById('cdjxpt-delete-photos').addEventListener('click', async () => {
      await runLocked(async () => {
        try {
          await deleteAllPhotos();
        } catch (e) {
          console.error(e);
          setStatus(String(e.message || e), true);
          alert(`执行失败：${e.message || e}`);
        }
      });
    });

    document.getElementById('cdjxpt-download-images').addEventListener('click', async () => {
      await runLocked(async () => {
        try {
          await downloadAllImages();
        } catch (e) {
          console.error(e);
          setStatus(String(e.message || e), true);
          alert(`导出失败：${e.message || e}`);
        }
      });
    });
  }

  mountPanel();
})();
