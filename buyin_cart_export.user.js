// ==UserScript==
// @name         百应选品车一键抓取导出（独立版）
// @namespace    ge.buyin.cart.export
// @version      4.7.1
// @description  巨量百应/罗盘选品车：预加载微轮次追底+卡片数进展；至「没有更多」；仅滚列表不滚整页；飞书写入
// @author       you
// @match        https://buyin.jinritemai.com/*
// @match        https://compass.jinritemai.com/*
// @run-at       document-start
// @grant        GM_addStyle
// @grant        GM_xmlhttpRequest
// @grant        GM_setValue
// @grant        GM_getValue
// @connect      open.feishu.cn
// ==/UserScript==

(function () {
  'use strict';

  const isJinritemaiCart = /^(buyin|compass)\.jinritemai\.com$/i.test(location.hostname);
  if (!isJinritemaiCart) return;

  /** 需要「字段调试」/addRows 日志时改为 true */
  const GE_CART_DEBUG = false;
  /** 仅在一次「扫描页面」/自动扫描任务内为 true */
  let GE_CART_SCAN_ACTIVE = false;
  /** 与 GE_CART_DEBUG 配合，每次扫描仅输出前 10 条「字段调试」日志 */
  let GE_FIELD_DEBUG_LOG_COUNT = 0;

  /** 为 true 时使用 getCardsInDocV2（按可见 div + 特征列） */
  const GE_USE_CARDS_V2 = true;

  // ---------- 拦截选品车列表 JSON（含 cart_view_card），比纯 DOM 更稳 ----------
  function rowFromApiProductCard(pc) {
    if (!pc || typeof pc !== 'object') return null;
    const base = pc.base_product_info || {};
    const biz = pc.business_operation_info || {};
    const shop = pc.base_shop_info || {};
    const priceFen = Number(base.price);
    let handPrice = '';
    if (Number.isFinite(priceFen)) {
      handPrice = (priceFen / 100).toFixed(2);
    }
    const cosRatio = Number(biz.cos_ratio);
    let commissionRate = '';
    if (Number.isFinite(cosRatio)) {
      const pct = cosRatio / 100;
      commissionRate = (pct % 1 === 0 ? String(Math.round(pct)) : pct.toFixed(2)) + '%';
    }
    const pid = s(pc.product_id || base.product_id);
    const detail = s(base.detail_url);
    const sid = shop.shop_id;
    const shopIdStr = sid != null && sid !== '' ? String(sid) : '';
    let shopLink = '';
    if (shopIdStr && isValidRealShopId(shopIdStr)) {
      shopLink = normalizeUrl(
        'https://buyin.jinritemai.com/dashboard/merch-picking-library/shop-detail?shop_id=' + encodeURIComponent(shopIdStr)
      );
    }
    let goodRate = s(biz.good_ratio);
    if (goodRate && goodRate.indexOf('%') < 0) goodRate = goodRate + '%';
    const img = s(base.cover);
    return {
      productInfo: s(base.title),
      productId: pid,
      productLink: detail ? normalizeUrl(detail) : '',
      image: img,
      productImageLink: img,
      shop: s(shop.shop_name),
      shopId: shopIdStr,
      shopLink,
      commissionRate,
      handPrice,
      monthlySales: biz.monthly_sale != null ? String(biz.monthly_sale) : '',
      goodRate,
      experienceScore: s(shop.exp_score),
      source: 'api_cart',
    };
  }

  function ingestCartApiJson(json) {
    if (!json || json.code !== 0 || !json.data) return;
    const d = json.data;
    const views = d.cart_view_card;
    if (!Array.isArray(views)) return;
    const rows = [];
    for (let vi = 0; vi < views.length; vi++) {
      const v = views[vi];
      const cards = v && v.cards;
      if (!Array.isArray(cards)) continue;
      for (let ci = 0; ci < cards.length; ci++) {
        const c = cards[ci];
        if (c && c.entity_type === 1 && c.product_card) {
          const r = rowFromApiProductCard(c.product_card);
          if (r) rows.push(r);
        }
      }
    }
    if (rows.length) {
      addRows(rows);
      try {
        updateStatus();
      } catch (_) {}
    }
  }

  function tryIngestCartResponseText(text) {
    if (!text || String(text).indexOf('cart_view_card') < 0) return;
    let j;
    try {
      j = JSON.parse(text);
    } catch (_) {
      return;
    }
    ingestCartApiJson(j);
  }

  (function installCartApiHooks() {
    try {
      const origOpen = XMLHttpRequest.prototype.open;
      const origSend = XMLHttpRequest.prototype.send;
      XMLHttpRequest.prototype.open = function (method, url) {
        try {
          this._geCartXhrUrl = typeof url === 'string' ? url : '';
        } catch (_) {
          this._geCartXhrUrl = '';
        }
        return origOpen.apply(this, arguments);
      };
      XMLHttpRequest.prototype.send = function () {
        const xhr = this;
        xhr.addEventListener(
          'load',
          function () {
            try {
              tryIngestCartResponseText(xhr.responseText);
            } catch (_) {}
          },
          { once: true }
        );
        return origSend.apply(this, arguments);
      };
    } catch (_) {}
    try {
      const ofetch = window.fetch;
      if (typeof ofetch !== 'function') return;
      window.fetch = function () {
        const p = ofetch.apply(this, arguments);
        try {
          const req = arguments[0];
          let u = typeof req === 'string' ? req : req && req.url;
          u = u ? String(u) : '';
          const rel = u && !/^https?:/i.test(u) ? location.origin + (u[0] === '/' ? u : '/' + u) : u;
          const low = rel.toLowerCase();
          if (
            rel &&
            (/buyin\.jinritemai\.com/i.test(rel) || rel.indexOf(location.origin) === 0) &&
            (low.indexOf('cart') >= 0 ||
              low.indexOf('picking') >= 0 ||
              low.indexOf('merch') >= 0 ||
              low.indexOf('selection') >= 0)
          ) {
            return p.then(function (res) {
              try {
                const cl = res.clone();
                cl.text().then(tryIngestCartResponseText).catch(function () {});
              } catch (_) {}
              return res;
            });
          }
        } catch (_) {}
        return p;
      };
    } catch (_) {}
  })();

  /** 无真实商品链接时，用商品 ID 拼接（你提供的模板） */
  const PRODUCT_LINK_TEMPLATE =
    'https://buyin.jinritemai.com/dashboard/merch-picking-library/merch-promoting?id=';

  const STORE = {
    rows: [],
    keySet: new Set(),
    mounted: false,
  };

  const BAD_LINK_TEXT_RE = /添加分销|联系商家|橱窗|删除|操作|管理/i;

  function s(v) {
    return v == null ? '' : String(v).trim();
  }

  function toNumLike(v) {
    const t = s(v).replace(/[,\s]/g, '');
    if (!t) return '';
    const m = t.match(/(\d+(?:\.\d+)?)/);
    return m ? m[1] : '';
  }

  function normalizeUrl(u) {
    const x = s(u);
    if (!x) return '';
    if (/^https?:\/\//i.test(x)) return x;
    if (/^\//.test(x)) return 'https://buyin.jinritemai.com' + x;
    return x;
  }

  function isHttp(u) {
    return /^https?:\/\//i.test(s(u));
  }

  function toAbs(u) {
    if (!u) return '';
    const x = s(u);
    if (x.startsWith('//')) return 'https:' + x;
    if (x.startsWith('/')) return location.origin + x;
    return x;
  }

  function clean(v) {
    return s(v).replace(/\s+/g, ' ');
  }

  function text(el) {
    return clean(el && (el.innerText || el.textContent));
  }

  function attr(el, name) {
    return s(el && el.getAttribute && el.getAttribute(name));
  }

  function isValidRealProductId(id) {
    if (!id) return false;
    const t = s(id);
    if (!/^\d{8,22}$/.test(t)) return false;
    if (/^1000+/.test(t)) return false;
    return true;
  }

  function isValidRealShopId(id) {
    if (!id) return false;
    const t = s(id);
    if (!/^\d{6,22}$/.test(t)) return false;
    if (/^1000+/.test(t)) return false;
    return true;
  }

  /** 页面常用「-」「—」表示无 ID，与空字符串同等视为无值 */
  function geCartIdFieldEmpty(v) {
    const t = s(v);
    if (!t) return true;
    if (/^[-—－]+$/.test(t)) return true;
    return false;
  }

  /** 失效：商品ID、店铺ID 均为空或占位横线（不依赖位数校验） */
  function geCartRowBothIdsEmpty(r) {
    if (!r || typeof r !== 'object') return true;
    return geCartIdFieldEmpty(r.productId) && geCartIdFieldEmpty(r.shopId);
  }

  /** 可导出/可入库：至少填了商品ID或店铺ID之一（非双空） */
  function geIsValidCartExportRow(r) {
    if (!r || typeof r !== 'object') return false;
    return !geCartRowBothIdsEmpty(r);
  }

  const geSleep = (ms) =>
    new Promise(function (r) {
      setTimeout(r, ms);
    });

  function rowDedupeKey(r) {
    const pid = s(r.productId);
    if (isValidRealProductId(pid)) return 'pid:' + pid;
    return [s(r.productInfo), s(r.shop), s(r.handPrice), s(r.monthlySales)].join('|');
  }

  function dedupeFinalRows(rows) {
    const out = [];
    const seen = new Set();
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!geIsValidCartExportRow(r)) continue;
      const k = rowDedupeKey(r);
      if (!k || seen.has(k)) continue;
      seen.add(k);
      out.push(r);
    }
    return out;
  }

  /** 有商品ID且无真实链接时，用模板补全 */
  function fillProductLinkFromId(row) {
    if (!row) return row;
    const pid = s(row.productId);
    if (!s(row.productLink) && isValidRealProductId(pid)) {
      row.productLink = PRODUCT_LINK_TEMPLATE + encodeURIComponent(pid);
    }
    return row;
  }

  function mergeRowField(oldV, newV) {
    if (s(oldV)) return oldV;
    return newV;
  }

  function addRows(rows) {
    let inc = 0;
    let mergedDup = 0;
    let skipped = 0;
    const skipReasons = [];
    if (GE_CART_DEBUG && GE_CART_SCAN_ACTIVE) {
      console.log('[addRows] incoming=', rows ? rows.length : 0);
    }
    for (let i = 0; i < rows.length; i++) {
      let r = rows[i];
      if (!r) {
        skipped++;
        if (GE_CART_DEBUG && GE_CART_SCAN_ACTIVE) skipReasons.push('i' + i + ':null');
        continue;
      }
      fillProductLinkFromId(r);
      if (!geIsValidCartExportRow(r)) {
        skipped++;
        if (GE_CART_DEBUG && GE_CART_SCAN_ACTIVE) skipReasons.push('i' + i + ':bothIdsEmpty');
        continue;
      }
      if (!s(r.productInfo) && !s(r.productId) && !s(r.image) && !s(r.productImageLink)) {
        skipped++;
        if (GE_CART_DEBUG && GE_CART_SCAN_ACTIVE) skipReasons.push('i' + i + ':noTitleIdImg');
        continue;
      }
      if (/\.(png|jpe?g|webp|gif)(\?|$)/i.test(s(r.productInfo)) || /^https?:\/\//i.test(s(r.productInfo))) {
        skipped++;
        if (GE_CART_DEBUG && GE_CART_SCAN_ACTIVE) skipReasons.push('i' + i + ':badTitle');
        continue;
      }
      if (!s(r.productInfo) && !s(r.shop) && !s(r.productId)) {
        skipped++;
        if (GE_CART_DEBUG && GE_CART_SCAN_ACTIVE) skipReasons.push('i' + i + ':noCore');
        continue;
      }

      const k = rowDedupeKey(r);
      if (GE_CART_DEBUG && GE_CART_SCAN_ACTIVE) {
        console.log('[addRows] key=', k, 'productInfo=', s(r.productInfo).slice(0, 48), 'productId=', s(r.productId));
      }
      if (!k || k === '|||') {
        skipped++;
        if (GE_CART_DEBUG && GE_CART_SCAN_ACTIVE) skipReasons.push('i' + i + ':emptyKey');
        continue;
      }

      if (STORE.keySet.has(k)) {
        mergedDup++;
        const idx = STORE.rows.findIndex(function (x) {
          return rowDedupeKey(x) === k;
        });
        if (idx >= 0) {
          const cur = STORE.rows[idx];
          STORE.rows[idx] = {
            ...cur,
            productLink: mergeRowField(cur.productLink, r.productLink),
            productId: mergeRowField(cur.productId, r.productId),
            handPrice: mergeRowField(cur.handPrice, r.handPrice),
            commissionRate: mergeRowField(cur.commissionRate, r.commissionRate),
            monthlySales: mergeRowField(cur.monthlySales, r.monthlySales),
            goodRate: mergeRowField(cur.goodRate, r.goodRate),
            experienceScore: mergeRowField(cur.experienceScore, r.experienceScore),
            shop: mergeRowField(cur.shop, r.shop),
            shopId: mergeRowField(cur.shopId, r.shopId),
            shopLink: mergeRowField(cur.shopLink, r.shopLink),
            image: mergeRowField(cur.image, r.image),
            productImageLink: mergeRowField(cur.productImageLink, r.productImageLink),
            productInfo: mergeRowField(cur.productInfo, r.productInfo),
            source: mergeRowField(cur.source, r.source),
          };
          fillProductLinkFromId(STORE.rows[idx]);
        }
        continue;
      }
      STORE.keySet.add(k);
      STORE.rows.push(r);
      inc++;
    }
    if (GE_CART_DEBUG && GE_CART_SCAN_ACTIVE) {
      console.log(
        '[addRows] 新增入库=',
        inc,
        '合并已存在=',
        mergedDup,
        '跳过=',
        skipped,
        'STORE.rows.length=',
        STORE.rows.length
      );
      if (skipReasons.length) console.log('[addRows] skipReasons(前30)=', skipReasons.slice(0, 30));
    }
    if (inc > 0) updateStatus();
    return inc;
  }

  /** 子树内「商品信息块」数量（含自身若匹配），querySelectorAll 不含祖先自身故需单独算 */
  function countSelectedPrdInfoInSubtree(el) {
    if (!el) return 0;
    let n = 0;
    try {
      const cls = el.className;
      const cs = typeof cls === 'string' ? cls : cls && cls.baseVal != null ? String(cls.baseVal) : '';
      if (cs.indexOf('selectedPrdInfo') >= 0) n++;
      n += el.querySelectorAll('div[class*="selectedPrdInfo"]').length;
    } catch (_) {}
    return n;
  }

  function hasRowFeatures(el) {
    if (!el) return false;
    return (
      !!el.querySelector('[class*="priceWrap"]') ||
      !!el.querySelector('[class*="ratioWrapper"]') ||
      !!el.querySelector('[class*="evaluationWrap"]') ||
      !!el.querySelector('[class*="btnWrap"]') ||
      !!el.querySelector('[class*="shopTitle"]') ||
      !!el.querySelector('[class*="shopItem"]')
    );
  }

  function hasRowClassHint(el) {
    if (!el) return false;
    const cs = typeof el.className === 'string' ? el.className : String(el.className || '');
    return (
      cs.indexOf('tableLine') >= 0 ||
      cs.indexOf('tableItem') >= 0 ||
      cs.indexOf('lineItem') >= 0 ||
      cs.indexOf('listItem') >= 0 ||
      cs.indexOf('tableRow') >= 0 ||
      (cs.indexOf('row') >= 0 && (cs.indexOf('table') >= 0 || cs.indexOf('Line') >= 0 || cs.indexOf('List') >= 0))
    );
  }

  /** 从 selectedPrdInfo 向上找「单行」根：子树内仅 1 个商品信息块 + 含店铺/价格等列；优先带 tableLine/tableItem 等行 class */
  function findCardRoot(info, idxForDebug) {
    if (!info) return null;
    let p = info;
    for (let i = 0; i < 48 && p; i++) {
      const infoCount = countSelectedPrdInfoInSubtree(p);
      if (infoCount !== 1) {
        p = p.parentElement;
        continue;
      }
      if (!hasRowFeatures(p)) {
        p = p.parentElement;
        continue;
      }
      if (hasRowClassHint(p)) {
        if (GE_CART_DEBUG && GE_CART_SCAN_ACTIVE && idxForDebug < 5) {
          console.log('[findCardRoot]', {
            idx: idxForDebug,
            infoText: text(info).slice(0, 50),
            root: p,
            rootClass: typeof p.className === 'string' ? String(p.className).slice(0, 100) : '',
            infoCount: infoCount,
          });
        }
        return p;
      }
      p = p.parentElement;
    }
    p = info;
    for (let i = 0; i < 48 && p; i++) {
      const infoCount = countSelectedPrdInfoInSubtree(p);
      if (infoCount !== 1) {
        p = p.parentElement;
        continue;
      }
      if (!hasRowFeatures(p)) {
        p = p.parentElement;
        continue;
      }
      if (GE_CART_DEBUG && GE_CART_SCAN_ACTIVE && idxForDebug < 5) {
        console.log('[findCardRoot]', {
          idx: idxForDebug,
          infoText: text(info).slice(0, 50),
          root: p,
          rootClass: typeof p.className === 'string' ? String(p.className).slice(0, 100) : '',
          infoCount: infoCount,
          note: 'fallback-no-row-class',
        });
      }
      return p;
    }
    const nearestLine = info.closest(
      'div[class*="tableLine"],div[class*="tableItem"],div[class*="lineItem"],div[class*="listItem"],div[class*="row"]'
    );
    if (nearestLine && countSelectedPrdInfoInSubtree(nearestLine) === 1) {
      if (GE_CART_DEBUG && GE_CART_SCAN_ACTIVE && idxForDebug < 5) {
        console.log('[findCardRoot]', {
          idx: idxForDebug,
          infoText: text(info).slice(0, 50),
          root: nearestLine,
          infoCount: countSelectedPrdInfoInSubtree(nearestLine),
          note: 'closest',
        });
      }
      return nearestLine;
    }
    const last = info.parentElement || info;
    if (GE_CART_DEBUG && GE_CART_SCAN_ACTIVE && idxForDebug < 5) {
      console.log('[findCardRoot]', {
        idx: idxForDebug,
        infoText: text(info).slice(0, 50),
        root: last,
        infoCount: countSelectedPrdInfoInSubtree(last),
        note: 'last-resort',
      });
    }
    return last;
  }

  function geUniqElements(arr) {
    const uniq = [];
    const seen = new Set();
    for (let i = 0; i < arr.length; i++) {
      const el = arr[i];
      if (!el || seen.has(el)) continue;
      seen.add(el);
      uniq.push(el);
    }
    return uniq;
  }

  function getDomPath(el) {
    if (!el) return '';
    const parts = [];
    let cur = el;
    let depth = 0;
    while (cur && cur.nodeType === 1 && depth < 8) {
      let idx = 1;
      let sib = cur;
      while ((sib = sib.previousElementSibling)) idx++;
      parts.unshift(cur.tagName.toLowerCase() + ':nth-child(' + idx + ')');
      cur = cur.parentElement;
      depth++;
    }
    return parts.join(' > ');
  }

  function geIsElementVisible(el) {
    if (!el || el.nodeType !== 1) return false;
    try {
      const win = el.ownerDocument.defaultView;
      if (!win) return false;
      const st = win.getComputedStyle(el);
      if (st.display === 'none' || st.visibility === 'hidden' || parseFloat(st.opacity) === 0) return false;
      const r = el.getBoundingClientRect();
      return r.width > 0 && r.height > 0;
    } catch (_) {
      return false;
    }
  }

  /** 店铺/佣金/到手价/月销好评/操作 五类特征中命中几类（用于 V2） */
  function countFeatureColumnHits(el) {
    if (!el) return 0;
    let n = 0;
    if (el.querySelector('[class*="shopTitle"]') || el.querySelector('[class*="shopItem"]')) n++;
    if (el.querySelector('[class*="ratioWrapper"]')) n++;
    if (el.querySelector('[class*="priceWrap"]')) n++;
    if (el.querySelector('[class*="evaluationWrap"]')) n++;
    if (el.querySelector('[class*="btnWrap"]')) n++;
    return n;
  }

  /**
   * 不依赖 selectedPrdInfo 反推 root：遍历可见 div，子树含 selectedPrdInfo 且至少命中 3 类列特征；
   * 去掉「仍包含其它候选节点」的外层大容器，保留最内层行节点。
   */
  function getCardsInDocV2(root) {
    const doc = root && root.querySelectorAll ? root : document;
    const candidates = [];
    const allDivs = doc.querySelectorAll('div');
    for (let i = 0; i < allDivs.length; i++) {
      const el = allDivs[i];
      if (!geIsElementVisible(el)) continue;
      if (!el.querySelector('div[class*="selectedPrdInfo"]')) continue;
      if (countFeatureColumnHits(el) < 3) continue;
      candidates.push(el);
    }
    return candidates.filter(function (el) {
      return !candidates.some(function (other) {
        return other !== el && el.contains(other);
      });
    });
  }

  /** 直接收集「行级」容器：每行仅 1 个 selectedPrdInfo，且含价格/店铺等列；去掉被外层行包含的内层重复节点 */
  function collectRowCandidatesFromDOM(root) {
    const sel =
      'div[class*="tableLine"],div[class*="tableItem"],div[class*="lineItem"],div[class*="listItem"],div[class*="row"]';
    const all = Array.from(root.querySelectorAll(sel));
    const out = [];
    for (let i = 0; i < all.length; i++) {
      const el = all[i];
      if (countSelectedPrdInfoInSubtree(el) !== 1) continue;
      if (!el.querySelector('div[class*="selectedPrdInfo"]')) continue;
      if (!hasRowFeatures(el)) continue;
      out.push(el);
    }
    return out.filter(function (el) {
      return !out.some(function (other) {
        return other !== el && other.contains(el);
      });
    });
  }

  function getCardsInDoc(d) {
    const root = d || document;
    if (GE_USE_CARDS_V2) {
      const v2 = getCardsInDocV2(root);
      if (GE_CART_DEBUG && GE_CART_SCAN_ACTIVE) {
        console.log('[getCardsInDocV2] 行节点数=', v2.length);
      }
      return v2;
    }
    const infos = Array.from(root.querySelectorAll('div[class*="selectedPrdInfo"]'));
    const rowDirect = collectRowCandidatesFromDOM(root);
    if (GE_CART_DEBUG && GE_CART_SCAN_ACTIVE) {
      console.log('[getCardsInDoc] info节点数=', infos.length);
      console.log('[getCardsInDoc] 行容器直选=', rowDirect.length);
    }
    if (rowDirect.length >= 1 && (infos.length === 0 || rowDirect.length >= infos.length)) {
      const uniq = geUniqElements(rowDirect);
      if (GE_CART_DEBUG && GE_CART_SCAN_ACTIVE) {
        console.log('[getCardsInDoc] card root去重前=', rowDirect.length);
        console.log('[getCardsInDoc] card root去重后=', uniq.length);
      }
      return uniq;
    }
    const cards = infos
      .map(function (info, idx) {
        return findCardRoot(info, idx);
      })
      .filter(Boolean);
    if (GE_CART_DEBUG && GE_CART_SCAN_ACTIVE) {
      console.log('[getCardsInDoc] card root去重前=', cards.length);
    }
    const uniq2 = geUniqElements(cards);
    if (GE_CART_DEBUG && GE_CART_SCAN_ACTIVE) {
      console.log('[getCardsInDoc] card root去重后=', uniq2.length);
    }
    return uniq2;
  }

  function getImage(card) {
    const imgs = Array.from(card.querySelectorAll('img'));
    for (let i = 0; i < imgs.length; i++) {
      const img = imgs[i];
      const cands = [
        img.src,
        attr(img, 'src'),
        attr(img, 'data-src'),
        attr(img, 'data-lazy-src'),
        attr(img, 'data-original'),
        attr(img, 'data-url'),
      ].filter(Boolean);
      for (let j = 0; j < cands.length; j++) {
        const abs = toAbs(cands[j]);
        if (isHttp(abs)) return normalizeUrl(abs);
      }
      const srcset = attr(img, 'srcset');
      if (srcset) {
        const first = (srcset.split(',')[0] || '').trim().split(/\s+/)[0];
        if (first) {
          const abs = toAbs(first);
          if (isHttp(abs)) return normalizeUrl(abs);
        }
      }
    }
    const all = card.querySelectorAll('*');
    for (let i = 0; i < all.length; i++) {
      try {
        const bg = getComputedStyle(all[i]).backgroundImage || '';
        const m = bg.match(/url\(["']?(.*?)["']?\)/);
        if (m && m[1]) {
          const abs = toAbs(m[1]);
          if (isHttp(abs)) return normalizeUrl(abs);
        }
      } catch (_) {}
    }
    return '';
  }

  function getTitle(card) {
    const titleEl =
      card.querySelector('div[class*="selectedPrdTitle"]') ||
      card.querySelector('div[class*="titleHover"]') ||
      Array.from(card.querySelectorAll('div[class*="selectedPrdDescContent"] div')).find(function (el) {
        const t = text(el);
        return t && t.length >= 6 && t !== 'ID' && !/^店铺$/.test(t);
      });
    return text(titleEl);
  }

  function getShopName(card, title) {
    const shopEl =
      card.querySelector('div[class*="shopTitle"]') ||
      Array.from(card.querySelectorAll('div[class*="shopItem"] div, div[class*="selectedPrdDesc"] div')).find(function (el) {
        const t = text(el);
        return t && t !== title && t !== 'ID' && t.length <= 50;
      });
    return text(shopEl);
  }

  function getExpectedEarn(card) {
    const wrap = card.querySelector('div[class*="priceWrap"]');
    if (!wrap) return '';
    const all = text(wrap);
    const m = all.match(/预计赚\s*[¥￥]?\s*([\d.]+)/);
    return m ? m[1] : '';
  }

  function extractHandPriceFromPriceWrap(w) {
    if (!w) return '';
    const raw = text(w);
    if (/预计赚/.test(raw)) return '';
    const joined = Array.from(w.querySelectorAll('span,div'))
      .map(text)
      .filter(Boolean)
      .join(' ');
    let m = joined.match(/[¥￥]\s*([\d.]+)/);
    if (!m) m = raw.match(/[¥￥]\s*([\d.]+)/);
    if (m) return m[1];
    const priceEl = w.querySelector('span[class*="priceML"]') || w.querySelector('span[class*="price"]');
    if (priceEl) {
      const pt = text(priceEl);
      if (/预计赚/.test(pt)) return '';
      const mm = pt.match(/([\d.]+)/);
      if (mm) return mm[1];
    }
    return '';
  }

  function normalizePriceText(raw) {
    let t = s(raw);
    if (!t) return '';
    return t
      .replace(/[。．]/g, '.')
      .replace(/[¥￥]/g, '')
      .replace(/\s+/g, '')
      .replace(/预计赚/g, '')
      .trim();
  }

  function extractPriceFromNode(node) {
    if (!node) return '';
    const raw = text(node);
    const norm = normalizePriceText(raw);
    if (!norm) return '';
    let m = norm.match(/(\d+\.\d{1,2})/);
    if (m) return m[1];
    const nums = norm.match(/\d+/g);
    if (nums && nums.length >= 2) {
      const a = nums[0];
      const b = nums[1];
      if (a && b && b.length <= 2) {
        var bp = b.length === 1 ? '0' + b : b;
        return a + '.' + bp;
      }
    }
    if (nums && nums.length === 1) {
      return nums[0];
    }
    return '';
  }

  function isLikelyPrice(v) {
    const t = s(v);
    if (!t) return false;
    if (/%/.test(t)) return false;
    if (/万/.test(t)) return false;
    const n = Number(t);
    return Number.isFinite(n) && n > 0 && n < 100000;
  }

  /** 到手价：按列优先（priceWrap → weightWrap → cosratio），不做整卡 ¥ 兜底 */
  function getPrice(card) {
    if (!card) return '';

    const priceWraps = Array.from(card.querySelectorAll('div[class*="priceWrap"], span[class*="priceWrap"]'));
    let i;
    for (i = 0; i < priceWraps.length; i++) {
      const el = priceWraps[i];
      if (el.closest('[class*="ratioWrapper"]')) continue;
      let v = extractPriceFromNode(el);
      if (!isLikelyPrice(v)) v = extractHandPriceFromPriceWrap(el);
      if (isLikelyPrice(v)) return v;
    }

    const weightWraps = Array.from(card.querySelectorAll('div[class*="weightWrap"], span[class*="weightWrap"]'));
    for (i = 0; i < weightWraps.length; i++) {
      const el = weightWraps[i];
      if (el.closest('[class*="ratioWrapper"]')) continue;
      if (el.querySelector('div[class*="multiRow"]')) continue;
      const v = extractPriceFromNode(el);
      if (isLikelyPrice(v)) return v;
    }

    const vals = Array.from(
      card.querySelectorAll(
        'span[class*="cosratio"], div[class*="cosratio"], span[class*="comratio"], div[class*="comratio"]'
      )
    );
    for (i = 0; i < vals.length; i++) {
      const el = vals[i];
      if (el.closest('[class*="ratioWrapper"]')) continue;
      const raw = text(el);
      if (/%/.test(raw) || /万/.test(raw) || /团长/.test(raw)) continue;
      const v = extractPriceFromNode(el);
      if (isLikelyPrice(v)) return v;
    }

    return '';
  }

  function getCommissionRate(card) {
    const ratioWrap = card.querySelector('div[class*="ratioWrapper"]');
    if (!ratioWrap) return '';
    const raw = text(ratioWrap).replace(/\s+/g, '');
    const m = raw.match(/(\d+(?:\.\d+)?)%/);
    if (m) return m[1] + '%';
    const vals = Array.from(
      ratioWrap.querySelectorAll(
        'div[class*="cosratio"], div[class*="comratio"], span[class*="cosratio"], span[class*="comratio"]'
      )
    );
    for (let i = 0; i < vals.length; i++) {
      const t = text(vals[i]).replace(/\s+/g, '');
      if (/^\d+(?:\.\d+)?$/.test(t)) return t + '%';
      const mm = t.match(/(\d+(?:\.\d+)?)%/);
      if (mm) return mm[1] + '%';
    }
    return '';
  }

  function getSales(card) {
    const wraps = Array.from(card.querySelectorAll('div[class*="weightWrap"]'));
    for (let i = 0; i < wraps.length; i++) {
      const wrap = wraps[i];
      const multiRow = wrap.querySelector('div[class*="multiRow"]');
      if (!multiRow) continue;
      const valueEl = multiRow.querySelector('div[class*="cosratio"], div[class*="comratio"]');
      const v = text(valueEl);
      if (v && /^[\d.]+万?$/.test(v)) return v;
    }
    return '';
  }

  function getGoodRate(card) {
    const wrap = card.querySelector('div[class*="evaluationWrap"]');
    if (!wrap) return '';
    const valueEl = wrap.querySelector('div[class*="cosratio"], div[class*="comratio"]');
    const value = text(valueEl);
    return value ? (value.indexOf('%') >= 0 ? value : value + '%') : '';
  }

  function getExperienceScore(card) {
    const full = text(card);
    const m = full.match(/体验分\s*[:：]?\s*(\d+(?:\.\d+)?)/);
    return m ? m[1] : '';
  }

  function getAllPossibleUrlsFromNode(el) {
    if (!el) return [];
    const vals = [
      attr(el, 'href'),
      attr(el, 'data-url'),
      attr(el, 'data-href'),
      attr(el, 'data-link'),
      attr(el, 'jump-url'),
      attr(el, 'data-jump-url'),
      attr(el, 'data-router'),
      attr(el, 'data-target-url'),
    ].filter(Boolean);
    return vals.map(toAbs).filter(isHttp);
  }

  function getAllPossibleUrlsFromCard(card) {
    const urls = [];
    const nodes = [card].concat(Array.from(card.querySelectorAll('*')));
    for (let i = 0; i < nodes.length; i++) {
      urls.push.apply(urls, getAllPossibleUrlsFromNode(nodes[i]));
    }
    const aLinks = Array.from(card.querySelectorAll('a[href]')).map(function (a) {
      return toAbs(a.href || attr(a, 'href'));
    }).filter(isHttp);
    urls.push.apply(urls, aLinks);
    return Array.from(new Set(urls));
  }

  function getProductLink(card) {
    const urls = getAllPossibleUrlsFromCard(card);
    const productFirst = urls.find(function (u) {
      return /detail|ecommerce|trade|product|goods|promotion|jinritemai/i.test(u);
    });
    return productFirst ? normalizeUrl(productFirst) : '';
  }

  function extractProductIdFromUrl(url) {
    const u = s(url);
    if (!u) return '';
    try {
      const x = new URL(u, location.origin);
      const keys = ['id', 'product_id', 'item_id', 'goods_id'];
      for (let i = 0; i < keys.length; i++) {
        const v = s(x.searchParams.get(keys[i]) || '');
        if (isValidRealProductId(v)) return v;
      }
    } catch (_) {}
    const m = u.match(/(?:id|product_id|item_id|goods_id)[=/](\d{8,22})/i);
    return m && isValidRealProductId(m[1]) ? s(m[1]) : '';
  }

  function getShopLink(card) {
    const shopEl = card.querySelector('div[class*="shopTitle"]');
    const urls = getAllPossibleUrlsFromNode(shopEl).concat(getAllPossibleUrlsFromCard(card));
    const uniq = Array.from(new Set(urls));
    const shopFirst = uniq.find(function (u) {
      return /shop|store|merchant|author/i.test(u);
    });
    return shopFirst ? normalizeUrl(shopFirst) : '';
  }

  function getProductInfoBlock(card) {
    return (
      card.querySelector('div[class*="selectedPrdInfo"]') ||
      card.querySelector('div[class*="selectedPrdDesc"]') ||
      card
    );
  }

  function extractProductIdFromString(str) {
    if (!str) return '';
    const m = String(str).match(/商品ID\s*[:：]?\s*(\d{8,22})/i);
    if (m && isValidRealProductId(m[1])) return s(m[1]);
    return '';
  }

  function extractShopIdFromString(str) {
    if (!str) return '';
    const m =
      String(str).match(/店铺ID\s*[:：]?\s*(\d{6,22})/i) || String(str).match(/商家ID\s*[:：]?\s*(\d{6,22})/i);
    if (m && isValidRealShopId(m[1])) return s(m[1]);
    return '';
  }

  function extractShopIdFromUrl(href) {
    const u = s(href);
    if (!u) return '';
    try {
      const x = new URL(u, 'https://buyin.jinritemai.com');
      const v = s(x.searchParams.get('shop_id') || x.searchParams.get('shopId') || '');
      return isValidRealShopId(v) ? v : '';
    } catch (_) {
      const m = u.match(/[?&]shop_id=([^&]+)/i);
      if (!m) return '';
      const v = s(decodeURIComponent(m[1]));
      return isValidRealShopId(v) ? v : '';
    }
  }

  /** 只收集锚点附近的可见 tooltip 文案（按与锚点距离排序），禁止用全页第一个 tooltip */
  function getVisibleTooltipTextsNear(doc, anchorEl) {
    const d = doc && doc.body ? doc : document;
    const body = d.body;
    if (!body) return [];
    const all = body.querySelectorAll('*');
    const out = [];
    const ar = anchorEl && anchorEl.getBoundingClientRect ? anchorEl.getBoundingClientRect() : null;
    for (let ai = 0; ai < all.length; ai++) {
      const el = all[ai];
      const txt = text(el);
      if (!txt || !/(商品ID|店铺ID|商家ID)/i.test(txt)) continue;
      try {
        const win = el.ownerDocument.defaultView;
        if (!win) continue;
        const st = win.getComputedStyle(el);
        const r = el.getBoundingClientRect();
        if (st.display === 'none' || st.visibility === 'hidden' || r.width <= 0 || r.height <= 0) continue;
        const cls = String(el.className || '');
        const isTooltipLike =
          /tooltip|popover|trigger|overlay|semi|arco/i.test(cls) || attr(el, 'role') === 'tooltip';
        if (!isTooltipLike) continue;
        let dist = 0;
        if (ar) {
          const cx = r.left + r.width / 2;
          const cy = r.top + r.height / 2;
          const ax = ar.left + ar.width / 2;
          const ay = ar.top + ar.height / 2;
          const dx = cx - ax;
          const dy = cy - ay;
          dist = Math.sqrt(dx * dx + dy * dy);
        }
        out.push({ txt: txt, dist: dist });
      } catch (_) {}
    }
    out.sort(function (a, b) {
      return a.dist - b.dist;
    });
    const result = [];
    for (let i = 0; i < out.length; i++) {
      result.push(out[i].txt);
    }
    return result;
  }

  function getIdFromAttrs(card, type) {
    const nodes = [card].concat(Array.from(card.querySelectorAll('*')));
    for (let i = 0; i < nodes.length; i++) {
      const el = nodes[i];
      if (type === 'product') {
        const dp = attr(el, 'data-product-id');
        if (dp && isValidRealProductId(s(dp))) return s(dp);
      } else {
        const ds = attr(el, 'data-shop-id');
        if (ds && isValidRealShopId(s(ds))) return s(ds);
      }
      const vals = [
        attr(el, 'title'),
        attr(el, 'aria-label'),
        attr(el, 'data-title'),
        attr(el, 'data-tip'),
        attr(el, 'data-tooltip'),
        attr(el, 'data-clipboard-text'),
        attr(el, 'data-copy'),
        attr(el, 'data-product-id'),
        attr(el, 'data-shop-id'),
        attr(el, 'data-id'),
      ].filter(Boolean);
      for (let j = 0; j < vals.length; j++) {
        const v = vals[j];
        if (type === 'product') {
          const id = extractProductIdFromString(v);
          if (id) return id;
        } else {
          const id = extractShopIdFromString(v);
          if (id) return id;
        }
      }
    }
    return '';
  }

  function hoverIdMatchesCard(card, id, type) {
    if (!card || !id) return false;
    const t = text(card);
    if (t.indexOf(id) >= 0) return true;
    if (type === 'product') {
      const local = extractProductIdFromString(text(getProductInfoBlock(card)));
      if (local && local !== id) return false;
    } else {
      const local = extractShopIdFromString(t);
      if (local && local !== id) return false;
    }
    return true;
  }

  async function hoverReadId(card, type, doc) {
    let candidates = [];
    if (type === 'product') {
      candidates = Array.from(card.querySelectorAll('div,span,i')).filter(function (el) {
        return text(el) === 'ID';
      });
    } else {
      const shopBlock = (card.querySelector('div[class*="shopTitle"]') && card.querySelector('div[class*="shopTitle"]').parentElement) || card;
      candidates = Array.from(shopBlock.querySelectorAll('div,span,i')).filter(function (el) {
        return text(el) === 'ID';
      });
    }
    const rootDoc = doc || document;
    for (let i = 0; i < candidates.length; i++) {
      const node = candidates[i];
      const targets = [node, node.parentElement, node.nextElementSibling, node.previousElementSibling].filter(Boolean);
      for (let j = 0; j < targets.length; j++) {
        const target = targets[j];
        const beforeList = getVisibleTooltipTextsNear(rootDoc, target);
        const beforeSet = new Set(beforeList);
        try {
          target.dispatchEvent(new MouseEvent('mouseenter', { bubbles: true }));
          target.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
          target.dispatchEvent(new PointerEvent('pointerenter', { bubbles: true }));
          target.dispatchEvent(new PointerEvent('pointerover', { bubbles: true }));
        } catch (_) {}
        await geSleep(350);
        const afterList = getVisibleTooltipTextsNear(rootDoc, target);
        let listToUse = afterList.filter(function (t) {
          return !beforeSet.has(t);
        });
        if (!listToUse.length) listToUse = afterList;
        for (let k = 0; k < listToUse.length; k++) {
          const str = listToUse[k];
          if (type === 'product') {
            const id = extractProductIdFromString(str);
            if (id) return id;
          } else {
            const id = extractShopIdFromString(str);
            if (id) return id;
          }
        }
      }
    }
    return '';
  }

  async function getProductId(card, doc, productLink) {
    let id = getIdFromAttrs(card, 'product');
    if (isValidRealProductId(id)) return id;
    id = extractProductIdFromUrl(productLink || '');
    if (isValidRealProductId(id)) return id;
    const infoBlock = getProductInfoBlock(card);
    id = extractProductIdFromString(text(infoBlock));
    if (isValidRealProductId(id)) return id;
    const hoverId = await hoverReadId(card, 'product', doc);
    if (isValidRealProductId(hoverId) && hoverIdMatchesCard(card, hoverId, 'product')) return hoverId;
    return '';
  }

  function sanitizeCommissionRate(v) {
    const t = s(v);
    const m = t.match(/(\d+(?:\.\d+)?)%/);
    return m ? m[1] + '%' : '';
  }

  function sanitizeHandPrice(v) {
    const t = s(v);
    return isLikelyPrice(t) ? t : '';
  }

  async function getShopIdFromCard(card, doc, shopLink) {
    let id = extractShopIdFromUrl(shopLink || '');
    if (isValidRealShopId(id)) return id;
    id = getIdFromAttrs(card, 'shop');
    if (isValidRealShopId(id)) return id;
    id = extractShopIdFromString(text(card));
    if (isValidRealShopId(id)) return id;
    const hoverId = await hoverReadId(card, 'shop', doc);
    if (isValidRealShopId(hoverId) && hoverIdMatchesCard(card, hoverId, 'shop')) return hoverId;
    return '';
  }

  async function parseCardAsync(card, doc) {
    const 商品标题 = getTitle(card);
    const 店铺名称 = getShopName(card, 商品标题);
    const 商品图片链接 = getImage(card);
    let 商品链接 = getProductLink(card);
    let 店铺链接 = getShopLink(card);
    const 佣金率 = getCommissionRate(card);
    const 到手价 = getPrice(card);
    const 月销 = getSales(card);
    const 好评 = getGoodRate(card);
    const 体验分 = getExperienceScore(card);
    const 商品ID = await getProductId(card, doc, 商品链接);
    const 店铺ID = await getShopIdFromCard(card, doc, 店铺链接);

    if (!s(店铺链接)) {
      const sid = s(店铺ID);
      if (isValidRealShopId(sid)) {
        店铺链接 = normalizeUrl('https://buyin.jinritemai.com/dashboard/merch-picking-library/shop-detail?shop_id=' + encodeURIComponent(sid));
      }
    }

    if (!s(商品链接) && isValidRealProductId(商品ID)) {
      商品链接 = PRODUCT_LINK_TEMPLATE + encodeURIComponent(商品ID);
    }

    let outPid = s(商品ID);
    const outSid = s(店铺ID);
    if (outPid && outSid && outPid === outSid) {
      const fromUrl = extractProductIdFromUrl(商品链接 || '');
      if (isValidRealProductId(fromUrl) && fromUrl !== outSid) {
        outPid = fromUrl;
      } else {
        outPid = '';
      }
    }

    if (!s(商品链接) && isValidRealProductId(outPid)) {
      商品链接 = PRODUCT_LINK_TEMPLATE + encodeURIComponent(outPid);
    }

    const cleanCommissionRate = sanitizeCommissionRate(佣金率);
    const cleanHandPrice = sanitizeHandPrice(到手价);

    if (GE_CART_DEBUG && GE_FIELD_DEBUG_LOG_COUNT < 10) {
      GE_FIELD_DEBUG_LOG_COUNT++;
      console.log('[字段调试]', {
        title: 商品标题,
        productId: outPid,
        shopId: outSid,
        handPrice: cleanHandPrice,
        commissionRate: cleanCommissionRate,
        cardText: text(card).slice(0, 300),
      });
    }

    const img = s(商品图片链接);
    return {
      productInfo: s(商品标题),
      productId: outPid,
      productLink: s(商品链接),
      image: img,
      productImageLink: img,
      shop: s(店铺名称),
      shopId: outSid,
      shopLink: s(店铺链接),
      commissionRate: cleanCommissionRate,
      handPrice: cleanHandPrice,
      monthlySales: s(月销),
      goodRate: s(好评),
      experienceScore: s(体验分),
      source: 'dom_card',
    };
  }

  function getAllDocs() {
    const docs = [document];
    const ifs = Array.from(document.querySelectorAll('iframe'));
    for (let i = 0; i < ifs.length; i++) {
      try {
        const d = ifs[i].contentDocument;
        if (d && docs.indexOf(d) < 0) docs.push(d);
      } catch (_) {}
    }
    return docs;
  }

  function geDescribeScroller(el) {
    if (!el) return '(null)';
    try {
      const tag = el.tagName || '';
      const id = el.id ? '#' + el.id : '';
      const cls = el.className;
      const cs = typeof cls === 'string' ? cls : cls && cls.baseVal != null ? String(cls.baseVal) : '';
      const short = cs ? cs.replace(/\s+/g, ' ').trim().slice(0, 72) : '';
      return tag + id + (short ? '.' + short : '');
    } catch (_) {
      return '(?)';
    }
  }

  function geIsScrollableY(el, win) {
    if (!el || !win) return false;
    try {
      const st = win.getComputedStyle(el);
      const oy = st.overflowY;
      if (oy !== 'auto' && oy !== 'scroll' && oy !== 'overlay') return false;
      return el.scrollHeight > el.clientHeight + 2;
    } catch (_) {
      return false;
    }
  }

  const GE_HEADER_KEYWORDS = ['商品信息', '店铺', '佣金率', '到手价', '月销', '好评'];
  const GE_ACTION_TEXT_RE = /去带货|添加分组|联系商家|删除/;

  /** 不依赖 overflowY：按「选品车列表」特征给疑似滚动容器打分 */
  function geScrollCandidateScore(el, win) {
    if (!el || !win) return -1;
    try {
      const sh = el.scrollHeight;
      const ch = el.clientHeight;
      if (sh <= ch + 2) return -1;
      let score = 0;
      const cards = el.querySelectorAll('div[class*="selectedPrdInfo"]').length;
      score += cards * 14;
      const t = text(el);
      let hitHeader = 0;
      for (let hi = 0; hi < GE_HEADER_KEYWORDS.length; hi++) {
        if (t.indexOf(GE_HEADER_KEYWORDS[hi]) >= 0) hitHeader++;
      }
      score += hitHeader * 16;
      if (GE_ACTION_TEXT_RE.test(t)) score += 32;
      if (ch >= 200) score += Math.min(28, Math.floor(ch / 35));
      else if (ch >= 100) score += 10;
      try {
        const st = win.getComputedStyle(el);
        const oy = st.overflowY;
        if (oy === 'auto' || oy === 'scroll' || oy === 'overlay') score += 20;
      } catch (_) {}
      const tag = el.tagName;
      if (tag === 'BODY' || tag === 'HTML') score -= 45;
      return score;
    } catch (_) {
      return -1;
    }
  }

  function geCollectScrollCandidates(doc) {
    if (!doc || !doc.body) return [];
    const win = doc.defaultView;
    if (!win) return [];
    const infos = doc.querySelectorAll('div[class*="selectedPrdInfo"]');
    const seen = new Map();
    for (let i = 0; i < infos.length; i++) {
      let p = infos[i].parentElement;
      for (let depth = 0; depth < 56 && p && p !== doc.documentElement; depth++) {
        const sc = geScrollCandidateScore(p, win);
        if (sc > 0) {
          const prev = seen.get(p) || 0;
          if (sc > prev) seen.set(p, sc);
        }
        p = p.parentElement;
      }
    }
    const arr = Array.from(seen.entries()).map(function (kv) {
      return { el: kv[0], score: kv[1] };
    });
    arr.sort(function (a, b) {
      return b.score - a.score;
    });
    return arr;
  }

  function geLogScrollCandidates(doc, candidates) {
    let href = '(inline)';
    try {
      if (doc && doc.location) href = doc.location.href;
    } catch (_) {
      href = '(cross-origin)';
    }
    if (!candidates || !candidates.length) {
      console.log('[滚动候选] (无) doc=' + href);
      return;
    }
    console.log('[选品车抓取] 滚动候选共 ' + candidates.length + ' 个 doc=' + href);
    for (let i = 0; i < candidates.length; i++) {
      const c = candidates[i];
      const el = c.el;
      const t = text(el);
      let hitHeader = 0;
      for (let hi = 0; hi < GE_HEADER_KEYWORDS.length; hi++) {
        if (t.indexOf(GE_HEADER_KEYWORDS[hi]) >= 0) hitHeader++;
      }
      console.log('[滚动候选]', {
        rank: i + 1,
        score: c.score,
        tag: el.tagName,
        className: String(el.className || '').slice(0, 120),
        clientHeight: el.clientHeight,
        scrollHeight: el.scrollHeight,
        cardCount: el.querySelectorAll('div[class*="selectedPrdInfo"]').length,
        hitHeaderKeywords: hitHeader,
        text: t.slice(0, 80),
      });
    }
  }

  function gePickBestScrollCandidate(candidates) {
    if (!candidates || !candidates.length) return null;
    return candidates[0].el;
  }

  /** 从指定卡片节点向上试滚，锁定 scrollTop 能变化或卡片数能增长的层 */
  async function probeScrollableAncestorFromAnchor(doc, rootIndex) {
    if (!doc || !doc.body) return null;
    const roots = doc.querySelectorAll('div[class*="selectedPrdInfo"]');
    if (!roots.length) return null;
    const idx = Math.max(0, Math.min(rootIndex, roots.length - 1));
    const anchor = roots[idx];
    let p = anchor.parentElement;
    while (p && p !== doc.documentElement) {
      const max = Math.max(0, p.scrollHeight - p.clientHeight);
      if (max < 2) {
        p = p.parentElement;
        continue;
      }
      const before = p.scrollTop;
      const beforeCards = doc.querySelectorAll('div[class*="selectedPrdInfo"]').length;
      const step = Math.max(120, Math.min(400, Math.floor(p.clientHeight * 0.5)));
      const next = Math.min(before + step, max);
      try {
        p.scrollTop = next;
        try {
          p.dispatchEvent(new Event('scroll', { bubbles: true }));
        } catch (_) {}
        try {
          p.dispatchEvent(
            new WheelEvent('wheel', {
              bubbles: true,
              cancelable: true,
              deltaY: step,
            })
          );
        } catch (_) {}
      } catch (_) {}
      await geSleep(380);
      const afterCards = doc.querySelectorAll('div[class*="selectedPrdInfo"]').length;
      const moved = Math.abs(p.scrollTop - before) > 0.5;
      if (afterCards > beforeCards || moved) {
        try {
          p.scrollTop = before;
        } catch (_) {}
        console.log(
          '[选品车抓取] 试滚兜底命中 anchor#' +
            idx +
            ' ' +
            geDescribeScroller(p) +
            ' 试滚前卡片=' +
            beforeCards +
            ' 试滚后=' +
            afterCards +
            ' 位移=' +
            moved
        );
        return p;
      }
      try {
        p.scrollTop = before;
      } catch (_) {}
      p = p.parentElement;
    }
    return null;
  }

  async function probeScrollableAncestor(doc) {
    let r = await probeScrollableAncestorFromAnchor(doc, 0);
    if (r) return r;
    const roots = doc.querySelectorAll('div[class*="selectedPrdInfo"]');
    if (roots.length > 1) {
      r = await probeScrollableAncestorFromAnchor(doc, roots.length - 1);
    }
    return r;
  }

  /** 向下滚动一屏（触发虚拟列表与懒加载），返回是否发生位移或已尝试滚到底 */
  function scrollListStep(scroller) {
    if (!scroller) return false;
    const beforeTop = scroller.scrollTop;
    const max = Math.max(0, scroller.scrollHeight - scroller.clientHeight);
    const step = Math.max(120, Math.floor(scroller.clientHeight * 0.9));
    const next = Math.min(beforeTop + step, max);
    scroller.scrollTop = next;
    try {
      scroller.dispatchEvent(new Event('scroll', { bubbles: true }));
    } catch (_) {}
    try {
      scroller.dispatchEvent(
        new WheelEvent('wheel', {
          bubbles: true,
          cancelable: true,
          deltaY: step,
        })
      );
    } catch (_) {}
    const afterTop = scroller.scrollTop;
    return Math.abs(afterTop - beforeTop) > 0.5 || next > beforeTop;
  }

  function geScrollMainWindowStep() {
    const se = document.scrollingElement || document.documentElement;
    const before = se.scrollTop;
    const max = Math.max(0, se.scrollHeight - se.clientHeight);
    const step = Math.max(100, Math.floor(se.clientHeight * 0.85));
    const next = Math.min(before + step, max);
    se.scrollTop = next;
    try {
      se.dispatchEvent(new Event('scroll', { bubbles: true }));
    } catch (_) {}
    try {
      se.dispatchEvent(
        new WheelEvent('wheel', {
          bubbles: true,
          cancelable: true,
          deltaY: step,
        })
      );
    } catch (_) {}
    return Math.abs(se.scrollTop - before) > 0.5 || next > before;
  }

  /** 将列表容器瞬间滚到当前 scrollHeight 底部并派发事件（配合步进触发懒加载） */
  function scrollListSnapToBottom(scroller) {
    if (!scroller) return false;
    const max = Math.max(0, scroller.scrollHeight - scroller.clientHeight);
    const before = scroller.scrollTop;
    scroller.scrollTop = max;
    try {
      scroller.dispatchEvent(new Event('scroll', { bubbles: true }));
    } catch (_) {}
    try {
      const dh = Math.max(400, Math.floor(scroller.clientHeight * 0.95));
      scroller.dispatchEvent(
        new WheelEvent('wheel', {
          bubbles: true,
          cancelable: true,
          deltaY: dh,
        })
      );
    } catch (_) {}
    return Math.abs(scroller.scrollTop - before) > 0.5 || max > before + 0.5;
  }

  /** 将主文档滚动条拉到底（与列表内滚动并行，双保险） */
  function geScrollWindowSnapToBottom() {
    const se = document.scrollingElement || document.documentElement;
    const max = Math.max(0, se.scrollHeight - se.clientHeight);
    const before = se.scrollTop;
    se.scrollTop = max;
    try {
      se.dispatchEvent(new Event('scroll', { bubbles: true }));
    } catch (_) {}
    try {
      const dh = Math.max(400, Math.floor(se.clientHeight * 0.95));
      se.dispatchEvent(
        new WheelEvent('wheel', {
          bubbles: true,
          cancelable: true,
          deltaY: dh,
        })
      );
    } catch (_) {}
    return Math.abs(se.scrollTop - before) > 0.5;
  }

  function geCountSelectedPrdInfoAllDocs() {
    let n = 0;
    const docs = getAllDocs();
    for (let i = 0; i < docs.length; i++) {
      try {
        n += docs[i].querySelectorAll('div[class*="selectedPrdInfo"]').length;
      } catch (_) {}
    }
    return n;
  }

  /** 列表底部「没有更多(了)」提示：主文档与 iframe 全文检测 */
  function geDetectNoMoreFooterInDocs() {
    const re = /没有更多[了]?/;
    return getAllDocs().some(function (dd) {
      try {
        const root = dd.body || dd.documentElement;
        if (!root) return false;
        return re.test(text(root));
      } catch (_) {
        return false;
      }
    });
  }

  /** 已在列表底部时仍派发 wheel，促发「加载更多」（虚拟列表常见） */
  function geNudgeWheelWhenListAtBottom(scroller) {
    if (!scroller) return false;
    let max = 0;
    try {
      max = Math.max(0, scroller.scrollHeight - scroller.clientHeight);
    } catch (_) {
      return false;
    }
    if (scroller.scrollTop < max - 1.5) return false;
    for (let i = 0; i < 8; i++) {
      try {
        scroller.dispatchEvent(
          new WheelEvent('wheel', {
            bubbles: true,
            cancelable: true,
            deltaY: 100,
          })
        );
      } catch (_) {}
    }
    try {
      scroller.dispatchEvent(new Event('scroll', { bubbles: true }));
    } catch (_) {}
    return true;
  }

  /** 单段同步追底（步进 + 吸底 + 高度变化） */
  function geListScrollerBurstSync(el) {
    let moved = false;
    let beforeH = 0;
    try {
      beforeH = el.scrollHeight;
    } catch (_) {}
    for (let k = 0; k < 12; k++) {
      if (scrollListStep(el)) moved = true;
    }
    for (let z = 0; z < 2; z++) {
      if (scrollListSnapToBottom(el)) moved = true;
    }
    try {
      if (el.scrollHeight > beforeH) moved = true;
    } catch (_) {}
    try {
      if (geNudgeWheelWhenListAtBottom(el)) moved = true;
    } catch (_) {}
    return moved;
  }

  /** 多段微间隔追底，给虚拟列表留出布局/请求时间，避免首轮加载后「假死」 */
  async function geChaseListScrollerMicro(el) {
    let any = false;
    for (let burst = 0; burst < 5; burst++) {
      if (geListScrollerBurstSync(el)) any = true;
      await geSleep(95);
    }
    return any;
  }

  async function geChaseWindowMicro() {
    const se = document.scrollingElement || document.documentElement;
    let moved = false;
    let beforeDocH = 0;
    try {
      beforeDocH = se ? se.scrollHeight : 0;
    } catch (_) {}
    for (let burst = 0; burst < 4; burst++) {
      for (let k = 0; k < 10; k++) {
        if (geScrollMainWindowStep()) moved = true;
      }
      for (let z = 0; z < 2; z++) {
        if (geScrollWindowSnapToBottom()) moved = true;
      }
      await geSleep(95);
    }
    try {
      if (se && se.scrollHeight > beforeDocH) moved = true;
    } catch (_) {}
    return moved;
  }

  /**
   * 阶段①：持续滚动直到出现「没有更多」。
   * 若已识别列表内滚动容器：只滚该容器，不滚 document 窗口（避免双滚动条、白屏挡 UI）。
   * 若无列表容器：才滚主窗口。
   */
  async function geCartScrollPhaseLoadAll(scrollers, onProgress) {
    const maxRounds = 600;
    let endFlagStreak = 0;
    let noMoveStreak = 0;
    console.log(
      '[选品车抓取] ①预加载：持续滚动直至「没有更多」' +
        (scrollers.length ? '（仅列表内滚动，不滚整页）' : '（仅主窗口滚动）') +
        '…'
    );
    for (let round = 0; round < maxRounds; round++) {
      if (scrollers.length) {
        for (let si = 0; si < scrollers.length; si++) {
          try {
            const again = await findScrollableContainerAsync(scrollers[si].doc);
            if (again) scrollers[si].el = again;
          } catch (_) {}
        }
      }

      const beforeCards = geCountSelectedPrdInfoAllDocs();
      let moved = false;

      if (scrollers.length) {
        for (let si = 0; si < scrollers.length; si++) {
          let el = scrollers[si].el;
          try {
            if (!el.isConnected) {
              const again = await findScrollableContainerAsync(scrollers[si].doc);
              if (again) scrollers[si].el = again;
              el = scrollers[si].el;
            }
          } catch (_) {}
          try {
            if (await geChaseListScrollerMicro(el)) moved = true;
          } catch (_) {}
        }
      } else {
        try {
          if (await geChaseWindowMicro()) moved = true;
        } catch (_) {}
      }

      await geSleep(480);

      const totalCards = geCountSelectedPrdInfoAllDocs();
      if (totalCards > beforeCards) moved = true;

      const endFlag = geDetectNoMoreFooterInDocs();
      if (endFlag) endFlagStreak++;
      else endFlagStreak = 0;

      if (moved) noMoveStreak = 0;
      else noMoveStreak++;

      console.log(
        '[选品车抓取] ①预加载 第' +
          (round + 1) +
          '/' +
          maxRounds +
          ' 轮 DOM卡片≈' +
          totalCards +
          ' 本轮卡片Δ' +
          (totalCards - beforeCards) +
          ' 位移/进展=' +
          moved +
          ' 底部提示=' +
          endFlag +
          ' 底部提示连续=' +
          endFlagStreak +
          ' 无进展连续=' +
          noMoveStreak
      );

      if (typeof onProgress === 'function') {
        try {
          onProgress({
            round: round + 1,
            totalCards: totalCards,
            endFlag: endFlag,
            moved: moved,
            phase: 'preload',
          });
        } catch (_) {}
      }

      if (endFlag && endFlagStreak >= 2) break;
      if (noMoveStreak >= 200 && !endFlag) {
        console.warn('[选品车抓取] ①预加载：长时间无位移/增高且未检测到「没有更多」，停止（请检查列表容器）');
        break;
      }
    }
    console.log('[选品车抓取] ①预加载结束 底部提示=' + geDetectNoMoreFooterInDocs() + ' DOM卡片≈' + geCountSelectedPrdInfoAllDocs());
  }

  /** 对候选容器试滚一步，确认 scrollTop 真能变化（筛掉误中的外层包裹） */
  async function geVerifyScrollerMoves(doc, el) {
    if (!el) return false;
    const max = Math.max(0, el.scrollHeight - el.clientHeight);
    if (max < 2) return true;
    const before = el.scrollTop;
    if (before >= max - 0.5) return true;
    scrollListStep(el);
    await geSleep(140);
    const moved = Math.abs(el.scrollTop - before) > 0.5;
    try {
      el.scrollTop = before;
    } catch (_) {}
    return moved;
  }

  /** 在文档内找出列表主滚动容器：打分候选 → 试滚兜底（不再依赖仅 overflow 的投票） */
  async function findScrollableContainerAsync(doc) {
    if (!doc || !doc.documentElement) return null;
    const win = doc.defaultView;
    if (!win) return null;

    const candidates = geCollectScrollCandidates(doc);
    geLogScrollCandidates(doc, candidates);
    let best = gePickBestScrollCandidate(candidates);
    if (best) {
      console.log('[选品车抓取] 选用打分最高滚动容器:', geDescribeScroller(best));
      const okMove = await geVerifyScrollerMoves(doc, best);
      if (okMove) return best;
      console.warn('[选品车抓取] 打分最高容器试滚无位移，改用试滚兜底');
      const probed = await probeScrollableAncestor(doc);
      if (probed) return probed;
      return best;
    }

    const probed = await probeScrollableAncestor(doc);
    if (probed) return probed;

    let bestEl = null;
    let bestSpare = 0;
    let walked = 0;
    const maxWalk = 8000;
    function walk(el) {
      if (!el || walked >= maxWalk) return;
      walked++;
      if (geIsScrollableY(el, win)) {
        const spare = el.scrollHeight - el.clientHeight;
        if (spare > bestSpare) {
          bestSpare = spare;
          bestEl = el;
        }
      }
      for (let c = el.firstElementChild; c; c = c.nextElementSibling) walk(c);
    }
    if (doc.body) walk(doc.body);
    if (bestEl) console.log('[选品车抓取] 回退：overflow 可滚最大区域', geDescribeScroller(bestEl));
    return bestEl;
  }

  async function geScanAllByVisibleCards() {
    GE_CART_SCAN_ACTIVE = true;
    GE_FIELD_DEBUG_LOG_COUNT = 0;
    try {
    const docs = getAllDocs();
    const scrollers = [];
    for (let di = 0; di < docs.length; di++) {
      const el = await findScrollableContainerAsync(docs[di]);
      if (el) scrollers.push({ doc: docs[di], el: el });
    }
    const scrollerDesc = scrollers.length
      ? scrollers
          .map(function (s, i) {
            return '[doc' + i + '] ' + geDescribeScroller(s.el) + ' scrollH=' + s.el.scrollHeight + ' clientH=' + s.el.clientHeight;
          })
          .join(' | ')
      : '(未识别到内部滚动容器，将回退 window 滚动；若条数不增请看 [滚动候选] 与试滚日志)';
    console.log('[选品车抓取] 最终列表滚动容器:', scrollerDesc);

    let totalInc = 0;
    const baseCount = STORE.rows.length;

    if (!scrollers.length) {
      console.warn(
        '[选品车抓取] 未挂载任何列表内滚动容器，预加载阶段仅滚窗口；若条数偏少，请检查上方 [滚动候选] 是否为空或分数过低'
      );
    }

    function geReportPreloadProgress(info) {
      const line = document.getElementById('ge-cart-progress');
      if (!line) return;
      const tail = info.endFlag ? ' · 已检测到「没有更多」' : '';
      line.textContent =
        '①预加载 第 ' +
        info.round +
        ' 轮 · 页面约 ' +
        info.totalCards +
        ' 个商品块 · 滚动' +
        (info.moved ? '有位移' : '暂无明显位移') +
        tail;
      line.style.display = 'block';
    }

    const line0 = document.getElementById('ge-cart-progress');
    if (line0) {
      line0.textContent =
        '①预加载：持续滚动直到「没有更多」' +
        (scrollers.length ? '（只滚商品列表区域，不滚整页）' : '（滚主窗口）') +
        '…';
      line0.style.display = 'block';
    }

    await geCartScrollPhaseLoadAll(scrollers, geReportPreloadProgress);

    const line2 = document.getElementById('ge-cart-progress');
    if (line2) {
      line2.textContent = '②解析：正在读取商品卡片并写入脚本缓存…';
      line2.style.display = 'block';
    }

    console.log('[选品车抓取] ②解析：预加载结束后统一解析 DOM 并入库（接口拦截数据已并行合并）…');
    const docsNow = getAllDocs();
    const out = [];
    let cardCountTotal = 0;
    for (let d = 0; d < docsNow.length; d++) {
      const dd = docsNow[d];
      const cards = getCardsInDoc(dd);
      cardCountTotal += cards.length;
      for (let i = 0; i < cards.length; i++) {
        const row = await parseCardAsync(cards[i], dd);
        if (!row.productInfo && !row.image && !row.productImageLink) continue;
        out.push(row);
      }
    }

    totalInc = addRows(out);
    console.log(
      '[选品车抓取] ②解析完成 命中卡片=' + cardCountTotal + ' 待解析行=' + out.length + ' DOM入库新增=' + totalInc + ' 当前累计=' + STORE.rows.length
    );

    const deltaAll = STORE.rows.length - baseCount;
    console.log('[选品车抓取] 扫描结束 本轮任务累计新增=' + deltaAll + ' 扫描前基数=' + baseCount + ' 当前总行数=' + STORE.rows.length);

    return deltaAll;
    } finally {
      GE_CART_SCAN_ACTIVE = false;
    }
  }

  function parseCsvValue(v) {
    const t = s(v).replace(/\r?\n/g, ' ');
    if (/[,"\n]/.test(t)) return '"' + t.replace(/"/g, '""') + '"';
    return t;
  }

  function toCsv(rows) {
    const cfg = loadFeishuConfig();
    const enabled = parseFieldEnabledJson(cfg.fieldEnabledJson);
    const meta = EXPORT_FIELD_META.filter(function (item) {
      if (CSV_ALWAYS_EXPORT_KEYS.indexOf(item.key) >= 0) return true;
      return enabled[item.key] !== false;
    });
    const head = meta.map(function (m) {
      return m.title;
    });
    const lines = [head.join(',')];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      fillProductLinkFromId(r);
      const cells = [];
      for (let j = 0; j < meta.length; j++) {
        cells.push(parseCsvValue(geCsvRowValueForKey(r, meta[j].key)));
      }
      lines.push(cells.join(','));
    }
    return '\uFEFF' + lines.join('\n');
  }

  function downloadCsv() {
    if (!STORE.rows.length) {
      alert('暂无可导出数据，请先点侧栏「一键抓取（滚到底）」抓取选品车列表。');
      return;
    }
    const copy = STORE.rows.map(function (r) {
      return { ...r };
    });
    copy.forEach(fillProductLinkFromId);
    const rows = dedupeFinalRows(copy);
    const blob = new Blob([toCsv(rows)], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = '百应选品车_' + Date.now() + '.csv';
    a.click();
    setTimeout(function () {
      URL.revokeObjectURL(a.href);
    }, 2000);
  }

  // ---------- 飞书多维表格（与「百应+淘宝 商品一键导出」配置方式相同，独立存储键）----------
  const FEISHU_CFG_KEY = 'ge_buyin_cart_feishu_v1';
  const GE_CLEAR_AFTER_FEISHU_KEY = 'ge_buyin_cart_clear_after_feishu_v1';
  /** 与「写入后清空」同时勾选时，飞书写入成功后不再弹确认，直接尝试 DOM 批量删除 */
  const GE_CLEAR_AFTER_FEISHU_AUTO_KEY = 'ge_buyin_cart_clear_after_feishu_auto_v1';
  /** 一键抓取完成后自动调用写入飞书（需已配置飞书） */
  const GE_AUTO_FEISHU_AFTER_SCAN_KEY = 'ge_buyin_cart_auto_feishu_after_scan_v1';
  /** 一键抓取完成后逐条删除失效商品（商品ID与店铺ID均为空） */
  const GE_DELETE_INVALID_AFTER_SCAN_KEY = 'ge_buyin_cart_delete_invalid_after_scan_v1';
  const FEISHU_DEFAULT_FIELD_MAP = {
    productInfo: '商品信息',
    productId: '商品ID',
    handPrice: '到手价',
    commissionRate: '佣金率',
    productLink: '商品链接',
    shop: '店铺',
    shopId: '店铺ID',
    experienceScore: '体验分',
    goodRate: '好评',
    monthlySales: '月销',
    productImageLink: '商品图片链接',
  };
  const FEISHU_UPLOAD_FIELD_KEYS = [
    'productInfo',
    'productId',
    'handPrice',
    'commissionRate',
    'productLink',
    'shop',
    'shopId',
    'experienceScore',
    'goodRate',
    'monthlySales',
    'productImageLink',
  ];
  /** 配置弹窗左侧说明（右侧为用户可改的飞书列标题） */
  const FEISHU_FIELD_UI_LABELS = {
    productInfo: '商品信息',
    productId: '商品ID',
    handPrice: '到手价',
    commissionRate: '佣金率',
    productLink: '商品链接',
    shop: '店铺',
    shopId: '店铺ID',
    experienceScore: '体验分',
    goodRate: '好评',
    monthlySales: '月销',
    productImageLink: '商品图片链接',
  };

  const FEISHU_DEFAULT_FIELD_ENABLED = {
    productInfo: true,
    productId: true,
    handPrice: true,
    commissionRate: true,
    productLink: true,
    shop: true,
    shopId: true,
    experienceScore: true,
    goodRate: true,
    monthlySales: true,
    productImageLink: true,
  };

  /** CSV：店铺链接、来源不在映射勾选内，始终导出 */
  const CSV_ALWAYS_EXPORT_KEYS = ['shopLink', 'source'];

  const EXPORT_FIELD_META = [
    { key: 'productInfo', title: '商品信息' },
    { key: 'productId', title: '商品ID' },
    { key: 'handPrice', title: '到手价' },
    { key: 'commissionRate', title: '佣金率' },
    { key: 'productLink', title: '商品链接' },
    { key: 'shop', title: '店铺' },
    { key: 'shopId', title: '店铺ID' },
    { key: 'experienceScore', title: '体验分' },
    { key: 'goodRate', title: '好评' },
    { key: 'monthlySales', title: '月销' },
    { key: 'productImageLink', title: '商品图片链接' },
    { key: 'shopLink', title: '店铺链接' },
    { key: 'source', title: '来源' },
  ];

  function parseFieldEnabledJson(str) {
    const def = { ...FEISHU_DEFAULT_FIELD_ENABLED };
    if (!str || !String(str).trim()) return def;
    try {
      const o = JSON.parse(str);
      if (!o || typeof o !== 'object') return def;
      const out = { ...def };
      for (const k of Object.keys(def)) {
        if (typeof o[k] === 'boolean') out[k] = o[k];
      }
      return out;
    } catch (_) {
      return def;
    }
  }

  function getEnabledUploadFieldKeys(cfg) {
    const enabled = parseFieldEnabledJson(cfg.fieldEnabledJson);
    return FEISHU_UPLOAD_FIELD_KEYS.filter(function (k) {
      return enabled[k] !== false;
    });
  }

  function geCsvRowValueForKey(r, key) {
    if (!r) return '';
    if (key === 'productImageLink') return s(r.productImageLink || r.image);
    return s(r[key]);
  }

  function geGm() {
    return typeof GM !== 'undefined' ? GM : null;
  }
  function gmGet(key, def) {
    try {
      if (typeof GM_getValue === 'function') return GM_getValue(key, def);
      const g = geGm();
      if (g && typeof g.getValue === 'function') return g.getValue(key, def);
    } catch (_) {}
    try {
      const x = localStorage.getItem(key);
      return x != null ? x : def;
    } catch (_) {
      return def;
    }
  }
  function gmSet(key, val) {
    try {
      if (typeof GM_setValue === 'function') return void GM_setValue(key, val);
      const g = geGm();
      if (g && typeof g.setValue === 'function') return void g.setValue(key, val);
    } catch (_) {}
    try {
      localStorage.setItem(key, val);
    } catch (_) {}
  }
  function gmXhr(opts) {
    return new Promise(function (resolve, reject) {
      const done = function (r) {
        try {
          resolve({ status: r.status, responseText: r.responseText || '' });
        } catch (e) {
          reject(e);
        }
      };
      const req = {
        method: opts.method || 'GET',
        url: opts.url,
        headers: opts.headers || {},
        data: opts.body,
        onload: done,
        onerror: function () {
          reject(new Error('网络错误'));
        },
        ontimeout: function () {
          reject(new Error('请求超时'));
        },
      };
      if (typeof GM_xmlhttpRequest === 'function') {
        GM_xmlhttpRequest(req);
        return;
      }
      const g = geGm();
      if (g && typeof g.xmlHttpRequest === 'function') {
        g.xmlHttpRequest(req);
        return;
      }
      reject(new Error('需要 Tampermonkey / Violentmonkey 并授权访问 open.feishu.cn'));
    });
  }

  function loadFeishuConfig() {
    const raw = gmGet(FEISHU_CFG_KEY, '');
    if (!raw) {
      return {
        wikiNodeToken: '',
        appToken: '',
        tableId: '',
        accessToken: '',
        feishuAppId: '',
        feishuAppSecret: '',
        fieldMapJson: '',
        fieldEnabledJson: '',
        useHyperlink: false,
        feishuCoerceNumberFields: false,
        feishuAutoCreateFields: true,
      };
    }
    try {
      const j = JSON.parse(raw);
      return {
        wikiNodeToken: (j.wikiNodeToken || '').trim(),
        appToken: (j.appToken || '').trim(),
        tableId: (j.tableId || '').trim(),
        accessToken: (j.accessToken || '').trim(),
        feishuAppId: (j.feishuAppId || '').trim(),
        feishuAppSecret: (j.feishuAppSecret || '').trim(),
        fieldMapJson: typeof j.fieldMapJson === 'string' ? j.fieldMapJson : '',
        fieldEnabledJson: typeof j.fieldEnabledJson === 'string' ? j.fieldEnabledJson : '',
        useHyperlink: !!j.useHyperlink,
        feishuCoerceNumberFields: !!j.feishuCoerceNumberFields,
        feishuAutoCreateFields: j.feishuAutoCreateFields !== false,
      };
    } catch (_) {
      return {
        wikiNodeToken: '',
        appToken: '',
        tableId: '',
        accessToken: '',
        feishuAppId: '',
        feishuAppSecret: '',
        fieldMapJson: '',
        fieldEnabledJson: '',
        useHyperlink: false,
        feishuCoerceNumberFields: false,
        feishuAutoCreateFields: true,
      };
    }
  }

  function saveFeishuConfig(cfg) {
    gmSet(
      FEISHU_CFG_KEY,
      JSON.stringify({
        wikiNodeToken: (cfg.wikiNodeToken || '').trim(),
        appToken: (cfg.appToken || '').trim(),
        tableId: (cfg.tableId || '').trim(),
        accessToken: (cfg.accessToken || '').trim(),
        feishuAppId: (cfg.feishuAppId || '').trim(),
        feishuAppSecret: (cfg.feishuAppSecret || '').trim(),
        fieldMapJson: typeof cfg.fieldMapJson === 'string' ? cfg.fieldMapJson : '',
        fieldEnabledJson: typeof cfg.fieldEnabledJson === 'string' ? cfg.fieldEnabledJson : '',
        useHyperlink: !!cfg.useHyperlink,
        feishuCoerceNumberFields: !!cfg.feishuCoerceNumberFields,
        feishuAutoCreateFields: cfg.feishuAutoCreateFields !== false,
      })
    );
  }

  function parseFieldMapJson(str) {
    const def = { ...FEISHU_DEFAULT_FIELD_MAP };
    if (!str || !String(str).trim()) return def;
    try {
      const o = JSON.parse(str);
      if (!o || typeof o !== 'object') return def;
      const out = { ...def };
      for (const k of Object.keys(FEISHU_DEFAULT_FIELD_MAP)) {
        if (typeof o[k] === 'string' && o[k].trim()) out[k] = o[k].trim();
      }
      return out;
    } catch (_) {
      return def;
    }
  }

  function mergeFieldMap(cfg) {
    return parseFieldMapJson(cfg.fieldMapJson);
  }

  function chunk(arr, n) {
    const o = [];
    for (let i = 0; i < arr.length; i += n) o.push(arr.slice(i, i + n));
    return o;
  }

  function looksLikeFeishuAccessTokenNotWikiNode(tok) {
    const v = (tok || '').trim();
    if (!v) return false;
    if (/^t-g[a-z0-9]{8,}$/i.test(v)) return true;
    if (/^u-[a-z0-9_-]{10,}$/i.test(v)) return true;
    if (/^pat_[a-z0-9]{8,}$/i.test(v)) return true;
    return false;
  }

  async function wikiNodeToBitableAppToken(wikiNodeToken, accessToken) {
    const nodeTok = wikiNodeToken.trim();
    if (looksLikeFeishuAccessTokenNotWikiNode(nodeTok)) {
      throw new Error(
        '「Wiki 节点 token」误填成了鉴权串（如 t-g…）。请填浏览器地址栏 /wiki/ 与 ? 之间的节点 ID。'
      );
    }
    const tok = accessToken.trim();
    if (!tok) throw new Error('access_token 为空');
    const auth = tok.toLowerCase().startsWith('bearer ') ? tok : 'Bearer ' + tok;
    const url =
      'https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token=' + encodeURIComponent(nodeTok);
    const r = await gmXhr({ method: 'GET', url, headers: { Authorization: auth } });
    let json;
    try {
      json = JSON.parse(r.responseText || '{}');
    } catch (_) {
      throw new Error('Wiki 接口返回非 JSON（HTTP ' + r.status + '）');
    }
    if (json.code !== 0) throw new Error((json.msg || '获取 Wiki 节点失败') + ' code=' + json.code);
    const node = json.data && json.data.node;
    if (!node || node.obj_type !== 'bitable' || !node.obj_token) {
      throw new Error('Wiki 节点不是多维表格或未返回 obj_token');
    }
    return String(node.obj_token);
  }

  async function resolveBitableAppToken(cfg) {
    const wiki = (cfg.wikiNodeToken || '').trim();
    if (wiki) return wikiNodeToBitableAppToken(wiki, cfg.accessToken);
    return (cfg.appToken || '').trim();
  }

  async function feishuFetchTenantAccessToken(appId, appSecret) {
    const r = await gmXhr({
      method: 'POST',
      url: 'https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal',
      headers: { 'Content-Type': 'application/json; charset=utf-8' },
      body: JSON.stringify({ app_id: appId.trim(), app_secret: appSecret.trim() }),
    });
    let json;
    try {
      json = JSON.parse(r.responseText || '{}');
    } catch (_) {
      throw new Error('换取 token 返回非 JSON');
    }
    if (json.code !== 0) throw new Error((json.msg || '换取 tenant_access_token 失败') + ' code=' + json.code);
    if (!json.tenant_access_token) throw new Error('响应中无 tenant_access_token');
    return String(json.tenant_access_token);
  }

  async function ensureFeishuAccessToken(cfg) {
    const id = (cfg.feishuAppId || '').trim();
    const sec = (cfg.feishuAppSecret || '').trim();
    if (id && sec) {
      const t = await feishuFetchTenantAccessToken(id, sec);
      return { ...cfg, accessToken: t };
    }
    return cfg;
  }

  function feishuBitableTypeForMappedKey(internalKey, cfg) {
    if (
      cfg.feishuCoerceNumberFields &&
      (internalKey === 'handPrice' || internalKey === 'shopId' || internalKey === 'monthlySales')
    )
      return 2;
    if (cfg.useHyperlink && (internalKey === 'productLink' || internalKey === 'productImageLink')) return 15;
    return 1;
  }

  async function feishuBitableListFieldsPage(appToken, tableId, accessToken, pageToken) {
    let url =
      'https://open.feishu.cn/open-apis/bitable/v1/apps/' +
      encodeURIComponent(appToken) +
      '/tables/' +
      encodeURIComponent(tableId) +
      '/fields?page_size=100';
    if (pageToken) url += '&page_token=' + encodeURIComponent(pageToken);
    const tok = accessToken.trim();
    const auth = tok.toLowerCase().startsWith('bearer ') ? tok : 'Bearer ' + tok;
    const r = await gmXhr({ method: 'GET', url, headers: { Authorization: auth } });
    let json;
    try {
      json = JSON.parse(r.responseText || '{}');
    } catch (_) {
      throw new Error('列出字段失败：非 JSON');
    }
    if (json.code !== 0) throw new Error((json.msg || '列出字段失败') + ' code=' + json.code);
    const d = json.data || {};
    return {
      items: Array.isArray(d.items) ? d.items : [],
      hasMore: !!d.has_more,
      pageToken: d.page_token || '',
    };
  }

  function normFieldName(t) {
    return String(t || '')
      .replace(/[\u00a0\u200b\ufeff]/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  /** 拉取子表全部字段：精确名集合 + 规范化名 → 飞书返回的精确列名（避免空格/不可见字符导致误判缺失） */
  async function feishuBitableBuildFieldIndex(appToken, tableId, accessToken) {
    const exactSet = new Set();
    const normToExact = new Map();
    let pt = '';
    for (let guard = 0; guard < 50; guard++) {
      const page = await feishuBitableListFieldsPage(appToken, tableId, accessToken, pt);
      for (let i = 0; i < page.items.length; i++) {
        const it = page.items[i];
        const exact = (it && it.field_name != null ? String(it.field_name) : '').trim();
        if (!exact) continue;
        exactSet.add(exact);
        const nk = normFieldName(exact);
        if (!normToExact.has(nk)) normToExact.set(nk, exact);
      }
      if (!page.hasMore || !page.pageToken) break;
      pt = page.pageToken;
      await geSleep(120);
    }
    return { exactSet, normToExact };
  }

  function fieldIndexHasTitle(index, desiredTitle) {
    const t = String(desiredTitle || '').trim();
    if (!t) return true;
    if (index.exactSet.has(t)) return true;
    return index.normToExact.has(normFieldName(t));
  }

  function fieldIndexResolveTitle(index, desiredTitle) {
    const t = String(desiredTitle || '').trim();
    if (index.exactSet.has(t)) return t;
    const nk = normFieldName(t);
    if (index.normToExact.has(nk)) return index.normToExact.get(nk);
    return t;
  }

  async function feishuBitableCreateField(appToken, tableId, accessToken, fieldName, typeNum) {
    const url =
      'https://open.feishu.cn/open-apis/bitable/v1/apps/' +
      encodeURIComponent(appToken) +
      '/tables/' +
      encodeURIComponent(tableId) +
      '/fields';
    const tok = accessToken.trim();
    const auth = tok.toLowerCase().startsWith('bearer ') ? tok : 'Bearer ' + tok;
    const r = await gmXhr({
      method: 'POST',
      url,
      headers: { 'Content-Type': 'application/json; charset=utf-8', Authorization: auth },
      body: JSON.stringify({ field_name: fieldName, type: typeNum }),
    });
    let json;
    try {
      json = JSON.parse(r.responseText || '{}');
    } catch (_) {
      throw new Error('新增字段返回非 JSON');
    }
    if (json.code === 0) return true;
    const c = Number(json.code);
    if (c === 1254014 || /duplicate|duplicated/i.test(String(json.msg || ''))) return false;
    throw new Error((json.msg || '新增字段失败') + ' code=' + json.code);
  }

  /**
   * 检查映射中的每一列是否在子表中存在（按列名规范化比较）；
   * 开启自动建列时缺失则调用接口新建；关闭时若有缺失则一次性报错。
   * @returns {{ created: number, resolveMap: Record<string, string> }} resolveMap：内部键 → 飞书实际列名
   */
  async function ensureFeishuBitableMissingColumns(bitableAppToken, tableId, cfgTok) {
    const merged = mergeFieldMap(cfgTok);
    const autoCreate = cfgTok.feishuAutoCreateFields !== false;
    const uploadKeys = getEnabledUploadFieldKeys(cfgTok);
    let index = await feishuBitableBuildFieldIndex(bitableAppToken, tableId, cfgTok.accessToken);

    const missing = [];
    for (let ki = 0; ki < uploadKeys.length; ki++) {
      const k = uploadKeys[ki];
      const colTitle = (merged[k] || '').trim();
      if (!colTitle) continue;
      if (!fieldIndexHasTitle(index, colTitle)) missing.push({ k: k, title: colTitle, typ: feishuBitableTypeForMappedKey(k, cfgTok) });
    }

    if (missing.length && !autoCreate) {
      throw new Error(
        '多维表格缺少以下列（与映射中的列名一致或仅差空格）：' +
          missing.map(function (m) {
            return m.title;
          }).join('、') +
          '。请在飞书配置中勾选「写入前自动创建缺失列」，或在表中手动添加后再写入。'
      );
    }

    let created = 0;
    for (let mi = 0; mi < missing.length; mi++) {
      const m = missing[mi];
      const ok = await feishuBitableCreateField(bitableAppToken, tableId, cfgTok.accessToken, m.title, m.typ);
      if (ok) {
        created++;
        index.exactSet.add(m.title);
        index.normToExact.set(normFieldName(m.title), m.title);
      } else {
        index = await feishuBitableBuildFieldIndex(bitableAppToken, tableId, cfgTok.accessToken);
      }
      await geSleep(150);
    }

    if (created > 0) {
      await geSleep(280);
      index = await feishuBitableBuildFieldIndex(bitableAppToken, tableId, cfgTok.accessToken);
    }

    const resolveMap = {};
    for (let ki = 0; ki < uploadKeys.length; ki++) {
      const k = uploadKeys[ki];
      const colTitle = (merged[k] || '').trim();
      if (!colTitle) continue;
      resolveMap[k] = fieldIndexResolveTitle(index, colTitle);
    }
    return { created, resolveMap };
  }

  async function feishuBatchCreate(appToken, tableId, accessToken, records) {
    const url =
      'https://open.feishu.cn/open-apis/bitable/v1/apps/' +
      encodeURIComponent(appToken) +
      '/tables/' +
      encodeURIComponent(tableId) +
      '/records/batch_create';
    const tok = accessToken.trim();
    const auth = tok.toLowerCase().startsWith('bearer ') ? tok : 'Bearer ' + tok;
    const r = await gmXhr({
      method: 'POST',
      url,
      headers: { 'Content-Type': 'application/json; charset=utf-8', Authorization: auth },
      body: JSON.stringify({ records }),
    });
    let json;
    try {
      json = JSON.parse(r.responseText || '{}');
    } catch (_) {
      throw new Error('飞书返回非 JSON（HTTP ' + r.status + '）');
    }
    if (json.code !== 0) {
      throw new Error((json.msg || '飞书写入失败') + ' code=' + json.code);
    }
    return json;
  }

  function parseHandPriceForFeishuNumber(v) {
    const t = String(v || '')
      .replace(/[¥￥,\s]/g, '')
      .trim();
    if (!t) return null;
    const n = parseFloat(t);
    return Number.isFinite(n) ? n : null;
  }

  function parseShopIdForFeishuNumber(v) {
    const t = String(v || '').replace(/\s/g, '');
    if (!t) return null;
    const n = parseInt(t, 10);
    return Number.isFinite(n) && String(n) === t ? n : null;
  }

  function parseMonthlySalesForFeishuNumber(v) {
    const t = String(v || '').replace(/[,，\s]/g, '');
    if (!t) return null;
    const n = parseInt(t, 10);
    return Number.isFinite(n) ? n : null;
  }

  function feishuCartCell(internalKey, raw, mergedMap, cfg) {
    const col = mergedMap[internalKey];
    if (!col) return null;
    const v = raw == null ? '' : String(raw).trim();
    if (!v) return null;
    if (cfg.feishuCoerceNumberFields) {
      if (internalKey === 'handPrice') {
        const n = parseHandPriceForFeishuNumber(v);
        return n !== null ? n : null;
      }
      if (internalKey === 'shopId') {
        const n = parseShopIdForFeishuNumber(v);
        return n !== null ? n : null;
      }
      if (internalKey === 'monthlySales') {
        const n = parseMonthlySalesForFeishuNumber(v);
        return n !== null ? n : null;
      }
    }
    const hyp =
      cfg.useHyperlink &&
      (internalKey === 'productLink' || internalKey === 'productImageLink') &&
      /^https?:\/\//i.test(v);
    if (hyp) return { text: v.length > 120 ? v.slice(0, 117) + '...' : v, link: v };
    return v;
  }

  function rowToFeishuCartFields(row, mergedMap, cfg, resolveMap) {
    const keys = getEnabledUploadFieldKeys(cfg);
    const fields = {};
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      const raw =
        k === 'productImageLink' ? row.productImageLink || row.image || '' : row[k];
      const cell = feishuCartCell(k, raw, mergedMap, cfg);
      if (cell === null || cell === '') continue;
      const feishuCol = resolveMap && resolveMap[k] ? resolveMap[k] : mergedMap[k];
      if (!feishuCol) continue;
      fields[feishuCol] = cell;
    }
    return fields;
  }

  async function uploadCartRowsToFeishu(rows, cfg) {
    const cfgTok = await ensureFeishuAccessToken(cfg);
    const bitableAppToken = await resolveBitableAppToken(cfgTok);
    if (!bitableAppToken) throw new Error('请填写「Wiki 节点 token」或「多维表格 app_token」');
    if (!(cfgTok.tableId || '').trim()) throw new Error('请填写子表 table_id');
    const merged = mergeFieldMap(cfgTok);
    const colResult = await ensureFeishuBitableMissingColumns(bitableAppToken, cfgTok.tableId, cfgTok);
    const resolveMap = colResult.resolveMap || {};
    const batches = chunk(rows, 100);
    let total = 0;
    for (let i = 0; i < batches.length; i++) {
      const records = batches[i].map(function (row) {
        return { fields: rowToFeishuCartFields(row, merged, cfgTok, resolveMap) };
      });
      await feishuBatchCreate(bitableAppToken, cfgTok.tableId, cfgTok.accessToken, records);
      total += records.length;
      if (i < batches.length - 1) await geSleep(320);
    }
    return { total, columnsCreated: colResult.created };
  }

  /**
   * 与侧栏「写入飞书」按钮相同逻辑；opts.showAlert 为 false 时不弹成功/清空提示（供抓取后自动写入时合并提示）。
   */
  async function geUploadCartToFeishuWithClearAfterFlow(opts) {
    const o = opts || {};
    const showAlert = o.showAlert !== false;
    const rows = dedupeFinalRows(
      STORE.rows.map(function (r) {
        return { ...r };
      })
    );
    if (!rows.length) {
      if (showAlert) alert('暂无数据，请等待列表加载或点「一键抓取（滚到底）」');
      throw new Error('暂无数据');
    }
    const fcfg = loadFeishuConfig();
    if (!(fcfg.wikiNodeToken || fcfg.appToken) || !fcfg.tableId) {
      if (showAlert) alert('请先点「飞书配置」填写 Wiki 节点 token 或 app_token，以及 table_id');
      throw new Error('缺少飞书配置');
    }
    if (!(fcfg.accessToken || '').trim() && !((fcfg.feishuAppId || '').trim() && (fcfg.feishuAppSecret || '').trim())) {
      if (showAlert) alert('请填写 access_token，或填写 App ID + App Secret');
      throw new Error('缺少鉴权');
    }
    rows.forEach(fillProductLinkFromId);
    const res = await uploadCartRowsToFeishu(rows, fcfg);
    if (showAlert) {
      let msg = '已写入飞书 ' + res.total + ' 条';
      if (res.columnsCreated) msg += '（新建列 ' + res.columnsCreated + ' 个）';
      alert(msg);
    }
    const chkClear = document.getElementById('ge-buyin-clear-after');
    const chkClearAuto = document.getElementById('ge-buyin-clear-auto');
    if (chkClear && chkClear.checked) {
      const skipConfirm = chkClearAuto && chkClearAuto.checked;
      const goClear =
        skipConfirm ||
        confirm(
          '是否在网页上全选并删除当前选品车中的商品？（仅当前页已加载条目；多页需切换分页后再次写入并清空）'
        );
      if (goClear) {
        try {
          await clearBuyinCartDom();
          STORE.rows = [];
          STORE.keySet.clear();
          updateStatus();
          if (showAlert) {
            alert(skipConfirm ? '已按设置自动清空当前页选品车；若有残留请翻页后重复或刷新' : '已尝试清空；若仍有残留请刷新页面再试');
          }
        } catch (e) {
          if (showAlert) alert('清空选品车失败：' + (e && e.message ? e.message : e));
          throw e;
        }
      }
    }
    return res;
  }

  function collectFeishuFieldMapFromForm(box) {
    const o = {};
    for (let i = 0; i < FEISHU_UPLOAD_FIELD_KEYS.length; i++) {
      const k = FEISHU_UPLOAD_FIELD_KEYS[i];
      const inp = box.querySelector('[data-ge-map-key="' + k + '"]');
      const v = inp ? String(inp.value || '').trim() : '';
      o[k] = v || FEISHU_DEFAULT_FIELD_MAP[k];
    }
    return o;
  }

  function collectFeishuFieldEnabledFromForm(box) {
    const out = {};
    for (let i = 0; i < FEISHU_UPLOAD_FIELD_KEYS.length; i++) {
      const k = FEISHU_UPLOAD_FIELD_KEYS[i];
      const ck = box.querySelector('[data-ge-enable-key="' + k + '"]');
      out[k] = !!(ck && ck.checked);
    }
    return out;
  }

  function geApplyMapInputEnabled(inp, enabled) {
    if (!inp) return;
    inp.disabled = !enabled;
    inp.style.opacity = enabled ? '1' : '0.55';
    inp.style.background = enabled ? '#fff' : '#f2f3f5';
  }

  function syncFeishuMapJsonTextarea(box) {
    const ta = box.querySelector('#ge-bf-map');
    if (!ta) return;
    try {
      ta.value = JSON.stringify(collectFeishuFieldMapFromForm(box), null, 2);
    } catch (_) {}
  }

  function syncFeishuFieldEnabledJsonTextarea(box) {
    const ta = box.querySelector('#ge-bf-field-enabled-json');
    if (!ta) return;
    try {
      ta.value = JSON.stringify(collectFeishuFieldEnabledFromForm(box), null, 2);
    } catch (_) {}
  }

  function buildFeishuMapFormRows(box, cfg) {
    const wrap = box.querySelector('#ge-bf-map-ui');
    if (!wrap) return;
    const merged = mergeFieldMap(cfg);
    const enabled = parseFieldEnabledJson(cfg.fieldEnabledJson);
    wrap.innerHTML = '';
    for (let i = 0; i < FEISHU_UPLOAD_FIELD_KEYS.length; i++) {
      const k = FEISHU_UPLOAD_FIELD_KEYS[i];
      const row = document.createElement('div');
      row.className = 'ge-bf-map-row';
      row.style.cssText =
        'display:grid;grid-template-columns:22px minmax(72px,30%) 1fr;gap:8px 10px;align-items:center;margin-bottom:10px';
      const ck = document.createElement('input');
      ck.type = 'checkbox';
      ck.setAttribute('data-ge-enable-key', k);
      ck.checked = enabled[k] !== false;
      ck.title = '勾选：导出 CSV / 写入飞书 / 检查列';
      const lab = document.createElement('div');
      lab.style.cssText = 'font-size:12px;color:#4e5969;line-height:1.4';
      lab.textContent = FEISHU_FIELD_UI_LABELS[k] || k;
      lab.title = '内部字段 ' + k + ' → 飞书列标题';
      const inp = document.createElement('input');
      inp.type = 'text';
      inp.setAttribute('data-ge-map-key', k);
      inp.style.cssText =
        'width:100%;box-sizing:border-box;padding:8px 10px;border:1px solid #e5e6eb;border-radius:6px;font-size:13px';
      inp.placeholder = FEISHU_DEFAULT_FIELD_MAP[k] || '';
      inp.value = merged[k] || FEISHU_DEFAULT_FIELD_MAP[k] || '';
      geApplyMapInputEnabled(inp, ck.checked);
      row.appendChild(ck);
      row.appendChild(lab);
      row.appendChild(inp);
      wrap.appendChild(row);
    }
    wrap.addEventListener('input', function () {
      syncFeishuMapJsonTextarea(box);
    });
    wrap.addEventListener('change', function (e) {
      const t = e.target;
      if (!t || !t.getAttribute) return;
      const ek = t.getAttribute('data-ge-enable-key');
      if (!ek) return;
      const inp = box.querySelector('[data-ge-map-key="' + ek + '"]');
      geApplyMapInputEnabled(inp, !!t.checked);
      syncFeishuFieldEnabledJsonTextarea(box);
    });
    syncFeishuFieldEnabledJsonTextarea(box);
  }

  function bindFeishuFieldEnableShortcutButtons(box) {
    function setAll(on) {
      for (let i = 0; i < FEISHU_UPLOAD_FIELD_KEYS.length; i++) {
        const k = FEISHU_UPLOAD_FIELD_KEYS[i];
        const ck = box.querySelector('[data-ge-enable-key="' + k + '"]');
        const inp = box.querySelector('[data-ge-map-key="' + k + '"]');
        if (ck) {
          ck.checked = on;
          geApplyMapInputEnabled(inp, on);
        }
      }
      syncFeishuFieldEnabledJsonTextarea(box);
    }
    const all = box.querySelector('#ge-bf-map-enable-all');
    const none = box.querySelector('#ge-bf-map-enable-none');
    const rst = box.querySelector('#ge-bf-map-enable-reset');
    if (all) all.addEventListener('click', function () { setAll(true); });
    if (none) none.addEventListener('click', function () { setAll(false); });
    if (rst) rst.addEventListener('click', function () { setAll(true); });
  }

  function showFeishuCartConfigModal() {
    if (document.getElementById('ge-buyin-feishu-modal')) return;
    const cfg = loadFeishuConfig();
    const mask = document.createElement('div');
    mask.id = 'ge-buyin-feishu-modal';
    mask.style.cssText =
      'position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:2147483646;display:flex;align-items:center;justify-content:center;padding:12px;box-sizing:border-box';
    const box = document.createElement('div');
    box.style.cssText =
      'background:#fff;border-radius:12px;max-width:min(640px,calc(100vw - 24px));width:100%;max-height:92vh;overflow:auto;padding:16px;font:13px/1.5 -apple-system,sans-serif;color:#1f2329';
    box.innerHTML =
      '<div style="font-weight:700;font-size:16px;margin-bottom:10px">选品车 → 飞书多维表格</div>' +
      '<p style="margin:0 0 10px;color:#646a73;font-size:12px">与项目内「百应+淘宝 商品一键导出」相同：开放平台应用需开通 bitable:app，并把应用加入表格协作者。写入前会<strong>拉取子表全部字段</strong>并与下方<strong>列名映射</strong>比对；勾选「缺列则自动新建」时按右侧填写的名称创建字段。地址栏 <code>/base/xxx</code> 为 app_token，<code>table=</code> 后为 table_id；Wiki 打开则填节点 token。</p>' +
      '<label style="display:block;margin:6px 0">Wiki 节点 token（与 app_token 二选一）</label>' +
      '<input id="ge-bf-wiki" type="text" style="width:100%;box-sizing:border-box;padding:8px;border:1px solid #e5e6eb;border-radius:6px" placeholder="/wiki/ 与 ? 之间" />' +
      '<label style="display:block;margin:8px 0 6px">多维表格 app_token（非 Wiki 时）</label>' +
      '<input id="ge-bf-app" type="text" style="width:100%;box-sizing:border-box;padding:8px;border:1px solid #e5e6eb;border-radius:6px" />' +
      '<label style="display:block;margin:8px 0 6px">子表 table_id</label>' +
      '<input id="ge-bf-tbl" type="text" style="width:100%;box-sizing:border-box;padding:8px;border:1px solid #e5e6eb;border-radius:6px" />' +
      '<label style="display:block;margin:8px 0 6px">App ID / App Secret（可选，填写则每次写入前自动换 token）</label>' +
      '<input id="ge-bf-aid" type="text" style="width:100%;box-sizing:border-box;padding:8px;border:1px solid #e5e6eb;border-radius:6px;margin-bottom:6px" placeholder="cli_xxx" />' +
      '<input id="ge-bf-sec" type="password" style="width:100%;box-sizing:border-box;padding:8px;border:1px solid #e5e6eb;border-radius:6px" placeholder="Secret" />' +
      '<label style="display:block;margin:8px 0 6px">access_token（tenant_access_token；若已填 App ID/Secret 可仅作备用）</label>' +
      '<input id="ge-bf-tok" type="password" style="width:100%;box-sizing:border-box;padding:8px;border:1px solid #e5e6eb;border-radius:6px" />' +
      '<div style="font-weight:600;margin:16px 0 6px;font-size:13px;color:#1f2329">飞书列名映射</div>' +
      '<p style="margin:0 0 12px;color:#646a73;font-size:11px;line-height:1.55">左起：勾选启用 → 字段含义 → 飞书列标题。未勾选字段不导出 CSV、不写入飞书、不检查/自动建列。「恢复默认列名」仅重置列名，不改变勾选。</p>' +
      '<div id="ge-bf-map-ui" class="ge-bf-map-ui"></div>' +
      '<div style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px;align-items:center">' +
      '<button type="button" id="ge-bf-map-reset" style="padding:6px 12px;border:1px solid #dcdfe6;border-radius:6px;background:#f7f8fa;cursor:pointer;font-size:12px">恢复默认列名</button>' +
      '<button type="button" id="ge-bf-map-enable-all" style="padding:6px 12px;border:1px solid #dcdfe6;border-radius:6px;background:#f7f8fa;cursor:pointer;font-size:12px">全选</button>' +
      '<button type="button" id="ge-bf-map-enable-none" style="padding:6px 12px;border:1px solid #dcdfe6;border-radius:6px;background:#f7f8fa;cursor:pointer;font-size:12px">全不选</button>' +
      '<button type="button" id="ge-bf-map-enable-reset" style="padding:6px 12px;border:1px solid #dcdfe6;border-radius:6px;background:#f7f8fa;cursor:pointer;font-size:12px">恢复默认勾选</button>' +
      '</div>' +
      '<details style="margin:4px 0 12px;border:1px solid #e5e6eb;border-radius:8px;padding:10px 12px;background:#fafbfc">' +
      '<summary style="cursor:pointer;font-size:12px;color:#3370ff;font-weight:500;outline:none">高级：JSON 映射（与表单双向同步）</summary>' +
      '<textarea id="ge-bf-map" style="width:100%;box-sizing:border-box;min-height:140px;margin-top:10px;padding:8px;border:1px solid #e5e6eb;border-radius:6px;font:12px Consolas,monospace"></textarea>' +
      '<p style="margin:8px 0 4px;font-size:11px;color:#646a73">字段启用 JSON（与左侧勾选同步）</p>' +
      '<textarea id="ge-bf-field-enabled-json" style="width:100%;box-sizing:border-box;min-height:72px;margin-top:4px;padding:8px;border:1px solid #e5e6eb;border-radius:6px;font:12px Consolas,monospace"></textarea>' +
      '<p style="margin:8px 0;font-size:11px;color:#646a73">保存时以<strong>上方表单</strong>为准。若改了 JSON，请先点「从 JSON 应用到表单」再保存。</p>' +
      '<button type="button" id="ge-bf-map-apply-json" style="padding:6px 12px;border:1px solid #3370ff;border-radius:6px;background:#fff;color:#3370ff;cursor:pointer;font-size:12px">从 JSON 应用到表单</button>' +
      '</details>' +
      '<div style="margin-top:10px;display:flex;flex-wrap:wrap;gap:12px;align-items:center;font-size:12px">' +
      '<label><input type="checkbox" id="ge-bf-hyp" /> 商品链接用超链接列</label>' +
      '<label><input type="checkbox" id="ge-bf-num" /> 到手价、店铺ID、月销按数字列写入</label>' +
      '<label><input type="checkbox" id="ge-bf-auto" checked /> 写入前列检查：缺列则自动新建</label>' +
      '</div>' +
      '<div style="margin-top:14px;display:flex;gap:8px;justify-content:flex-end">' +
      '<button type="button" id="ge-bf-cancel" style="padding:8px 16px;border:1px solid #dcdfe6;border-radius:6px;background:#fff;cursor:pointer">取消</button>' +
      '<button type="button" id="ge-bf-save" style="padding:8px 16px;border:none;border-radius:6px;background:#3370ff;color:#fff;cursor:pointer">保存</button>' +
      '</div>';
    mask.appendChild(box);
    document.body.appendChild(mask);
    const $ = function (id) {
      return document.getElementById(id);
    };
    $('ge-bf-wiki').value = cfg.wikiNodeToken;
    $('ge-bf-app').value = cfg.appToken;
    $('ge-bf-tbl').value = cfg.tableId;
    $('ge-bf-aid').value = cfg.feishuAppId;
    $('ge-bf-sec').value = cfg.feishuAppSecret;
    $('ge-bf-tok').value = cfg.accessToken;
    $('ge-bf-hyp').checked = cfg.useHyperlink;
    $('ge-bf-num').checked = cfg.feishuCoerceNumberFields;
    $('ge-bf-auto').checked = cfg.feishuAutoCreateFields !== false;

    buildFeishuMapFormRows(box, cfg);
    syncFeishuMapJsonTextarea(box);
    bindFeishuFieldEnableShortcutButtons(box);

    function close() {
      mask.remove();
    }
    $('ge-bf-cancel').addEventListener('click', close);
    mask.addEventListener('click', function (e) {
      if (e.target === mask) close();
    });
    box.querySelector('#ge-bf-map-reset').addEventListener('click', function () {
      for (let i = 0; i < FEISHU_UPLOAD_FIELD_KEYS.length; i++) {
        const k = FEISHU_UPLOAD_FIELD_KEYS[i];
        const inp = box.querySelector('[data-ge-map-key="' + k + '"]');
        if (inp) inp.value = FEISHU_DEFAULT_FIELD_MAP[k] || '';
      }
      syncFeishuMapJsonTextarea(box);
    });
    box.querySelector('#ge-bf-map-apply-json').addEventListener('click', function () {
      const ta = box.querySelector('#ge-bf-map');
      let o;
      try {
        o = JSON.parse(ta.value || '{}');
      } catch (e) {
        alert('JSON 无法解析，请检查格式');
        return;
      }
      if (!o || typeof o !== 'object') {
        alert('JSON 须为对象');
        return;
      }
      for (let i = 0; i < FEISHU_UPLOAD_FIELD_KEYS.length; i++) {
        const k = FEISHU_UPLOAD_FIELD_KEYS[i];
        if (typeof o[k] !== 'string' || !String(o[k]).trim()) continue;
        const inp = box.querySelector('[data-ge-map-key="' + k + '"]');
        if (inp) inp.value = String(o[k]).trim();
      }
      syncFeishuMapJsonTextarea(box);
      alert('已应用到表单');
    });
    $('ge-bf-save').addEventListener('click', function () {
      const mapObj = collectFeishuFieldMapFromForm(box);
      const enabledObj = collectFeishuFieldEnabledFromForm(box);
      saveFeishuConfig({
        wikiNodeToken: $('ge-bf-wiki').value,
        appToken: $('ge-bf-app').value,
        tableId: $('ge-bf-tbl').value,
        feishuAppId: $('ge-bf-aid').value,
        feishuAppSecret: $('ge-bf-sec').value,
        accessToken: $('ge-bf-tok').value,
        fieldMapJson: JSON.stringify(mapObj),
        fieldEnabledJson: JSON.stringify(enabledObj),
        useHyperlink: $('ge-bf-hyp').checked,
        feishuCoerceNumberFields: $('ge-bf-num').checked,
        feishuAutoCreateFields: $('ge-bf-auto').checked,
      });
      alert('已保存');
      close();
    });
  }

  /** 仅调整「内部字段 → 飞书列标题」，不打开完整 Wiki/app 配置 */
  function showFeishuColumnMapOnlyModal() {
    if (document.getElementById('ge-buyin-feishu-map-modal')) return;
    const cfg = loadFeishuConfig();
    const mask = document.createElement('div');
    mask.id = 'ge-buyin-feishu-map-modal';
    mask.style.cssText =
      'position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:2147483646;display:flex;align-items:center;justify-content:center;padding:12px;box-sizing:border-box';
    const box = document.createElement('div');
    box.style.cssText =
      'background:#fff;border-radius:12px;max-width:min(520px,calc(100vw - 24px));width:100%;max-height:88vh;overflow:auto;padding:16px;font:13px/1.5 -apple-system,sans-serif;color:#1f2329';
    box.innerHTML =
      '<div style="font-weight:700;font-size:15px;margin-bottom:8px">飞书列名映射</div>' +
      '<p style="margin:0 0 12px;color:#646a73;font-size:12px;line-height:1.55">勾选启用 → 列标题映射。未勾选字段不导出、不写入飞书。店铺链接与来源列始终出现在 CSV，不在此勾选。</p>' +
      '<div id="ge-bf-map-ui" class="ge-bf-map-ui"></div>' +
      '<div style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px;align-items:center">' +
      '<button type="button" id="ge-bf-map-reset" style="padding:6px 12px;border:1px solid #dcdfe6;border-radius:6px;background:#f7f8fa;cursor:pointer;font-size:12px">恢复默认列名</button>' +
      '<button type="button" id="ge-bf-map-enable-all" style="padding:6px 12px;border:1px solid #dcdfe6;border-radius:6px;background:#f7f8fa;cursor:pointer;font-size:12px">全选</button>' +
      '<button type="button" id="ge-bf-map-enable-none" style="padding:6px 12px;border:1px solid #dcdfe6;border-radius:6px;background:#f7f8fa;cursor:pointer;font-size:12px">全不选</button>' +
      '<button type="button" id="ge-bf-map-enable-reset" style="padding:6px 12px;border:1px solid #dcdfe6;border-radius:6px;background:#f7f8fa;cursor:pointer;font-size:12px">恢复默认勾选</button>' +
      '</div>' +
      '<details style="margin:4px 0 12px;border:1px solid #e5e6eb;border-radius:8px;padding:10px 12px;background:#fafbfc">' +
      '<summary style="cursor:pointer;font-size:12px;color:#3370ff;font-weight:500;outline:none">高级：JSON 映射</summary>' +
      '<textarea id="ge-bf-map" style="width:100%;box-sizing:border-box;min-height:120px;margin-top:10px;padding:8px;border:1px solid #e5e6eb;border-radius:6px;font:12px Consolas,monospace"></textarea>' +
      '<p style="margin:8px 0 4px;font-size:11px;color:#646a73">字段启用 JSON</p>' +
      '<textarea id="ge-bf-field-enabled-json" style="width:100%;box-sizing:border-box;min-height:72px;margin-top:4px;padding:8px;border:1px solid #e5e6eb;border-radius:6px;font:12px Consolas,monospace"></textarea>' +
      '<button type="button" id="ge-bf-map-apply-json" style="margin-top:8px;padding:6px 12px;border:1px solid #3370ff;border-radius:6px;background:#fff;color:#3370ff;cursor:pointer;font-size:12px">从 JSON 应用到表单</button>' +
      '</details>' +
      '<div style="margin-top:12px;display:flex;gap:8px;justify-content:flex-end">' +
      '<button type="button" id="ge-bf2-cancel" style="padding:8px 16px;border:1px solid #dcdfe6;border-radius:6px;background:#fff;cursor:pointer">取消</button>' +
      '<button type="button" id="ge-bf2-save" style="padding:8px 16px;border:none;border-radius:6px;background:#3370ff;color:#fff;cursor:pointer">保存列名</button>' +
      '</div>';
    mask.appendChild(box);
    document.body.appendChild(mask);

    buildFeishuMapFormRows(box, cfg);
    syncFeishuMapJsonTextarea(box);
    bindFeishuFieldEnableShortcutButtons(box);

    function close() {
      mask.remove();
    }
    box.querySelector('#ge-bf2-cancel').addEventListener('click', close);
    mask.addEventListener('click', function (e) {
      if (e.target === mask) close();
    });
    box.querySelector('#ge-bf-map-reset').addEventListener('click', function () {
      for (let i = 0; i < FEISHU_UPLOAD_FIELD_KEYS.length; i++) {
        const k = FEISHU_UPLOAD_FIELD_KEYS[i];
        const inp = box.querySelector('[data-ge-map-key="' + k + '"]');
        if (inp) inp.value = FEISHU_DEFAULT_FIELD_MAP[k] || '';
      }
      syncFeishuMapJsonTextarea(box);
    });
    box.querySelector('#ge-bf-map-apply-json').addEventListener('click', function () {
      const ta = box.querySelector('#ge-bf-map');
      let o;
      try {
        o = JSON.parse(ta.value || '{}');
      } catch (e) {
        alert('JSON 无法解析，请检查格式');
        return;
      }
      if (!o || typeof o !== 'object') {
        alert('JSON 须为对象');
        return;
      }
      for (let i = 0; i < FEISHU_UPLOAD_FIELD_KEYS.length; i++) {
        const k = FEISHU_UPLOAD_FIELD_KEYS[i];
        if (typeof o[k] !== 'string' || !String(o[k]).trim()) continue;
        const inp = box.querySelector('[data-ge-map-key="' + k + '"]');
        if (inp) inp.value = String(o[k]).trim();
      }
      syncFeishuMapJsonTextarea(box);
      alert('已应用到表单');
    });
    box.querySelector('#ge-bf2-save').addEventListener('click', function () {
      const mapObj = collectFeishuFieldMapFromForm(box);
      const enabledObj = collectFeishuFieldEnabledFromForm(box);
      const latest = loadFeishuConfig();
      saveFeishuConfig({
        ...latest,
        fieldMapJson: JSON.stringify(mapObj),
        fieldEnabledJson: JSON.stringify(enabledObj),
      });
      alert('列名映射已保存');
      close();
    });
  }

  /** 删除选品车期间站点常调 location.reload / 同页 replace，导致脚本中断；临时拦截 */
  let GE_CART_RELOAD_GUARD_DEPTH = 0;

  function geCartReloadGuardEnter() {
    GE_CART_RELOAD_GUARD_DEPTH++;
  }

  function geCartReloadGuardLeave() {
    GE_CART_RELOAD_GUARD_DEPTH = Math.max(0, GE_CART_RELOAD_GUARD_DEPTH - 1);
  }

  function geSameDocumentUrl(url) {
    try {
      const next = new URL(String(url), location.href);
      const cur = new URL(location.href);
      return next.origin === cur.origin && next.pathname === cur.pathname && next.search === cur.search;
    } catch (_) {
      return false;
    }
  }

  function geInstallCartReloadGuardOnce() {
    if (window._geBuyinCartReloadGuard) return;
    window._geBuyinCartReloadGuard = true;
    try {
      const origReload = location.reload.bind(location);
      location.reload = function () {
        if (GE_CART_RELOAD_GUARD_DEPTH > 0) {
          console.warn('[选品车抓取] 已阻止删除流程中的 location.reload()');
          return;
        }
        return origReload.apply(location, arguments);
      };
    } catch (_) {}
    try {
      const proto = Location.prototype;
      const origReplace = proto.replace;
      const origAssign = proto.assign;
      if (typeof origReplace === 'function') {
        proto.replace = function (url) {
          if (GE_CART_RELOAD_GUARD_DEPTH > 0 && geSameDocumentUrl(url)) {
            console.warn('[选品车抓取] 已阻止删除流程中的 location.replace(同页)');
            return;
          }
          return origReplace.apply(this, arguments);
        };
      }
      if (typeof origAssign === 'function') {
        proto.assign = function (url) {
          if (GE_CART_RELOAD_GUARD_DEPTH > 0 && geSameDocumentUrl(url)) {
            console.warn('[选品车抓取] 已阻止删除流程中的 location.assign(同页)');
            return;
          }
          return origAssign.apply(this, arguments);
        };
      }
    } catch (_) {}
    try {
      const origGo = History.prototype.go;
      History.prototype.go = function (delta) {
        if (GE_CART_RELOAD_GUARD_DEPTH > 0 && (delta === 0 || delta === undefined)) {
          console.warn('[选品车抓取] 已阻止删除流程中的 history.go(刷新)');
          return;
        }
        return origGo.apply(this, arguments);
      };
    } catch (_) {}
  }

  async function clearBuyinCartDom() {
    geInstallCartReloadGuardOnce();
    geCartReloadGuardEnter();
    try {
    const bars = Array.from(document.querySelectorAll('div')).filter(function (div) {
      const r = div.getBoundingClientRect();
      if (r.width < 200 || r.height < 24) return false;
      const t = div.innerText || '';
      if (!/全选/.test(t) || !/已选/.test(t) || !/删除/.test(t)) return false;
      return r.bottom >= window.innerHeight - 200;
    });
    bars.sort(function (a, b) {
      return b.getBoundingClientRect().bottom - a.getBoundingClientRect().bottom;
    });
    const bar = bars[0];
    if (!bar) throw new Error('未找到底部操作栏（请打开选品车页并可见列表）');
    const cb = bar.querySelector('input[type="checkbox"]');
    if (cb) {
      if (!cb.checked) {
        cb.click();
        await geSleep(500);
      }
    } else {
      const lab = Array.from(bar.querySelectorAll('label,span,div')).find(function (el) {
        return /^全选/.test((el.textContent || '').trim());
      });
      if (lab) {
        lab.click();
        await geSleep(500);
      }
    }
    await geSleep(200);
    const del = Array.from(bar.querySelectorAll('button, [role="button"], span, div')).find(function (el) {
      const t = (el.textContent || '').replace(/\s/g, '');
      return t === '删除';
    });
    if (!del) throw new Error('未找到批量「删除」按钮');
    del.click();
    await geSleep(700);
    await geConfirmCartDeleteDialogIfAny();
    } finally {
      geCartReloadGuardLeave();
    }
  }

  /** 取当前最上层可见弹层（避免点到被遮挡的旧 Modal） */
  function geGetTopVisibleDialogRoot() {
    const dialogs = Array.from(
      document.querySelectorAll('[role="dialog"], [class*="Modal"], [class*="modal"], [class*="Dialog"]')
    );
    const visible = dialogs.filter(function (d) {
      try {
        const r = d.getBoundingClientRect();
        const st = window.getComputedStyle(d);
        if (r.width < 8 || r.height < 8) return false;
        if (st.visibility === 'hidden' || st.display === 'none' || Number(st.opacity) === 0) return false;
        return true;
      } catch (_) {
        return false;
      }
    });
    return visible.length ? visible[visible.length - 1] : null;
  }

  async function geConfirmCartDeleteDialogIfAny() {
    for (let attempt = 0; attempt < 6; attempt++) {
      if (attempt > 0) await geSleep(450);
      else await geSleep(600);
      const dialog = geGetTopVisibleDialogRoot();
      const scope = dialog || document.body;
      const pool = Array.from(scope.querySelectorAll('button, span, [role="button"], a, div'));
      let confirmBtn = pool.find(function (el) {
        const raw = (el.textContent || '').replace(/\s+/g, '');
        if (!raw) return false;
        if (/^(确认删除|确定删除)$/.test(raw)) return true;
        if (/^(确定|确认)$/.test(raw)) return true;
        if (dialog && /^删除$/.test(raw) && dialog.contains(el)) return true;
        return false;
      });
      if (!confirmBtn && dialog) {
        const primary = dialog.querySelector(
          'button[type="button"], button[class*="primary"], button[class*="Primary"], .ant-btn-primary'
        );
        if (primary && /确认|删除/.test(s(primary.textContent))) confirmBtn = primary;
      }
      if (confirmBtn) {
        try {
          confirmBtn.scrollIntoView({ block: 'center', inline: 'nearest', behavior: 'auto' });
        } catch (_) {}
        await geSleep(60);
        try {
          confirmBtn.click();
        } catch (_) {}
        return;
      }
    }
  }

  /** 操作列内「删除」：避免点到商品卡/标题上的外链 <a> */
  function geFindOperationColumnRoot(rowRoot) {
    if (!rowRoot) return null;
    try {
      const by =
        rowRoot.querySelector('[class*="btnWrap"]') ||
        rowRoot.querySelector('[class*="btn_line"]') ||
        rowRoot.querySelector('[class*="btnLine"]') ||
        rowRoot.querySelector('[class*="operation"]') ||
        rowRoot.querySelector('[class*="Operation"]') ||
        rowRoot.querySelector('[class*="action"]');
      return by || rowRoot;
    } catch (_) {
      return rowRoot;
    }
  }

  function geAnchorLooksLikePageNavigation(a) {
    if (!a || a.tagName !== 'A') return false;
    const h = s(a.getAttribute('href'));
    if (!h || h === '#' || /^javascript:/i.test(h)) return false;
    if (/^https?:\/\//i.test(h)) {
      try {
        const u = new URL(h, location.href);
        if (u.pathname.indexOf('merch') >= 0 || u.pathname.indexOf('shop') >= 0) return true;
        if (u.hostname !== location.hostname && !/jinritemai|douyin|bytedance/i.test(u.hostname)) return true;
      } catch (_) {
        return true;
      }
    } else if (h[0] === '/') return true;
    return false;
  }

  function geFindDeleteButtonInCartRow(rowRoot) {
    if (!rowRoot) return null;
    try {
      const op = geFindOperationColumnRoot(rowRoot);
      const candidates = Array.from(op.querySelectorAll('span, button, a, [role="button"], div'));
      const hits = [];
      for (let i = 0; i < candidates.length; i++) {
        const el = candidates[i];
        const raw = s(el.textContent).replace(/\s+/g, '');
        if (raw !== '删除') continue;
        if (el.tagName === 'A' && geAnchorLooksLikePageNavigation(el)) continue;
        const len = s(el.textContent).length;
        hits.push({ el: el, len: len, tag: el.tagName });
      }
      if (!hits.length) {
        for (let j = 0; j < candidates.length; j++) {
          const el = candidates[j];
          if (el.tagName !== 'SPAN' && el.tagName !== 'BUTTON') continue;
          const raw = s(el.textContent).replace(/\s+/g, '');
          if (raw !== '删除') continue;
          hits.push({ el: el, len: s(el.textContent).length, tag: el.tagName });
        }
      }
      if (!hits.length) return null;
      hits.sort(function (a, b) {
        if (a.len !== b.len) return a.len - b.len;
        const order = { SPAN: 0, BUTTON: 1, A: 2, DIV: 3 };
        return (order[a.tag] || 9) - (order[b.tag] || 9);
      });
      return hits[0].el;
    } catch (_) {
      return null;
    }
  }

  async function geClickDeleteInRowThenConfirm(rowRoot) {
    const btn = geFindDeleteButtonInCartRow(rowRoot);
    if (!btn) return false;
    try {
      rowRoot.scrollIntoView({ block: 'center', inline: 'nearest', behavior: 'smooth' });
    } catch (_) {}
    await geSleep(120);
    try {
      btn.scrollIntoView({ block: 'center', inline: 'nearest', behavior: 'auto' });
    } catch (_) {}
    await geSleep(160);
    try {
      btn.click();
    } catch (_) {
      return false;
    }
    await geSleep(650);
    await geConfirmCartDeleteDialogIfAny();
    return true;
  }

  /**
   * 双空失效行无法全选批量删：由后往前逐条点「删除」，弹窗需点「确认删除」等，每删一条重扫 DOM。
   * @returns {number} 成功点击删除的次数
   */
  async function geDeleteInvalidCartRowsDom() {
    geInstallCartReloadGuardOnce();
    geCartReloadGuardEnter();
    try {
    let total = 0;
    const maxPasses = 120;
    for (let pass = 0; pass < maxPasses; pass++) {
      const docs = getAllDocs();
      let deletedThisPass = false;
      outer: for (let d = 0; d < docs.length; d++) {
        const dd = docs[d];
        const cards = getCardsInDoc(dd);
        for (let i = cards.length - 1; i >= 0; i--) {
          const row = await parseCardAsync(cards[i], dd);
          if (geIsValidCartExportRow(row)) continue;
          const ok = await geClickDeleteInRowThenConfirm(cards[i]);
          if (!ok) continue;
          total++;
          deletedThisPass = true;
          await geSleep(400);
          break outer;
        }
      }
      if (!deletedThisPass) break;
    }
    return total;
    } finally {
      geCartReloadGuardLeave();
    }
  }

  function updateStatus() {
    const node = document.getElementById('ge-cart-status');
    if (node) node.textContent = '已抓取：' + STORE.rows.length;
  }

  function geClearCartProgress() {
    const line = document.getElementById('ge-cart-progress');
    if (!line) return;
    line.textContent = '';
    line.style.display = 'none';
  }

  function geBuildCartDebugPayload() {
    return {
      rows: STORE.rows,
      keys: Array.from(STORE.keySet),
      rowKeys: STORE.rows.map(function (r) {
        return { title: r.productInfo, pid: r.productId, key: rowDedupeKey(r) };
      }),
      flags: { GE_USE_CARDS_V2: GE_USE_CARDS_V2, GE_CART_DEBUG: GE_CART_DEBUG },
    };
  }

  function mountPanel() {
    if (STORE.mounted || document.getElementById('ge-cart-panel')) return;
    STORE.mounted = true;
    const panel = document.createElement('div');
    panel.id = 'ge-cart-panel';
    const clearAfter = gmGet(GE_CLEAR_AFTER_FEISHU_KEY, '1') === '1';
    const clearAutoSaved = gmGet(GE_CLEAR_AFTER_FEISHU_AUTO_KEY, '0') === '1';
    const clearAuto = clearAfter && clearAutoSaved;
    const autoFeishuAfterScan = gmGet(GE_AUTO_FEISHU_AFTER_SCAN_KEY, '0') === '1';
    const deleteInvalidAfterScan = gmGet(GE_DELETE_INVALID_AFTER_SCAN_KEY, '0') === '1';
    panel.innerHTML =
      '<div class="ge-hd">选品车抓取</div>' +
      '<p class="ge-tip">请先打开选品车列表，再点「一键抓取」：会先<strong>只在商品列表区域内滚动</strong>（不滚整页，避免双滚动条/白屏），直到出现「没有更多」再解析。<strong>商品ID与店铺ID不能同时为空</strong>才会进入导出/飞书。</p>' +
      '<div id="ge-cart-status" class="ge-st">已抓取：0</div>' +
      '<div id="ge-cart-progress" class="ge-progress" style="display:none"></div>' +
      '<details class="ge-fields"><summary>默认写入字段（列名可改）</summary>' +
      '<ul class="ge-fields-ul">' +
      '<li>商品信息、商品ID、到手价、佣金率、商品链接</li>' +
      '<li>店铺、店铺ID、体验分、好评、月销、商品图片链接</li>' +
      '</ul></details>' +
      '<div class="ge-row">' +
      '<button id="ge-cart-scan" type="button" class="ge-btn-primary">一键抓取（滚到底）</button>' +
      '<button id="ge-cart-export" type="button">导出CSV</button>' +
      '</div>' +
      '<div class="ge-row">' +
      '<button id="ge-cart-feishu-cfg" type="button">飞书配置</button>' +
      '<button id="ge-cart-feishu-upload" type="button" class="ge-btn-primary">写入飞书</button>' +
      '</div>' +
      '<div class="ge-row">' +
      '<button id="ge-cart-feishu-map" type="button">列名映射</button>' +
      '</div>' +
      '<label class="ge-chk"><input type="checkbox" id="ge-buyin-auto-feishu-after-scan" ' +
      (autoFeishuAfterScan ? 'checked' : '') +
      ' /> 抓取完成后自动写入飞书（需已配置）</label>' +
      '<label class="ge-chk"><input type="checkbox" id="ge-buyin-delete-invalid-after-scan" ' +
      (deleteInvalidAfterScan ? 'checked' : '') +
      ' /> 抓取后尝试删除失效商品（逐条点删除）</label>' +
      '<label class="ge-chk"><input type="checkbox" id="ge-buyin-clear-after" ' +
      (clearAfter ? 'checked' : '') +
      ' /> 写入飞书成功后清空选品车</label>' +
      '<label class="ge-chk ge-chk-sub"><input type="checkbox" id="ge-buyin-clear-auto" ' +
      (clearAuto ? 'checked' : '') +
      ' /> 清空时不弹确认（直接全选并删当前页）</label>' +
      '<div class="ge-row">' +
      '<button id="ge-cart-delete-invalid" type="button">删除失效商品(页)</button>' +
      '<button id="ge-cart-page-clear" type="button">清空选品车(页)</button>' +
      '</div>' +
      '<div class="ge-row">' +
      '<button id="ge-cart-copy" type="button">复制JSON</button>' +
      '</div>' +
      '<div class="ge-row">' +
      '<button id="ge-cart-debug-copy" type="button">复制调试信息</button>' +
      '</div>' +
      '<div class="ge-row">' +
      '<button id="ge-cart-clear" type="button">清空缓存</button>' +
      '</div>';

    (document.body || document.documentElement).appendChild(panel);
    const byId = (id) => document.getElementById(id);
    const chkClear = byId('ge-buyin-clear-after');
    const chkClearAuto = byId('ge-buyin-clear-auto');
    function syncClearSubDisabled() {
      chkClearAuto.disabled = !chkClear.checked;
      if (!chkClear.checked) chkClearAuto.checked = false;
    }
    chkClear.addEventListener('change', function () {
      gmSet(GE_CLEAR_AFTER_FEISHU_KEY, chkClear.checked ? '1' : '0');
      if (!chkClear.checked) {
        gmSet(GE_CLEAR_AFTER_FEISHU_AUTO_KEY, '0');
        chkClearAuto.checked = false;
      }
      syncClearSubDisabled();
    });
    chkClearAuto.addEventListener('change', function () {
      gmSet(GE_CLEAR_AFTER_FEISHU_AUTO_KEY, chkClearAuto.checked ? '1' : '0');
    });
    syncClearSubDisabled();
    const chkAutoFeishuScan = byId('ge-buyin-auto-feishu-after-scan');
    if (chkAutoFeishuScan) {
      chkAutoFeishuScan.addEventListener('change', function () {
        gmSet(GE_AUTO_FEISHU_AFTER_SCAN_KEY, chkAutoFeishuScan.checked ? '1' : '0');
      });
    }
    const chkDelInvScan = byId('ge-buyin-delete-invalid-after-scan');
    if (chkDelInvScan) {
      chkDelInvScan.addEventListener('change', function () {
        gmSet(GE_DELETE_INVALID_AFTER_SCAN_KEY, chkDelInvScan.checked ? '1' : '0');
      });
    }
    byId('ge-cart-feishu-cfg').addEventListener('click', function () {
      showFeishuCartConfigModal();
    });
    byId('ge-cart-feishu-map').addEventListener('click', function () {
      showFeishuColumnMapOnlyModal();
    });
    byId('ge-cart-feishu-upload').addEventListener('click', async function () {
      const btn = byId('ge-cart-feishu-upload');
      const old = btn.textContent;
      btn.disabled = true;
      btn.textContent = '写入中…';
      try {
        await geUploadCartToFeishuWithClearAfterFlow({ showAlert: true });
      } catch (e) {
        const msg = e && e.message ? e.message : String(e);
        if (['暂无数据', '缺少飞书配置', '缺少鉴权'].indexOf(msg) < 0) {
          alert('写入飞书失败：' + msg);
        }
      } finally {
        btn.disabled = false;
        btn.textContent = old;
      }
    });
    byId('ge-cart-delete-invalid').addEventListener('click', async function () {
      if (
        !confirm(
          '将从当前列表由后往前逐条删除「商品ID与店铺ID均为空」的失效行，每条可能弹出确认。无法使用全选批量删。是否继续？'
        )
      ) {
        return;
      }
      const btn = byId('ge-cart-delete-invalid');
      const old = btn.textContent;
      btn.disabled = true;
      btn.textContent = '删除中…';
      try {
        const n = await geDeleteInvalidCartRowsDom();
        alert('处理完成，共触发删除 ' + n + ' 次（若当前无失效行则为 0）');
      } catch (e) {
        alert('失败：' + (e && e.message ? e.message : e));
      } finally {
        btn.disabled = false;
        btn.textContent = old;
      }
    });
    byId('ge-cart-page-clear').addEventListener('click', async function () {
      if (!confirm('将在当前页全选并批量删除选品车商品，是否继续？')) return;
      const btn = byId('ge-cart-page-clear');
      const old = btn.textContent;
      btn.disabled = true;
      btn.textContent = '处理中…';
      try {
        await clearBuyinCartDom();
        STORE.rows = [];
        STORE.keySet.clear();
        updateStatus();
        alert('已执行清空操作');
      } catch (e) {
        alert('失败：' + (e && e.message ? e.message : e));
      } finally {
        btn.disabled = false;
        btn.textContent = old;
      }
    });
    byId('ge-cart-scan').addEventListener('click', async function () {
      const btn = byId('ge-cart-scan');
      const stEl = byId('ge-cart-status');
      const old = btn.textContent;
      btn.disabled = true;
      btn.textContent = '执行中…';
      if (stEl) stEl.textContent = '执行中：请看下方蓝色进度行';
      let inc = 0;
      let feishuExtra = '';
      let deleteInvalidExtra = '';
      try {
        inc = await geScanAllByVisibleCards();
        const chkAutoFs = byId('ge-buyin-auto-feishu-after-scan');
        if (chkAutoFs && chkAutoFs.checked) {
          const pr = byId('ge-cart-progress');
          if (pr) {
            pr.textContent = '③写入飞书：正在上传多维表格…';
            pr.style.display = 'block';
          }
          try {
            const res = await geUploadCartToFeishuWithClearAfterFlow({ showAlert: false });
            feishuExtra = ' 并已写入飞书 ' + res.total + ' 条';
            if (res.columnsCreated) feishuExtra += '（新建列 ' + res.columnsCreated + ' 个）';
          } catch (e) {
            const msg = e && e.message ? e.message : String(e);
            alert(
              '抓取已完成，但自动写入飞书未成功：' +
                msg +
                '。请检查「飞书配置」后手动点「写入飞书」。'
            );
          }
        }
        const chkDelInv = byId('ge-buyin-delete-invalid-after-scan');
        if (chkDelInv && chkDelInv.checked) {
          if (
            confirm(
              '是否在页面上逐条删除「商品ID与店铺ID均为空」的失效行？（每条可能弹确认，无法批量勾选）'
            )
          ) {
            const pr2 = byId('ge-cart-progress');
            if (pr2) {
              pr2.textContent = '④删除失效商品：逐条操作中…';
              pr2.style.display = 'block';
            }
            try {
              const n = await geDeleteInvalidCartRowsDom();
              deleteInvalidExtra = ' 已逐条删除失效约 ' + n + ' 次';
            } catch (e) {
              alert('删除失效商品时出错：' + (e && e.message ? e.message : e));
            }
          }
        }
      } finally {
        btn.disabled = false;
        btn.textContent = old;
        updateStatus();
        geClearCartProgress();
      }
      alert(
        '抓取结束：本轮新增 ' +
          inc +
          ' 条；当前累计 ' +
          STORE.rows.length +
          ' 条。' +
          feishuExtra +
          deleteInvalidExtra +
          '\n若条数少于「全部商品」，请确认已滚到底并出现「没有更多了」。'
      );
    });
    byId('ge-cart-export').addEventListener('click', function () {
      downloadCsv();
    });
    byId('ge-cart-copy').addEventListener('click', async function () {
      const validRows = STORE.rows.filter(geIsValidCartExportRow);
      if (!validRows.length) {
        alert('暂无有效数据可复制（商品ID与店铺ID不能同时为空）');
        return;
      }
      validRows.forEach(fillProductLinkFromId);
      try {
        await navigator.clipboard.writeText(JSON.stringify(validRows, null, 2));
        alert('已复制 ' + validRows.length + ' 条有效 JSON');
      } catch (e) {
        alert('复制失败：' + (e && e.message ? e.message : e));
      }
    });
    byId('ge-cart-debug-copy').addEventListener('click', async function () {
      try {
        const text = JSON.stringify(geBuildCartDebugPayload(), null, 2);
        await navigator.clipboard.writeText(text);
        alert('已复制调试信息（含 STORE.rows / keySet / 最近扫描摘要）');
      } catch (e) {
        alert('复制失败：' + (e && e.message ? e.message : e));
      }
    });
    byId('ge-cart-clear').addEventListener('click', function () {
      if (!confirm('仅清空脚本侧已抓取缓存，不删网页选品车。继续？')) return;
      STORE.rows = [];
      STORE.keySet.clear();
      updateStatus();
    });
    updateStatus();
  }

  function ensureStyle() {
    const css = `
      #ge-cart-panel{
        position:fixed;right:16px;top:110px;z-index:2147483647;
        width:260px;background:#fff;border:1px solid #e5e6eb;border-radius:10px;
        box-shadow:0 8px 24px rgba(0,0,0,.12);padding:10px;font:12px/1.5 -apple-system,BlinkMacSystemFont,sans-serif;
        overflow:visible;max-height:none;
        pointer-events:none;
      }
      #ge-cart-panel button,
      #ge-cart-panel input,
      #ge-cart-panel label,
      #ge-cart-panel summary,
      #ge-cart-panel details,
      #ge-cart-panel select,
      #ge-cart-panel textarea{
        pointer-events:auto;
      }
      #ge-cart-panel .ge-hd{font-weight:700;color:#1f2329;margin-bottom:4px}
      #ge-cart-panel .ge-tip{margin:0 0 8px;font-size:11px;color:#86909c;line-height:1.45}
      #ge-cart-panel .ge-fields{margin:0 0 8px;font-size:11px;color:#4e5969}
      #ge-cart-panel .ge-fields summary{cursor:pointer;user-select:none;color:#3370ff}
      #ge-cart-panel .ge-fields-ul{margin:6px 0 0;padding-left:18px;line-height:1.5}
      #ge-cart-panel .ge-st{color:#4e5969;margin-bottom:4px}
      #ge-cart-panel .ge-progress{display:none;font-size:11px;color:#3370ff;line-height:1.5;margin-bottom:8px;padding:6px 8px;background:#f0f6ff;border-radius:6px;border:1px solid #bedaff;word-break:break-word}
      #ge-cart-panel .ge-chk-sub{margin-left:8px;padding-left:4px;opacity:.95}
      #ge-cart-panel .ge-row{display:flex;gap:6px;margin-bottom:6px}
      #ge-cart-panel .ge-chk{display:flex;align-items:center;gap:6px;margin:0 0 8px;font-size:11px;color:#4e5969;cursor:pointer}
      #ge-cart-panel button{
        flex:1;cursor:pointer;border:1px solid #dcdfe6;border-radius:6px;background:#f7f8fa;
        color:#1f2329;padding:5px 4px;font-size:12px
      }
      #ge-cart-panel .ge-btn-primary{background:#3370ff;color:#fff;border-color:#3370ff}
    `;
    if (typeof GM_addStyle === 'function') GM_addStyle(css);
    else {
      const st = document.createElement('style');
      st.textContent = css;
      (document.head || document.documentElement).appendChild(st);
    }
  }

  function boot() {
    geInstallCartReloadGuardOnce();
    ensureStyle();
    const timer = setInterval(function () {
      if (!document.body) return;
      mountPanel();
      if (STORE.mounted) clearInterval(timer);
    }, 250);
  }

  boot();
})();
