/* =====================================================================
   presentations.js — Presentation Repository for the AVD-II (B) dashboard
   ---------------------------------------------------------------------
   Loaded on demand (the first time the "Presentations" tab is opened),
   so it adds nothing to the dashboard's normal load time.

   HOW IT WORKS
   - The original .pptx is stored untouched in Firestore, in ~800 KB
     pieces (Presentations/{id}/chunks). No Firebase Storage needed.
   - On upload, every slide's text, tables and pictures are read out and
     saved as one small document per slide (Presentations/{id}/slides).
   - Officers see every slide with a "Changes suggested" column beside
     it (Presentations/{id}/suggestions), and can edit the text of any
     slide live. Edits are saved per slide; the original text is kept
     so any slide can be reverted.
   - "Download edited PPTX" re-opens the ORIGINAL file and applies only
     the text edits (plus slide order / removal) to it, so the original
     design, fonts, layouts, charts and animations are preserved.

   KNOWN LIMITS (by design)
   - The on-screen view shows content (text, tables, pictures), not the
     exact slide design. The downloaded file keeps the real design.
   - Text inside charts, SmartArt and grouped images is not editable.
   - New slides cannot be added here (use "Changes suggested" for that).
   ===================================================================== */
(function () {
  "use strict";
  if (window.PresModule) return;

  const COL = "Presentations";
  const CHUNK = 800 * 1024;
  const MAX_FILE = 25 * 1024 * 1024;
  const MAX_SLIDES = 400;
  const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const MIME_PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  const NAME_KEY = "avd2_officer_name";
  const JSZIP_URL = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";

  const state = { view: "list", list: null, deck: null, editing: {}, name: "" };
  try { state.name = localStorage.getItem(NAME_KEY) || ""; } catch (e) {}

  /* ------------------------------ helpers ------------------------------ */
  const esc = s => escapeHtml(s == null ? "" : s);
  const pad = (n, w) => String(n).padStart(w, "0");
  const chunkId = i => pad(i, 4);
  const slideKey = i => "s" + pad(i, 3);
  const nowIso = () => new Date().toISOString();
  const fmtDate = iso => { const d = new Date(iso); return isNaN(d) ? "—" : d.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" }); };
  const fmtDateTime = iso => { const d = new Date(iso); return isNaN(d) ? "—" : d.toLocaleString("en-IN", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }); };
  const fmtSize = n => n >= 1048576 ? (n / 1048576).toFixed(1) + " MB" : Math.max(1, Math.round(n / 1024)) + " KB";
  const clone = o => JSON.parse(JSON.stringify(o));
  const explain = e => (e && e.code === "permission-denied")
    ? "Firestore rules do not allow Presentations yet — publish the updated firestore.rules (it adds the Presentations collection)."
    : ((e && e.message) || String(e));

  async function runLimited(tasks, n) {
    let i = 0;
    const workers = Array.from({ length: Math.min(n, tasks.length) }, async () => {
      while (i < tasks.length) { const k = i++; await tasks[k](); }
    });
    await Promise.all(workers);
  }
  function loadScript(src) {
    return new Promise((res, rej) => {
      const s = document.createElement("script");
      s.src = src; s.onload = res; s.onerror = () => rej(new Error("Could not load " + src));
      document.head.appendChild(s);
    });
  }
  let zipPromise = null;
  function ensureJSZip() {
    if (window.JSZip) return Promise.resolve();
    if (!zipPromise) zipPromise = loadScript(JSZIP_URL).catch(e => { zipPromise = null; throw e; });
    return zipPromise;
  }

  /* ------------------------------ XML helpers ------------------------------ */
  function parseXml(s) {
    const d = new DOMParser().parseFromString(s, "application/xml");
    if (d.getElementsByTagName("parsererror").length) throw new Error("This file has XML the portal cannot read.");
    return d;
  }
  function serializeXml(d) {
    let s = new XMLSerializer().serializeToString(d);
    if (!/^<\?xml/.test(s)) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s;
    return s;
  }
  const kids = (el, ln) => Array.from(el.childNodes).filter(n => n.nodeType === 1 && (!ln || n.localName === ln));
  const allByName = (doc, ln) => Array.from(doc.getElementsByTagNameNS("*", ln));
  function resolvePath(baseDir, target) {
    if (target[0] === "/") return target.slice(1);
    const out = [];
    (baseDir + target).split("/").forEach(p => { if (p === "..") out.pop(); else if (p && p !== ".") out.push(p); });
    return out.join("/");
  }
  async function readRels(zip, path) {
    const out = {};
    const f = zip.file(path);
    if (!f) return out;
    const doc = parseXml(await f.async("string"));
    kids(doc.documentElement, "Relationship").forEach(r => {
      out[r.getAttribute("Id")] = { type: r.getAttribute("Type") || "", target: r.getAttribute("Target") || "" };
    });
    return out;
  }

  /* ------------------------------ PPTX parsing ------------------------------ */
  function paraText(p) {
    let s = "";
    kids(p).forEach(n => {
      if (n.localName === "r" || n.localName === "fld") s += kids(n, "t").map(t => t.textContent).join("");
      else if (n.localName === "br") s += "\n";
    });
    return s;
  }
  function lvlOf(p) {
    const pPr = kids(p, "pPr")[0];
    return pPr ? (parseInt(pPr.getAttribute("lvl") || "0", 10) || 0) : 0;
  }
  function deriveTitle(blocks) {
    const t = blocks.find(b => b.kind === "title") || blocks[0];
    if (!t) return "(no text)";
    const line = ((t.paras.find(p => p.t.trim()) || { t: "" }).t).replace(/\s+/g, " ").trim();
    return line ? (line.length > 90 ? line.slice(0, 88) + "…" : line) : "(no text)";
  }
  function makeThumb(blob) {
    return new Promise(resolve => {
      const url = URL.createObjectURL(blob);
      const img = new Image();
      img.onload = () => {
        try {
          let maxW = 720, q = 0.65, out = "";
          for (let i = 0; i < 5; i++) {
            const sc = Math.min(1, maxW / img.naturalWidth);
            const w = Math.max(1, Math.round(img.naturalWidth * sc)), h = Math.max(1, Math.round(img.naturalHeight * sc));
            const c = document.createElement("canvas"); c.width = w; c.height = h;
            const x = c.getContext("2d"); x.fillStyle = "#fff"; x.fillRect(0, 0, w, h); x.drawImage(img, 0, 0, w, h);
            out = c.toDataURL("image/jpeg", q);
            if (out.length <= 140000) break;
            maxW *= 0.8; q -= 0.08;
          }
          resolve(out);
        } catch (e) { resolve(""); }
        URL.revokeObjectURL(url);
      };
      img.onerror = () => { URL.revokeObjectURL(url); resolve(""); };
      img.src = url;
    });
  }
  const IMG_MIME = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", webp: "image/webp" };

  async function parseSlide(zip, part, xml, idx) {
    const doc = parseXml(xml);
    const tbls = allByName(doc, "tbl");
    const blocks = [];
    allByName(doc, "txBody").forEach((tx, bi) => {
      const paras = kids(tx, "p").map((p, src) => ({ t: paraText(p), l: lvlOf(p), src }));
      if (!paras.some(p => p.t.trim())) return;
      const holder = tx.parentNode;
      let kind = "text", label = "Text", tbl = null;
      if (holder && holder.localName === "tc") {
        const tr = holder.parentNode, tb = tr.parentNode;
        kind = "cell"; label = "Table cell";
        tbl = {
          k: tbls.indexOf(tb),
          r: kids(tb, "tr").indexOf(tr),
          c: kids(tr, "tc").indexOf(holder)
        };
      } else if (holder && holder.localName === "sp") {
        const nv = kids(holder, "nvSpPr")[0];
        const cNv = nv && kids(nv, "cNvPr")[0];
        label = (cNv && cNv.getAttribute("name")) || "Text";
        const nvPr = nv && kids(nv, "nvPr")[0];
        const ph = nvPr && kids(nvPr, "ph")[0];
        if (ph) {
          const ty = ph.getAttribute("type") || "body";
          if (ty === "sldNum" || ty === "dt" || ty === "ftr") return;
          if (ty === "title" || ty === "ctrTitle") kind = "title";
          else if (ty === "subTitle") kind = "subtitle";
        }
      }
      blocks.push({ bi, kind, label, tbl, paras });
    });

    // Pictures (small previews only — the real pictures stay in the original file)
    const imgs = [];
    const rels = await readRels(zip, part.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels");
    const embeds = [];
    allByName(doc, "blip").forEach(b => { const id = b.getAttributeNS(NS_R, "embed"); if (id && !embeds.includes(id)) embeds.push(id); });
    let budget = 600000;
    for (const id of embeds) {
      if (imgs.length >= 4) break;
      const rel = rels[id];
      if (!rel || !/\/image$/.test(rel.type)) continue;
      const path = resolvePath("ppt/slides/", rel.target);
      const ext = path.split(".").pop().toLowerCase();
      if (!IMG_MIME[ext]) continue;
      const f = zip.file(path);
      if (!f) continue;
      const thumb = await makeThumb(new Blob([await f.async("uint8array")], { type: IMG_MIME[ext] }));
      if (thumb && thumb.length <= budget) { imgs.push(thumb); budget -= thumb.length; }
    }

    return {
      order: idx, part, title: deriveTitle(blocks), blocks,
      orig: blocks.map(b => ({ bi: b.bi, paras: clone(b.paras) })),
      imgs, deleted: false, edited: false, editedBy: "", editedAt: ""
    };
  }

  async function parseDeck(zip, onProgress) {
    const presFile = zip.file("ppt/presentation.xml");
    if (!presFile) throw new Error("This does not look like a PowerPoint (.pptx) file.");
    const pdoc = parseXml(await presFile.async("string"));
    const rels = await readRels(zip, "ppt/_rels/presentation.xml.rels");
    const lst = kids(pdoc.documentElement, "sldIdLst")[0];
    const entries = lst ? kids(lst, "sldId") : [];
    const slides = [];
    for (let i = 0; i < entries.length; i++) {
      const rel = rels[entries[i].getAttributeNS(NS_R, "id")];
      if (!rel) continue;
      const part = resolvePath("ppt/", rel.target);
      const f = zip.file(part);
      if (!f) continue;
      slides.push(await parseSlide(zip, part, await f.async("string"), slides.length));
      if (onProgress) onProgress(i + 1, entries.length);
    }
    return slides;
  }

  /* ------------------------------ Firestore I/O ------------------------------ */
  const deckRef = (id) => db.collection(COL).doc(id || state.deck.id);
  const slideRef = sid => deckRef().collection("slides").doc(sid);

  async function loadList() {
    const snap = await db.collection(COL).get();
    state.list = snap.docs.map(d => ({ id: d.id, ...d.data() }))
      .sort((a, b) => (b.uploadedAt || "").localeCompare(a.uploadedAt || ""));
  }

  async function uploadDeck({ file, title, description, by }, progress) {
    await ensureJSZip();
    progress("Reading the file…", 0.02);
    const buf = new Uint8Array(await file.arrayBuffer());
    let zip;
    try { zip = await JSZip.loadAsync(buf); }
    catch (e) { throw new Error("Could not open this file. It may be password-protected or not a valid .pptx."); }
    const slides = await parseDeck(zip, (i, n) => progress(`Reading slide ${i} of ${n}…`, 0.02 + 0.2 * (i / n)));
    if (!slides.length) throw new Error("No slides were found in this file.");
    if (slides.length > MAX_SLIDES) throw new Error("Too many slides (limit " + MAX_SLIDES + ").");

    const ref = db.collection(COL).doc();
    const nChunks = Math.ceil(buf.length / CHUNK);
    const total = nChunks + slides.length;
    let done = 0;
    const tick = () => { done++; progress(`Saving… ${done} of ${total}`, 0.25 + 0.7 * (done / total)); };
    try {
      await runLimited(Array.from({ length: nChunks }, (_, i) => () =>
        ref.collection("chunks").doc(chunkId(i)).set({
          i, data: firebase.firestore.Blob.fromUint8Array(buf.slice(i * CHUNK, (i + 1) * CHUNK))
        }).then(tick)), 3);
      await runLimited(slides.map((s, k) => () => ref.collection("slides").doc(slideKey(k)).set(s).then(tick)), 4);
      const meta = {
        title, description, fileName: file.name, fileSize: buf.length, chunkCount: nChunks,
        slideCount: slides.length, uploadedBy: by, uploadedAt: nowIso(), updatedAt: nowIso(),
        lastEditedBy: "", sugCount: 0
      };
      await ref.set(meta);
      return { id: ref.id, ...meta };
    } catch (e) {
      // best-effort cleanup of a half-finished upload
      try { await deleteSubdocs(ref, nChunks, slides.length, []); } catch (_) {}
      throw e;
    }
  }

  async function deleteSubdocs(ref, nChunks, nSlides, sugIds) {
    const tasks = [];
    for (let i = 0; i < nChunks; i++) tasks.push(() => ref.collection("chunks").doc(chunkId(i)).delete());
    for (let i = 0; i < nSlides; i++) tasks.push(() => ref.collection("slides").doc(slideKey(i)).delete());
    sugIds.forEach(id => tasks.push(() => ref.collection("suggestions").doc(id).delete()));
    await runLimited(tasks, 6);
  }

  async function deleteDeck(id) {
    const meta = state.list.find(x => x.id === id);
    if (!meta) return;
    if (!confirm('Delete "' + meta.title + '" with all its slides, suggestions and the stored file? This cannot be undone.')) return;
    showToast("Deleting…");
    try {
      const ref = deckRef(id);
      const sug = await ref.collection("suggestions").get();
      await deleteSubdocs(ref, meta.chunkCount || 0, meta.slideCount || 0, sug.docs.map(d => d.id));
      await ref.delete();
      state.list = state.list.filter(x => x.id !== id);
      showToast("Presentation deleted");
      renderList();
    } catch (e) { showToast("Delete failed — " + explain(e)); }
  }

  async function loadDeckData() {
    const ref = deckRef();
    const [ss, sg] = await Promise.all([ref.collection("slides").get(), ref.collection("suggestions").get()]);
    state.deck.slides = ss.docs.map(d => ({ id: d.id, ...d.data() })).sort((a, b) => a.order - b.order);
    state.deck.sug = sg.docs.map(d => ({ id: d.id, ...d.data() }));
    const open = state.deck.sug.filter(x => !x.done).length;
    if ((state.deck.meta.sugCount || 0) !== open) {
      state.deck.meta.sugCount = open;
      ref.set({ sugCount: open }, { merge: true }).catch(() => {});
    }
  }

  async function openDeck(id) {
    const meta = state.list.find(x => x.id === id);
    if (!meta) return;
    state.view = "deck";
    state.deck = { id, meta, slides: null, sug: null };
    state.editing = {};
    render();
    try {
      await loadDeckData();
    } catch (e) {
      showToast("Could not open — " + explain(e));
      goList();
      return;
    }
    if (state.deck && state.deck.id === id) render();
  }

  function goList() {
    state.view = "list"; state.deck = null; state.editing = {}; state.list = null;
    if (typeof currentTab !== "undefined" && currentTab === "presentations") render();
  }

  /* ------------------------------ editing ------------------------------ */
  const sameParas = (a, b) => a.length === b.length && a.every((p, i) => p.t === b[i].t && p.src === b[i].src);
  const getSlide = sid => state.deck.slides.find(s => s.id === sid);
  const sortedSlides = () => state.deck.slides.slice().sort((a, b) => a.order - b.order);
  function needName() {
    if (state.name.trim()) return true;
    showToast("Please type your name at the top first");
    const el = document.getElementById("presName");
    if (el) el.focus();
    return false;
  }
  async function touchDeck() {
    state.deck.meta.updatedAt = nowIso();
    state.deck.meta.lastEditedBy = state.name;
    await deckRef().set({ updatedAt: state.deck.meta.updatedAt, lastEditedBy: state.name }, { merge: true });
  }

  async function saveSlide(sid) {
    if (!needName()) return;
    const s = getSlide(sid), draft = state.editing[sid];
    if (!s || !draft) return;
    const blocks = s.blocks.map(b => {
      const d = draft.find(x => x.bi === b.bi);
      return { ...b, paras: d.paras.map(p => ({ t: p.t, l: p.l, src: p.src })) };
    });
    const edited = blocks.some(b => !sameParas(b.paras, s.orig.find(o => o.bi === b.bi).paras));
    const patch = { blocks, edited, editedBy: edited ? state.name : "", editedAt: edited ? nowIso() : "", title: deriveTitle(blocks) };
    try {
      await slideRef(sid).set(patch, { merge: true });
      Object.assign(s, patch);
      delete state.editing[sid];
      if (edited) await touchDeck();
      refreshSlide(sid);
      showToast(edited ? "Slide saved" : "No changes to save");
    } catch (e) { showToast("Save failed — " + explain(e)); }
  }

  async function revertSlide(sid) {
    if (!needName()) return;
    const s = getSlide(sid);
    if (!s || !confirm("Revert this slide's text to the original?")) return;
    const blocks = s.blocks.map(b => ({ ...b, paras: clone(s.orig.find(o => o.bi === b.bi).paras) }));
    const patch = { blocks, edited: false, editedBy: "", editedAt: "", title: deriveTitle(blocks) };
    try {
      await slideRef(sid).set(patch, { merge: true });
      Object.assign(s, patch);
      delete state.editing[sid];
      await touchDeck();
      refreshSlide(sid);
      showToast("Slide reverted to original");
    } catch (e) { showToast("Revert failed — " + explain(e)); }
  }

  async function toggleRemoved(sid) {
    if (!needName()) return;
    const s = getSlide(sid);
    try {
      await slideRef(sid).set({ deleted: !s.deleted }, { merge: true });
      s.deleted = !s.deleted;
      await touchDeck();
      refreshSlide(sid);
      showToast(s.deleted ? "Slide marked for removal" : "Slide restored");
    } catch (e) { showToast("Failed — " + explain(e)); }
  }

  async function moveSlide(sid, dir) {
    if (!needName()) return;
    const arr = sortedSlides();
    const i = arr.findIndex(s => s.id === sid), j = i + dir;
    if (j < 0 || j >= arr.length) return;
    const a = arr[i], b = arr[j], oa = a.order, ob = b.order;
    try {
      await Promise.all([slideRef(a.id).set({ order: ob }, { merge: true }), slideRef(b.id).set({ order: oa }, { merge: true })]);
      a.order = ob; b.order = oa;
      await touchDeck();
      renderDeck();
      showToast("Slide moved");
    } catch (e) { showToast("Move failed — " + explain(e)); }
  }

  /* ------------------------------ suggestions ------------------------------ */
  async function addSuggestion(sid) {
    const row = document.getElementById("pslide-" + sid);
    const ta = row && row.querySelector("[data-sug-input]");
    const text = ta ? ta.value.trim() : "";
    if (!text) { showToast("Type your suggestion first"); return; }
    if (!needName()) return;
    const rec = { slideId: sid, text, by: state.name.trim(), at: nowIso(), done: false };
    try {
      const r = await deckRef().collection("suggestions").add(rec);
      state.deck.sug.push({ id: r.id, ...rec });
      await deckRef().set({ sugCount: firebase.firestore.FieldValue.increment(1) }, { merge: true });
      state.deck.meta.sugCount = (state.deck.meta.sugCount || 0) + 1;
      ta.value = "";   // clear the box, otherwise refreshSlide would restore the old text
      refreshSlide(sid);
      updateStats();
      showToast("Suggestion added");
    } catch (e) { showToast("Could not add — " + explain(e)); }
  }
  async function toggleSuggestion(sugId) {
    if (!needName()) return;
    const x = state.deck.sug.find(s => s.id === sugId);
    if (!x) return;
    const nowDone = !x.done;
    const patch = { done: nowDone, doneBy: nowDone ? state.name.trim() : "" };
    try {
      await deckRef().collection("suggestions").doc(sugId).set(patch, { merge: true });
      Object.assign(x, patch);
      await deckRef().set({ sugCount: firebase.firestore.FieldValue.increment(nowDone ? -1 : 1) }, { merge: true });
      state.deck.meta.sugCount = Math.max(0, (state.deck.meta.sugCount || 0) + (nowDone ? -1 : 1));
      refreshSlide(x.slideId);
      updateStats();
    } catch (e) { showToast("Failed — " + explain(e)); }
  }
  async function deleteSuggestion(sugId) {
    const x = state.deck.sug.find(s => s.id === sugId);
    if (!x || !confirm("Delete this suggestion?")) return;
    try {
      await deckRef().collection("suggestions").doc(sugId).delete();
      state.deck.sug = state.deck.sug.filter(s => s.id !== sugId);
      if (!x.done) {
        await deckRef().set({ sugCount: firebase.firestore.FieldValue.increment(-1) }, { merge: true });
        state.deck.meta.sugCount = Math.max(0, (state.deck.meta.sugCount || 0) - 1);
      }
      refreshSlide(x.slideId);
      updateStats();
    } catch (e) { showToast("Failed — " + explain(e)); }
  }

  /* ------------------------------ download ------------------------------ */
  async function fetchBytes() {
    const snap = await deckRef().collection("chunks").get();
    const parts = snap.docs.map(d => d.data()).sort((a, b) => a.i - b.i).map(x => x.data.toUint8Array());
    if (parts.length !== state.deck.meta.chunkCount) throw new Error("Stored file is incomplete (" + parts.length + " of " + state.deck.meta.chunkCount + " parts).");
    const out = new Uint8Array(parts.reduce((s, p) => s + p.length, 0));
    let off = 0;
    parts.forEach(p => { out.set(p, off); off += p.length; });
    return out;
  }

  function setParaText(p, text) {
    const doc = p.ownerDocument;
    const els = Array.from(p.childNodes).filter(n => n.nodeType === 1);
    const firstRun = els.find(n => n.localName === "r");
    let rPrTpl = null;
    if (firstRun) {
      const rp = kids(firstRun, "rPr")[0];
      if (rp) rPrTpl = rp.cloneNode(true);
    } else {
      const end = els.find(n => n.localName === "endParaRPr");
      if (end) {
        rPrTpl = doc.createElementNS(NS_A, "a:rPr");
        Array.from(end.attributes).forEach(a => rPrTpl.setAttribute(a.name, a.value));
        Array.from(end.childNodes).forEach(c => rPrTpl.appendChild(c.cloneNode(true)));
      }
    }
    els.forEach(n => { if (n.localName === "r" || n.localName === "br" || n.localName === "fld") p.removeChild(n); });
    if (!text) return;
    const endPr = Array.from(p.childNodes).find(n => n.nodeType === 1 && n.localName === "endParaRPr") || null;
    String(text).split("\n").forEach((seg, i) => {
      if (i > 0) {
        const br = doc.createElementNS(NS_A, "a:br");
        if (rPrTpl) br.appendChild(rPrTpl.cloneNode(true));
        p.insertBefore(br, endPr);
      }
      const r = doc.createElementNS(NS_A, "a:r");
      if (rPrTpl) r.appendChild(rPrTpl.cloneNode(true));
      const t = doc.createElementNS(NS_A, "a:t");
      t.textContent = seg;
      r.appendChild(t);
      p.insertBefore(r, endPr);
    });
  }
  function rewriteTxBody(tx, paras) {
    const olds = kids(tx, "p");
    if (!olds.length) return;
    const news = (paras.length ? paras : [{ t: "", src: 0 }]).map(p => {
      const np = olds[Math.min(Math.max(p.src || 0, 0), olds.length - 1)].cloneNode(true);
      setParaText(np, p.t);
      return np;
    });
    olds.forEach(o => tx.removeChild(o));
    news.forEach(n => tx.appendChild(n));
  }

  async function applyStructure(zip, slides) {
    const presPath = "ppt/presentation.xml", relsPath = "ppt/_rels/presentation.xml.rels";
    const pdoc = parseXml(await zip.file(presPath).async("string"));
    const rdoc = parseXml(await zip.file(relsPath).async("string"));
    const relEls = kids(rdoc.documentElement, "Relationship");
    const relByPart = {};
    relEls.forEach(r => { relByPart[resolvePath("ppt/", r.getAttribute("Target"))] = r; });
    const lst = kids(pdoc.documentElement, "sldIdLst")[0];
    if (!lst) return;
    const entries = kids(lst, "sldId");
    const entryByRid = {};
    entries.forEach(e => { entryByRid[e.getAttributeNS(NS_R, "id")] = e; });

    const keep = [], drop = [];
    slides.forEach(s => {
      const rel = relByPart[s.part];
      const e = rel && entryByRid[rel.getAttribute("Id")];
      if (!e) return;
      (s.deleted ? drop : keep).push({ s, e, rel });
    });
    const keptNow = entries.filter(e => keep.some(k => k.e === e));
    const keptNew = keep.map(k => k.e);
    const orderChanged = keptNow.some((e, i) => e !== keptNew[i]);
    if (!drop.length && !orderChanged) return;

    const untouched = entries.filter(e => !keep.some(k => k.e === e) && !drop.some(d => d.e === e));
    entries.forEach(e => lst.removeChild(e));
    keptNew.concat(untouched).forEach(e => lst.appendChild(e));

    // Slide sections and custom shows point at slide ids; drop them rather than leave stale ids.
    allByName(pdoc, "sectionLst").forEach(el => {
      const ext = el.parentNode;
      if (ext && ext.localName === "ext" && ext.parentNode) ext.parentNode.removeChild(ext);
      else if (ext) ext.removeChild(el);
    });
    if (drop.length) allByName(pdoc, "custShowLst").forEach(el => el.parentNode.removeChild(el));

    const removedParts = [];
    for (const d of drop) {
      const sRelsPath = d.s.part.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
      const srels = await readRels(zip, sRelsPath);
      for (const r of Object.values(srels)) {
        if (/\/notesSlide$/.test(r.type)) {
          const np = resolvePath("ppt/slides/", r.target);
          removedParts.push(np);
          zip.remove(np);
          zip.remove(np.replace("ppt/notesSlides/", "ppt/notesSlides/_rels/") + ".rels");
        }
      }
      removedParts.push(d.s.part);
      zip.remove(d.s.part);
      zip.remove(sRelsPath);
      d.rel.parentNode.removeChild(d.rel);
    }
    if (removedParts.length) {
      const ctPath = "[Content_Types].xml";
      const ctDoc = parseXml(await zip.file(ctPath).async("string"));
      kids(ctDoc.documentElement, "Override").forEach(o => {
        if (removedParts.includes((o.getAttribute("PartName") || "").replace(/^\//, ""))) ctDoc.documentElement.removeChild(o);
      });
      zip.file(ctPath, serializeXml(ctDoc));
    }
    zip.file(presPath, serializeXml(pdoc));
    zip.file(relsPath, serializeXml(rdoc));
  }

  async function buildEdited() {
    await ensureJSZip();
    const ss = await deckRef().collection("slides").get();
    state.deck.slides = ss.docs.map(d => ({ id: d.id, ...d.data() })).sort((a, b) => a.order - b.order);
    const zip = await JSZip.loadAsync(await fetchBytes());
    for (const s of state.deck.slides) {
      if (s.deleted) continue;
      const changed = s.blocks.filter(b => { const o = s.orig.find(x => x.bi === b.bi); return !o || !sameParas(b.paras, o.paras); });
      if (!changed.length) continue;
      const f = zip.file(s.part);
      if (!f) continue;
      const doc = parseXml(await f.async("string"));
      const txs = allByName(doc, "txBody");
      changed.forEach(b => { if (txs[b.bi]) rewriteTxBody(txs[b.bi], b.paras); });
      zip.file(s.part, serializeXml(doc));
    }
    await applyStructure(zip, state.deck.slides);
    return zip.generateAsync({ type: "blob", mimeType: MIME_PPTX, compression: "DEFLATE" });
  }
  function saveBlob(blob, name) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
  }
  const safeName = s => String(s || "presentation").replace(/[\\/:*?"<>|]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 80);

  async function withBusy(btn, label, fn) {
    const old = btn.textContent;
    btn.disabled = true; btn.textContent = label;
    try { await fn(); } catch (e) { showToast("Failed — " + explain(e)); }
    finally { btn.disabled = false; btn.textContent = old; }
  }

  function csvCell(v) { return '"' + String(v == null ? "" : v).replace(/"/g, '""') + '"'; }
  function exportSuggestions() {
    const slides = sortedSlides();
    const rows = [["Slide No.", "Slide title", "Suggestion", "Suggested by", "Date", "Status"]];
    slides.forEach((s, i) => {
      state.deck.sug.filter(x => x.slideId === s.id).sort((a, b) => (a.at || "").localeCompare(b.at || "")).forEach(x => {
        rows.push([i + 1, s.title, x.text, x.by, fmtDateTime(x.at), x.done ? "Done (" + (x.doneBy || "") + ")" : "Open"]);
      });
    });
    if (rows.length === 1) { showToast("No suggestions to export yet"); return; }
    const csv = "\ufeff" + rows.map(r => r.map(csvCell).join(",")).join("\r\n");
    saveBlob(new Blob([csv], { type: "text/csv;charset=utf-8" }), safeName(state.deck.meta.title) + " - suggestions.csv");
  }

  /* ------------------------------ rendering ------------------------------ */
  const ICON = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="12" rx="2"/><path d="M12 16v4M8 20h8"/><path d="M7 9h5M7 12h8"/></svg>';

  function injectStyles() {
    if (document.getElementById("presStyles")) return;
    const st = document.createElement("style");
    st.id = "presStyles";
    st.textContent = `
.pres-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:14px;}
.pres-card{background:var(--paper-raised);border:1px solid var(--border);border-radius:var(--radius-md);padding:16px 18px;box-shadow:var(--shadow-sm);cursor:pointer;display:flex;flex-direction:column;gap:8px;transition:box-shadow .15s ease,transform .15s ease,border-color .15s ease;min-width:0;}
.pres-card:hover{box-shadow:var(--shadow-md);transform:translateY(-2px);border-color:var(--border-strong);}
.pres-card-icon{width:34px;height:34px;border-radius:9px;display:flex;align-items:center;justify-content:center;background:var(--blue-soft);color:var(--blue);}
.pres-card-icon svg{width:17px;height:17px;}
.pres-card h3{font-family:var(--serif);font-size:17px;margin:0;color:var(--ink);word-break:break-word;}
.pres-desc{font-size:12.5px;color:var(--slate);line-height:1.45;white-space:pre-wrap;word-break:break-word;}
.pres-meta{font-size:11.5px;color:var(--slate-light);line-height:1.5;}
.pres-card-actions{display:flex;gap:8px;margin-top:4px;}
.pres-chip{display:inline-block;font-family:var(--mono);font-size:10.5px;font-weight:600;padding:3px 8px;border-radius:20px;background:var(--border);color:var(--slate);margin-right:4px;}
.pres-chip.warn{background:var(--amber-soft);color:var(--amber);}
.pres-chip.ok{background:var(--sage-soft);color:var(--sage);}
.pres-chip.bad{background:var(--rust-soft);color:var(--rust);}
.pres-bar{display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap;background:var(--paper-raised);border:1px solid var(--border);border-radius:var(--radius-md);padding:14px 16px;margin-bottom:12px;box-shadow:var(--shadow-sm);}
.pres-name{display:flex;flex-direction:column;gap:4px;flex:1;min-width:170px;max-width:260px;}
.pres-name label{font-size:11px;font-weight:600;color:var(--ink-soft);}
.pres-name input{padding:8px 10px;border:1px solid var(--border-strong);border-radius:var(--radius-sm);font-family:var(--sans);font-size:13px;background:var(--paper-raised);color:var(--slate);}
.pres-note{font-size:12px;color:var(--slate-light);line-height:1.55;margin:0 2px 8px;}
.pres-stats{font-size:12px;color:var(--slate);margin:0 2px 16px;}
.pres-slide{display:grid;grid-template-columns:minmax(0,1.7fr) minmax(0,1fr);border:1px solid var(--border);border-radius:var(--radius-md);background:var(--paper-raised);margin-bottom:16px;box-shadow:var(--shadow-sm);overflow:hidden;}
.pres-slide.removed .pres-slide-body{opacity:.5;text-decoration:line-through;}
.pres-slide-main{padding:16px 18px;border-right:1px solid var(--border);min-width:0;}
.pres-slide-head{display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin-bottom:12px;}
.pres-slide-no{font-family:var(--mono);font-size:11.5px;font-weight:700;background:var(--border);color:var(--ink);padding:3px 9px;border-radius:6px;}
.pres-spacer{flex:1;}
.pres-title{font-family:var(--serif);font-size:20px;font-weight:600;color:var(--ink);margin:0 0 8px;white-space:pre-wrap;word-break:break-word;}
.pres-subtitle{font-size:14px;color:var(--slate-light);margin:0 0 8px;white-space:pre-wrap;}
.pres-block{margin-bottom:12px;}
.pres-label{font-family:var(--mono);font-size:9.5px;letter-spacing:.06em;text-transform:uppercase;color:var(--slate-light);margin-bottom:4px;}
.pres-p{font-size:13.5px;line-height:1.55;color:var(--slate);white-space:pre-wrap;word-break:break-word;margin-bottom:3px;}
.pres-bul{display:inline-block;width:14px;color:var(--slate-light);}
table.pres-table{border-collapse:collapse;width:100%;margin:6px 0 12px;font-size:12.5px;}
table.pres-table td{border:1px solid var(--border-strong);padding:6px 8px;vertical-align:top;}
.pres-imgs{display:flex;gap:8px;flex-wrap:wrap;margin-top:8px;}
.pres-imgs img{max-width:100%;max-height:200px;border-radius:6px;border:1px solid var(--border);}
.pres-empty-slide{font-size:12.5px;color:var(--slate-light);font-style:italic;}
.pres-prow{display:flex;gap:6px;align-items:flex-start;margin-bottom:5px;}
.pres-prow textarea{flex:1;min-width:0;resize:none;overflow:hidden;padding:7px 9px;border:1px solid var(--border-strong);border-radius:var(--radius-sm);font-family:var(--sans);font-size:13px;line-height:1.45;background:var(--paper-raised);color:var(--slate);}
.pres-prow.is-title textarea{font-family:var(--serif);font-size:16px;font-weight:600;}
.pres-prow-btns{display:flex;flex-direction:column;gap:3px;}
.pres-prow-btns button{width:24px;height:24px;border:1px solid var(--border-strong);border-radius:6px;background:var(--paper-raised);color:var(--slate);cursor:pointer;font-size:14px;line-height:1;padding:0;}
.pres-prow-btns button:hover{background:var(--border);}
.pres-orig{margin-top:10px;font-size:12px;color:var(--slate-light);}
.pres-orig summary{cursor:pointer;font-weight:600;}
.pres-orig .pres-p{font-size:12px;color:var(--slate-light);}
.pres-sug{padding:16px 18px;background:var(--paper);min-width:0;display:flex;flex-direction:column;gap:10px;}
.pres-sug-title{font-family:var(--serif);font-size:15px;font-weight:600;color:var(--ink);display:flex;align-items:center;gap:8px;}
.pres-sug-item{background:var(--paper-raised);border:1px solid var(--border);border-radius:8px;padding:9px 11px;}
.pres-sug-item.done{opacity:.65;}
.pres-sug-item.done .pres-sug-text{text-decoration:line-through;}
.pres-sug-text{font-size:12.5px;line-height:1.5;color:var(--slate);white-space:pre-wrap;word-break:break-word;}
.pres-sug-meta{font-size:11px;color:var(--slate-light);margin-top:4px;}
.pres-sug-actions{display:flex;gap:10px;margin-top:6px;}
.pres-sug-actions button{background:none;border:none;padding:0;font-size:11.5px;font-weight:600;color:var(--blue);cursor:pointer;}
.pres-sug-actions button.del{color:var(--rust);}
.pres-sug textarea{width:100%;min-height:62px;resize:vertical;padding:8px 10px;border:1px solid var(--border-strong);border-radius:var(--radius-sm);font-family:var(--sans);font-size:12.5px;background:var(--paper-raised);color:var(--slate);}
.pres-sug-none{font-size:12px;color:var(--slate-light);font-style:italic;}
.pres-prog{margin:14px 0 4px;}
.pres-prog-bar{height:8px;background:var(--border);border-radius:8px;overflow:hidden;}
.pres-prog-bar span{display:block;height:100%;width:0;background:var(--blue);transition:width .2s ease;}
.pres-prog-text{font-size:12px;color:var(--slate-light);margin-top:6px;}
@media (max-width:880px){.pres-slide{grid-template-columns:1fr;}.pres-slide-main{border-right:0;border-bottom:1px solid var(--border);}}
`;
    document.head.appendChild(st);
  }

  /* ---- slide content (view + edit) ---- */
  function sortedBlocks(blocks) {
    const rank = b => b.kind === "title" ? 0 : b.kind === "subtitle" ? 1 : 2;
    return blocks.map((b, i) => ({ b, i })).sort((x, y) => rank(x.b) - rank(y.b) || x.i - y.i).map(x => x.b);
  }
  function groupItems(blocks) {
    const items = [], seen = {};
    sortedBlocks(blocks).forEach(b => {
      if (b.tbl) {
        if (!seen[b.tbl.k]) { seen[b.tbl.k] = { type: "table", cells: [] }; items.push(seen[b.tbl.k]); }
        seen[b.tbl.k].cells.push(b);
      } else items.push({ type: "block", b });
    });
    return items;
  }
  function paraViewHtml(b, p) {
    if (!p.t.trim()) return "";
    if (b.kind === "title") return `<div class="pres-title">${esc(p.t)}</div>`;
    if (b.kind === "subtitle") return `<div class="pres-subtitle">${esc(p.t)}</div>`;
    const bullet = b.paras.length > 1 || p.l > 0;
    return `<div class="pres-p" style="margin-left:${(p.l || 0) * 16}px">${bullet ? '<span class="pres-bul">•</span>' : ""}${esc(p.t)}</div>`;
  }
  function paraEditHtml(b, p, j) {
    return `<div class="pres-prow ${b.kind === "title" ? "is-title" : ""}" data-b="${b.bi}" data-p="${j}" style="margin-left:${(p.l || 0) * 14}px"><textarea rows="1" data-edit="1" spellcheck="true">\n${esc(p.t)}</textarea><span class="pres-prow-btns"><button type="button" data-act="add-para" title="Add a paragraph below">+</button><button type="button" data-act="del-para" title="Remove this paragraph">×</button></span></div>`;
  }
  function blockHtml(b, editable) {
    const inner = b.paras.map((p, j) => editable ? paraEditHtml(b, p, j) : paraViewHtml(b, p)).join("");
    if (!inner.trim()) return "";
    const label = editable ? `<div class="pres-label">${esc(b.kind === "title" ? "Title" : b.label)}</div>` : "";
    return `<div class="pres-block">${label}${inner}</div>`;
  }
  function tableHtml(item, editable) {
    let R = 0, C = 0;
    item.cells.forEach(b => { R = Math.max(R, b.tbl.r + 1); C = Math.max(C, b.tbl.c + 1); });
    const grid = Array.from({ length: R }, () => Array(C).fill(null));
    item.cells.forEach(b => { grid[b.tbl.r][b.tbl.c] = b; });
    const rows = grid.map(row => "<tr>" + row.map(b => {
      if (!b) return "<td></td>";
      return "<td>" + b.paras.map((p, j) => editable ? paraEditHtml(b, p, j) : `<div class="pres-p">${esc(p.t)}</div>`).join("") + "</td>";
    }).join("") + "</tr>").join("");
    return `<table class="pres-table"><tbody>${rows}</tbody></table>`;
  }
  function contentHtml(blocks, editable) {
    const items = groupItems(blocks);
    const html = items.map(it => it.type === "table" ? tableHtml(it, editable) : blockHtml(it.b, editable)).join("");
    return html.trim() ? html : '<div class="pres-empty-slide">No text on this slide.</div>';
  }

  function sugHtml(x) {
    return `<div class="pres-sug-item ${x.done ? "done" : ""}" data-sugid="${esc(x.id)}">
      <div class="pres-sug-text">${esc(x.text)}</div>
      <div class="pres-sug-meta">${esc(x.by)} · ${esc(fmtDateTime(x.at))}${x.done ? " · done by " + esc(x.doneBy || "—") : ""}</div>
      <div class="pres-sug-actions"><button type="button" data-act="sug-done">${x.done ? "Reopen" : "Mark done"}</button><button type="button" class="del" data-act="sug-del">Delete</button></div>
    </div>`;
  }

  function slideHtml(s, pos, total) {
    const editing = !!state.editing[s.id];
    const sugs = state.deck.sug.filter(x => x.slideId === s.id).sort((a, b) => (a.at || "").localeCompare(b.at || ""));
    const openN = sugs.filter(x => !x.done).length;
    const blocks = editing ? state.editing[s.id] : s.blocks;
    const chips = [];
    if (s.deleted) chips.push('<span class="pres-chip bad">Marked for removal</span>');
    if (s.edited) chips.push(`<span class="pres-chip ok" title="${esc(fmtDateTime(s.editedAt))}">Edited by ${esc(s.editedBy)}</span>`);
    const actions = editing
      ? `<button type="button" class="btn small" data-act="save-slide">Save</button><button type="button" class="btn ghost small" data-act="cancel-edit">Cancel</button>`
      : `${s.deleted ? "" : '<button type="button" class="btn ghost small" data-act="edit-slide">Edit text</button>'}
         <button type="button" class="btn ghost small" data-act="move-up" ${pos === 0 ? "disabled" : ""} title="Move up">↑</button>
         <button type="button" class="btn ghost small" data-act="move-down" ${pos === total - 1 ? "disabled" : ""} title="Move down">↓</button>
         <button type="button" class="btn ghost small" data-act="toggle-removed">${s.deleted ? "Restore slide" : "Remove slide"}</button>
         ${s.edited ? '<button type="button" class="btn ghost small" data-act="revert-slide">Revert text</button>' : ""}`;
    let orig = "";
    if (s.edited && !editing) {
      const ob = s.blocks.map(b => ({ ...b, paras: s.orig.find(o => o.bi === b.bi).paras }));
      orig = `<details class="pres-orig"><summary>Show original text</summary>${contentHtml(ob, false)}</details>`;
    }
    const imgs = s.imgs && s.imgs.length ? `<div class="pres-imgs">${s.imgs.map(u => `<img src="${esc(u)}" alt="">`).join("")}</div>` : "";
    return `<section class="pres-slide ${s.deleted ? "removed" : ""}" id="pslide-${esc(s.id)}" data-sid="${esc(s.id)}">
      <div class="pres-slide-main">
        <div class="pres-slide-head"><span class="pres-slide-no">Slide ${pos + 1}</span>${chips.join("")}<span class="pres-spacer"></span>${actions}</div>
        <div class="pres-slide-body">${contentHtml(blocks, editing)}${editing ? "" : imgs}</div>
        ${orig}
      </div>
      <aside class="pres-sug">
        <div class="pres-sug-title">Changes suggested ${openN ? `<span class="pres-chip warn">${openN} open</span>` : ""}</div>
        ${sugs.length ? sugs.map(sugHtml).join("") : '<div class="pres-sug-none">No suggestions yet for this slide.</div>'}
        <textarea data-sug-input placeholder="Suggest a change for this slide…"></textarea>
        <button type="button" class="btn small" data-act="add-sug">Add suggestion</button>
      </aside>
    </section>`;
  }

  function growAll(root) {
    root.querySelectorAll("textarea[data-edit]").forEach(t => { t.style.height = "auto"; t.style.height = (t.scrollHeight + 2) + "px"; });
  }
  function refreshSlide(sid) {
    const el = document.getElementById("pslide-" + sid);
    if (!el || !state.deck.slides) return;
    const typed = (el.querySelector("[data-sug-input]") || {}).value || "";
    const arr = sortedSlides();
    const pos = arr.findIndex(s => s.id === sid);
    el.outerHTML = slideHtml(arr[pos], pos, arr.length);
    const fresh = document.getElementById("pslide-" + sid);
    if (fresh) {
      const ta = fresh.querySelector("[data-sug-input]");
      if (ta && typed) ta.value = typed;
      growAll(fresh);
    }
  }
  function statsText() {
    const sl = state.deck.slides, open = state.deck.sug.filter(x => !x.done).length;
    return `${sl.length} slides · ${open} open suggestion${open === 1 ? "" : "s"} · ${sl.filter(s => s.edited).length} slide${sl.filter(s => s.edited).length === 1 ? "" : "s"} edited · ${sl.filter(s => s.deleted).length} marked for removal`;
  }
  function updateStats() {
    const el = document.getElementById("presStats");
    if (el && state.deck.slides) el.textContent = statsText();
  }

  function renderDeck() {
    const content = document.getElementById("content");
    const y = window.scrollY;
    if (!state.deck.slides) { content.innerHTML = '<div class="register"><div class="empty">Opening presentation…</div></div>'; return; }
    const arr = sortedSlides();
    content.innerHTML = `
      <div class="pres-bar">
        <button type="button" class="btn ghost small" data-act="back">← All presentations</button>
        <div class="pres-name"><label for="presName">Your name (shown with your suggestions and edits)</label><input id="presName" value="${esc(state.name)}" placeholder="e.g. Vijay Kumar, US"></div>
        <button type="button" class="btn small" data-act="dl-edited">Download edited PPTX</button>
        <button type="button" class="btn ghost small" data-act="dl-orig">Download original</button>
        <button type="button" class="btn ghost small" data-act="dl-sug">Export suggestions (CSV)</button>
        <button type="button" class="btn ghost small" data-act="refresh">Refresh</button>
      </div>
      <div class="pres-note">${esc(state.deck.meta.description || "")}${state.deck.meta.description ? "<br>" : ""}This view shows each slide's text, tables and pictures — not the exact design. The downloaded PPTX keeps the original design and applies your text edits, slide order and removals to it. Text inside charts and SmartArt cannot be edited here.</div>
      <div class="pres-stats" id="presStats">${esc(statsText())}</div>
      ${arr.map((s, i) => slideHtml(s, i, arr.length)).join("")}`;
    growAll(content);
    window.scrollTo(0, y);
  }

  function listHtml() {
    if (!state.list.length) return `<div class="register"><div class="empty"><strong>No presentations yet</strong>Click "Add presentation" to upload the first one (.pptx).</div></div>`;
    return `<div class="pres-grid">${state.list.map(p => `
      <div class="pres-card" data-act="open" data-id="${esc(p.id)}">
        <div class="pres-card-icon">${ICON}</div>
        <h3>${esc(p.title)}</h3>
        ${p.description ? `<div class="pres-desc">${esc(p.description)}</div>` : ""}
        <div class="pres-meta">${esc(p.slideCount)} slides · ${esc(fmtSize(p.fileSize || 0))} · added by ${esc(p.uploadedBy || "—")} on ${esc(fmtDate(p.uploadedAt))}</div>
        <div class="pres-meta">${p.sugCount > 0 ? `<span class="pres-chip warn">${p.sugCount} open suggestion${p.sugCount === 1 ? "" : "s"}</span>` : '<span class="pres-chip">No open suggestions</span>'}${p.lastEditedBy ? `<span class="pres-chip ok">Edited by ${esc(p.lastEditedBy)}</span>` : ""}</div>
        <div class="pres-card-actions"><button type="button" class="btn small" data-act="open" data-id="${esc(p.id)}">Open</button><button type="button" class="btn ghost small" data-act="del-deck" data-id="${esc(p.id)}">Delete</button></div>
      </div>`).join("")}</div>`;
  }

  async function renderList() {
    const content = document.getElementById("content");
    if (state.list === null) {
      content.innerHTML = '<div class="register"><div class="empty">Loading presentations…</div></div>';
      try { await loadList(); }
      catch (e) { content.innerHTML = `<div class="register"><div class="empty"><strong>Could not load presentations</strong>${esc(explain(e))}</div></div>`; return; }
      if (currentTab !== "presentations" || state.view !== "list") return;
    }
    content.innerHTML = listHtml();
  }

  function render() {
    if (typeof currentTab === "undefined" || currentTab !== "presentations") return;
    injectStyles();
    const content = document.getElementById("content");
    const addBtn = document.getElementById("addBtn");
    const title = document.getElementById("pageTitle");
    const desc = document.getElementById("pageDesc");
    content.onclick = onClick; content.oninput = onInput; content.onkeydown = onKeyDown;
    if (DEMO_MODE) {
      addBtn.style.display = "none";
      content.innerHTML = '<div class="register"><div class="empty"><strong>Needs Firebase</strong>The Presentation Repository stores files in Firestore, so it is not available in demo mode.</div></div>';
      return;
    }
    if (state.view === "deck" && state.deck) {
      addBtn.style.display = "none";
      title.textContent = state.deck.meta.title;
      desc.textContent = "Review every slide, suggest changes, or edit live — then download the updated PPTX.";
      renderDeck();
    } else {
      state.view = "list";
      title.textContent = "Presentation Repository";
      desc.textContent = "Important presentations — review every slide, suggest changes, or edit live and download the updated PPTX.";
      const lbl = addBtn.querySelector("span");
      if (lbl) lbl.textContent = "Add presentation";
      addBtn.style.display = "inline-block";
      addBtn.onclick = openUpload;
      renderList();
    }
  }

  /* ------------------------------ events ------------------------------ */
  function onClick(e) {
    if (currentTab !== "presentations") return;
    const t = e.target.closest("[data-act]");
    if (!t) return;
    const act = t.dataset.act;
    const slideEl = t.closest("[data-sid]");
    const sid = slideEl && slideEl.dataset.sid;
    switch (act) {
      case "open": openDeck(t.dataset.id); break;
      case "del-deck": e.stopPropagation(); deleteDeck(t.dataset.id); break;
      case "back": goList(); break;
      case "refresh":
        state.deck.slides = null; renderDeck();
        loadDeckData().then(() => { state.editing = {}; renderDeck(); }).catch(err => showToast("Refresh failed — " + explain(err)));
        break;
      case "dl-orig":
        withBusy(t, "Preparing…", async () => {
          await ensureJSZip();
          saveBlob(new Blob([await fetchBytes()], { type: MIME_PPTX }), state.deck.meta.fileName || (safeName(state.deck.meta.title) + ".pptx"));
        });
        break;
      case "dl-edited":
        withBusy(t, "Preparing…", async () => {
          const blob = await buildEdited();
          saveBlob(blob, safeName(state.deck.meta.title) + " (edited).pptx");
          showToast("Edited PPTX downloaded");
        });
        break;
      case "dl-sug": exportSuggestions(); break;
      case "edit-slide": {
        const s = getSlide(sid);
        state.editing[sid] = clone(s.blocks);
        refreshSlide(sid);
        break;
      }
      case "cancel-edit": delete state.editing[sid]; refreshSlide(sid); break;
      case "save-slide": saveSlide(sid); break;
      case "revert-slide": revertSlide(sid); break;
      case "toggle-removed": toggleRemoved(sid); break;
      case "move-up": moveSlide(sid, -1); break;
      case "move-down": moveSlide(sid, 1); break;
      case "add-para": case "del-para": {
        const row = t.closest(".pres-prow");
        const b = state.editing[sid].find(x => x.bi === +row.dataset.b), j = +row.dataset.p;
        if (act === "add-para") b.paras.splice(j + 1, 0, { t: "", l: b.paras[j].l, src: b.paras[j].src });
        else if (b.paras.length > 1) b.paras.splice(j, 1);
        else b.paras[0].t = "";
        refreshSlide(sid);
        if (act === "add-para") {
          const next = document.querySelector(`#pslide-${sid} .pres-prow[data-b="${b.bi}"][data-p="${j + 1}"] textarea`);
          if (next) next.focus();
        }
        break;
      }
      case "add-sug": addSuggestion(sid); break;
      case "sug-done": toggleSuggestion(t.closest("[data-sugid]").dataset.sugid); break;
      case "sug-del": deleteSuggestion(t.closest("[data-sugid]").dataset.sugid); break;
    }
  }
  function onInput(e) {
    if (currentTab !== "presentations") return;
    const t = e.target;
    if (t.id === "presName") {
      state.name = t.value;
      try { localStorage.setItem(NAME_KEY, t.value); } catch (_) {}
      return;
    }
    if (t.matches("textarea[data-edit]")) {
      const row = t.closest(".pres-prow"), slideEl = t.closest("[data-sid]");
      const draft = state.editing[slideEl.dataset.sid];
      if (draft) draft.find(x => x.bi === +row.dataset.b).paras[+row.dataset.p].t = t.value;
      t.style.height = "auto"; t.style.height = (t.scrollHeight + 2) + "px";
    }
  }
  function onKeyDown(e) {
    if (currentTab !== "presentations") return;
    // Enter = new paragraph, Shift+Enter = line break inside the paragraph
    if (e.key === "Enter" && !e.shiftKey && e.target.matches && e.target.matches("textarea[data-edit]")) {
      e.preventDefault();
      const btn = e.target.closest(".pres-prow").querySelector('[data-act="add-para"]');
      if (btn) btn.click();
    }
  }

  /* ------------------------------ upload form ------------------------------ */
  function openUpload() {
    injectStyles();
    const old = document.getElementById("entryForm");
    const form = old.cloneNode(false);   // fresh form: drops listeners left by other drawers
    old.replaceWith(form);
    document.getElementById("drawerTitle").textContent = "Add presentation";
    document.getElementById("drawerSub").textContent = "Presentation Repository";
    form.innerHTML = `
      <div class="field"><label>PowerPoint file (.pptx)</label><input type="file" id="presFile" accept=".pptx" required>
        <div class="hint">Up to 25 MB. For an older .ppt file, open it in PowerPoint and Save As .pptx first. If it is too large, use Picture Format → Compress Pictures.</div></div>
      <div class="field"><label>Title</label><input id="presTitle" required></div>
      <div class="field"><label>Description / purpose</label><textarea id="presDesc"></textarea></div>
      <div class="field"><label>Your name</label><input id="presBy" value="${esc(state.name)}" required></div>
      <div class="pres-prog" id="presProg" style="display:none"><div class="pres-prog-bar"><span></span></div><div class="pres-prog-text"></div></div>
      <div class="drawer-actions"><button type="button" class="btn ghost" id="cancelBtn">Cancel</button><button type="submit" class="btn" id="presSubmit">Upload</button></div>`;
    const fileEl = form.querySelector("#presFile"), titleEl = form.querySelector("#presTitle");
    fileEl.onchange = () => { if (!titleEl.value && fileEl.files[0]) titleEl.value = fileEl.files[0].name.replace(/\.pptx$/i, ""); };
    form.querySelector("#cancelBtn").onclick = closeDrawer;
    form.onsubmit = async ev => {
      ev.preventDefault();
      const file = fileEl.files[0];
      const by = form.querySelector("#presBy").value.trim();
      if (!file) return;
      if (!/\.pptx$/i.test(file.name)) { showToast("Please choose a .pptx file"); return; }
      if (file.size > MAX_FILE) { showToast("File is larger than 25 MB — compress its pictures first"); return; }
      state.name = by; try { localStorage.setItem(NAME_KEY, by); } catch (_) {}
      const prog = form.querySelector("#presProg"), bar = prog.querySelector("span"), txt = prog.querySelector(".pres-prog-text");
      const submit = form.querySelector("#presSubmit");
      prog.style.display = "block"; submit.disabled = true;
      try {
        const meta = await uploadDeck({ file, title: titleEl.value.trim() || file.name, description: form.querySelector("#presDesc").value.trim(), by },
          (m, f) => { txt.textContent = m; bar.style.width = Math.round(f * 100) + "%"; });
        closeDrawer();
        showToast("Presentation added");
        if (state.list) state.list.unshift(meta);
        if (currentTab === "presentations" && state.view === "list") renderList();
      } catch (err) {
        txt.textContent = "Upload failed — " + explain(err);
        submit.disabled = false;
      }
    };
    document.getElementById("overlay").classList.add("open");
  }

  window.PresModule = { render, goList };
})();
