/* Receipt Capture - photo + time + GPS stamp, queued upload to SharePoint via Microsoft Graph */
(function () {
  "use strict";
  const CFG = window.RECEIPT_CONFIG;
  const TEST = window.__RECEIPT_TEST__ || null;           // used only by automated tests
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const SCOPES = ["Files.ReadWrite.All"];
  const MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const $ = (id) => document.getElementById(id);

  // ---------- small helpers ----------
  const pad = (n, w = 2) => String(n).padStart(w, "0");
  function setMsg(text, cls) { const m = $("msg"); m.textContent = text || ""; m.className = cls ? "msg-" + cls : ""; }
  function cleanNote(s) { return (s || "").replace(/[\\/:*?"<>|#%~&{}\u0000-\u001f]/g, " ").replace(/\s+/g, " ").trim().slice(0, 60); }
  function ascii(s) { return (s || "").normalize("NFKD").replace(/[^\x20-\x7e]/g, ""); }
  function fmtCoord(v) { return v.toFixed(6); }

  // Financial year folders: July-June. July 2026 = "2027 FY" / "01. July 2026"
  function folderFor(d) {
    const m = d.getMonth() + 1, y = d.getFullYear();
    const fy = m >= 7 ? y + 1 : y;
    const n = ((m - 7 + 12) % 12) + 1;
    return { fyName: fy + " FY", monthNum: pad(n), monthName: pad(n) + ". " + MONTHS[m - 1] + " " + y };
  }

  // ---------- IndexedDB queue ----------
  let dbp = null;
  function db() {
    if (!dbp) dbp = new Promise((res, rej) => {
      const r = indexedDB.open("receipt-capture", 1);
      r.onupgradeneeded = () => r.result.createObjectStore("queue", { keyPath: "id" });
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
    return dbp;
  }
  async function store(mode, fn) {
    const d = await db();
    return new Promise((res, rej) => {
      const tx = d.transaction("queue", mode);
      const out = fn(tx.objectStore("queue"));
      tx.oncomplete = () => res(out && "result" in out ? out.result : undefined);
      tx.onerror = () => rej(tx.error);
    });
  }
  const qPut = (item) => store("readwrite", (s) => s.put(item));
  const qDel = (id) => store("readwrite", (s) => s.delete(id));
  const qAll = () => store("readonly", (s) => s.getAll());

  // ---------- auth ----------
  let pca = null, account = null, needLogin = false;
  async function initAuth() {
    if (TEST) { account = { username: "test@commitsdc.com.au" }; return renderAccount(); }
    if (!CFG.clientId || CFG.clientId.indexOf("PASTE") === 0) {
      $("account").textContent = "Client ID not set in config.js";
      return;
    }
    const base = location.origin + location.pathname.replace(/[^/]*$/, "");
    pca = new msal.PublicClientApplication({
      auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/" + CFG.tenantId, redirectUri: base + "redirect.html",
              onRedirectNavigate: () => { saveSignInState(); return true; } },
      cache: { cacheLocation: "localStorage" }
    });
    await pca.initialize();
    try {
      const r = await pca.handleRedirectPromise();
      if (r && r.account) {
        pca.setActiveAccount(r.account);
        // Android finished sign-in in a browser tab rather than the installed app: tell Tim to close it
        if (sessionStorage.getItem("rc.restored") && !sessionStorage.removeItem("rc.restored"))
          setMsg("Signed in. If you opened the app from your home screen, tap X at the top to close this page and go back to it.", "ok");
      }
    } catch (e) { setMsg("Sign-in error: " + (e.message || e), "err"); }
    try { localStorage.removeItem("rc.signin"); } catch (_) {}
    account = pca.getActiveAccount() || pca.getAllAccounts()[0] || null;
    if (account) pca.setActiveAccount(account);
    renderAccount();
  }
  // Android can finish the Microsoft sign-in in a separate browser tab, which has none of this tab's sign-in state.
  // Keep a short-lived copy so redirect.html can finish the sign-in wherever it lands.
  function saveSignInState() {
    try {
      const keep = {};
      for (let i = 0; i < sessionStorage.length; i++) { const k = sessionStorage.key(i); if (k && k.indexOf("msal.") === 0) keep[k] = sessionStorage.getItem(k); }
      localStorage.setItem("rc.signin", JSON.stringify({ t: Date.now(), keep }));
      sessionStorage.setItem("rc.pending", "1");
    } catch (_) {}
  }
  function signIn() {
    if (!pca) return;
    pca.loginRedirect({ scopes: SCOPES, prompt: "select_account" }).catch((e) => setMsg("Sign-in error: " + (e.message || e), "err"));
  }
  // Pick up a sign-in that finished in another tab (tokens are shared through localStorage)
  function recheckAccount() {
    if (!pca || (account && !needLogin)) return;
    const a = pca.getActiveAccount() || pca.getAllAccounts()[0] || null;
    if (a) { account = a; pca.setActiveAccount(a); needLogin = false; renderAccount(); processQueue(); return; }
    // Sign-in was finished in another tab (it clears rc.signin): reload so this copy of the app loads the new sign-in
    try {
      if (sessionStorage.getItem("rc.pending") && !localStorage.getItem("rc.signin")) { sessionStorage.removeItem("rc.pending"); location.reload(); }
    } catch (_) {}
  }
  function renderAccount() {
    const el = $("account");
    el.innerHTML = "";
    if (account && !needLogin) {
      el.textContent = account.username;
    } else {
      const b = document.createElement("button");
      b.className = "primary"; b.textContent = "Sign in";
      b.onclick = signIn;
      el.appendChild(b);
    }
  }
  async function getToken() {
    if (TEST) return TEST.token;
    if (!pca || !account) { needLogin = true; renderAccount(); throw new Error("Not signed in"); }
    try {
      const r = await pca.acquireTokenSilent({ scopes: SCOPES, account });
      needLogin = false;
      return r.accessToken;
    } catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError || /interaction_required|login_required|no_tokens_found|consent_required/i.test(e.errorCode || e.message)) {
        needLogin = true; renderAccount();
        throw new Error("Sign-in expired, tap Sign in (photos are kept and will upload after)");
      }
      throw e;
    }
  }

  // ---------- Graph ----------
  async function graph(path, opts = {}) {
    const token = await getToken();
    const r = await fetch(path.indexOf("http") === 0 ? path : GRAPH + path, {
      ...opts, headers: { Authorization: "Bearer " + token, ...(opts.headers || {}) }
    });
    if (!r.ok) {
      let detail = "";
      try { const j = await r.json(); detail = (j.error && (j.error.message || j.error.code)) || ""; } catch (_) {}
      const err = new Error("SharePoint " + r.status + (detail ? ": " + detail : ""));
      err.status = r.status;
      throw err;
    }
    return r.status === 204 ? null : r.json();
  }
  const D = () => "/drives/" + encodeURIComponent(CFG.driveId);
  async function children(itemId) {
    let url = D() + "/items/" + itemId + "/children?$select=id,name,folder&$top=999";
    const all = [];
    while (url) { const j = await graph(url); all.push(...j.value); url = j["@odata.nextLink"] || null; }
    return all.filter((c) => c.folder);
  }
  async function createFolder(parentId, name) {
    try {
      return (await graph(D() + "/items/" + parentId + "/children", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" })
      })).id;
    } catch (e) {
      if (e.status === 409) { const hit = (await children(parentId)).find((c) => c.name === name); if (hit) return hit.id; }
      throw e;
    }
  }
  function cacheGet(k) { try { return JSON.parse(localStorage.getItem("rc.folders") || "{}")[k]; } catch (_) { return null; } }
  function cacheSet(k, v) { try { const c = JSON.parse(localStorage.getItem("rc.folders") || "{}"); c[k] = v; localStorage.setItem("rc.folders", JSON.stringify(c)); } catch (_) {} }
  function cacheClear() { try { localStorage.removeItem("rc.folders"); } catch (_) {} }

  // Walk by item ID (no colon-path lookups). Base path must already exist; FY and month folders are created if missing.
  async function resolveFolder(f) {
    const key = f.fyName + "/" + f.monthNum;
    const cached = cacheGet(key);
    if (cached) return cached;
    let id = cacheGet("__base");
    if (!id) {
      id = (await graph(D() + "/root?$select=id")).id;
      for (const part of CFG.basePath) {
        const hit = (await children(id)).filter((c) => c.name.trim().toLowerCase() === part.toLowerCase());
        if (!hit.length) throw new Error('Folder "' + part + '" not found in SharePoint path ' + CFG.basePath.join("/"));
        id = hit[0].id;
      }
      cacheSet("__base", id);
    }
    const fyHits = (await children(id)).filter((c) => c.name.trim().toLowerCase() === f.fyName.toLowerCase());
    const fyId = fyHits.length ? fyHits[0].id : await createFolder(id, f.fyName);
    // Month folders are "NN. Month YYYY" but some use short names (e.g. "07. Jan 2026"), so match on the "NN." prefix
    const mHits = (await children(fyId)).filter((c) => c.name.trim().indexOf(f.monthNum + ".") === 0);
    const exact = mHits.find((c) => c.name === f.monthName);
    const mId = exact ? exact.id : mHits.length ? mHits[0].id : await createFolder(fyId, f.monthName);
    cacheSet(key, mId);
    return mId;
  }
  async function uploadItem(item) {
    const put = async () => {
      const folderId = await resolveFolder(item.folder);
      return graph(D() + "/items/" + folderId + ":/" + encodeURIComponent(item.filename) + ":/content?@microsoft.graph.conflictBehavior=rename", {
        method: "PUT", headers: { "Content-Type": "image/jpeg" }, body: item.blob
      });
    };
    try { return await put(); }
    catch (e) { if (e.status === 404) { cacheClear(); return put(); } throw e; }
  }

  // ---------- queue processing ----------
  let busy = false;
  async function processQueue() {
    if (busy) return;
    busy = true;
    try {
      const items = (await qAll()).filter((i) => i.status !== "done").sort((a, b) => a.id - b.id);
      for (const item of items) {
        if (TEST ? TEST.offline : !navigator.onLine) { item.status = "pending"; item.error = "Offline, will upload when back in coverage"; await qPut(item); continue; }
        try {
          const r = await uploadItem(item);
          item.status = "done"; item.error = ""; item.savedAs = r && r.name ? r.name : item.filename; item.webUrl = r && r.webUrl;
          delete item.blob;                                  // free phone storage once it is in SharePoint
          await qPut(item);
          setMsg("Saved to " + item.folder.fyName + " / " + item.folder.monthName, "ok");
        } catch (e) {
          item.status = "failed"; item.error = e.message || String(e); item.tries = (item.tries || 0) + 1;
          await qPut(item);
          if (/Sign/.test(item.error)) break;
        }
        await renderQueue();
      }
      // tidy: forget finished items older than 30 days
      const cutoff = Date.now() - 30 * 864e5;
      for (const i of await qAll()) if (i.status === "done" && i.id < cutoff) await qDel(i.id);
    } finally {
      busy = false;
      await renderQueue();
    }
  }
  async function renderQueue() {
    const items = (await qAll()).sort((a, b) => b.id - a.id).slice(0, 25);
    const ul = $("queue");
    ul.innerHTML = "";
    if (!items.length) { ul.innerHTML = '<li><span class="sub">Nothing yet</span></li>'; return; }
    for (const i of items) {
      const li = document.createElement("li");
      const left = document.createElement("div");
      const nm = document.createElement("div"); nm.className = "name"; nm.textContent = i.savedAs || i.filename;
      const sub = document.createElement("div"); sub.className = "sub";
      sub.textContent = i.folder.fyName + " / " + i.folder.monthName + (i.error ? " - " + i.error : "");
      left.append(nm, sub);
      const right = document.createElement("div");
      const st = document.createElement("span");
      st.className = "st " + (i.status === "done" ? "done" : i.status === "failed" ? "failed" : "pending");
      st.textContent = i.status === "done" ? "Uploaded" : i.status === "failed" ? "Failed" : "Waiting";
      right.appendChild(st);
      if (i.status !== "done") {
        const del = document.createElement("button");
        del.textContent = "Discard";
        del.onclick = async () => { if (confirmDiscard(i)) { await qDel(i.id); renderQueue(); } };
        right.append(document.createElement("br"), del);
      }
      li.append(left, right);
      ul.appendChild(li);
    }
  }
  function confirmDiscard(i) { return TEST ? true : window.confirm("Discard " + i.filename + "? It has not been uploaded."); }

  // ---------- location ----------
  let fix = null, fixErr = null;
  function gpsStatus() {
    const el = $("gps");
    const age = fix ? Math.round((Date.now() - fix.time) / 1000) : null;
    if (fix) {
      const good = fix.acc <= 50;
      el.innerHTML = '<span class="dot ' + (good ? "ok" : "warn") + '"></span>' +
        fmtCoord(fix.lat) + ", " + fmtCoord(fix.lon) + " (±" + Math.round(fix.acc) + " m" + (age > 60 ? ", " + Math.round(age / 60) + " min old" : "") + ")";
    } else if (fixErr) {
      el.innerHTML = '<span class="dot err"></span>Location unavailable: ' + fixErr + ". Photos will save without GPS.";
    } else {
      el.innerHTML = '<span class="dot"></span>Getting location…';
    }
  }
  function onPos(p) { fix = { lat: p.coords.latitude, lon: p.coords.longitude, acc: p.coords.accuracy, time: p.timestamp || Date.now() }; fixErr = null; gpsStatus(); }
  function onPosErr(e) { if (!fix) { fixErr = e.code === 1 ? "permission denied" : e.message || "no fix"; gpsStatus(); } }
  function startGps() {
    if (!("geolocation" in navigator)) { fixErr = "not supported"; return gpsStatus(); }
    navigator.geolocation.watchPosition(onPos, onPosErr, { enableHighAccuracy: true, maximumAge: 10000, timeout: 30000 });
  }
  function freshFix() {
    if (fix && Date.now() - fix.time < 120000) return Promise.resolve(fix);
    // Hard 12 s limit: getCurrentPosition never answers while a location prompt is left open, which used to hang "Stamping photo"
    return new Promise((res) => {
      if (!("geolocation" in navigator)) return res(fix);
      const t = setTimeout(() => res(fix), 12000);
      navigator.geolocation.getCurrentPosition((p) => { clearTimeout(t); onPos(p); res(fix); }, () => { clearTimeout(t); res(fix); }, { enableHighAccuracy: true, maximumAge: 0, timeout: 12000 });
    });
  }
  async function placeName(f) {
    if (!f || !navigator.onLine || TEST && !TEST.geocode) return "";
    try {
      const ctl = new AbortController(); const t = setTimeout(() => ctl.abort(), 4000);
      const r = await fetch("https://nominatim.openstreetmap.org/reverse?format=jsonv2&zoom=18&addressdetails=1&lat=" + f.lat + "&lon=" + f.lon, { signal: ctl.signal, headers: { "Accept-Language": "en-AU" } });
      clearTimeout(t);
      if (!r.ok) return "";
      const j = await r.json(), a = j.address || {};
      const place = a.shop || a.amenity || a.fuel || a.building || a.tourism || "";
      const street = [a.house_number, a.road].filter(Boolean).join(" ");
      const town = a.suburb || a.town || a.city || a.village || a.hamlet || a.locality || "";
      return [place, street, town].filter(Boolean).join(", ");
    } catch (_) { return ""; }
  }

  // ---------- image processing ----------
  function blobToDataURL(b) { return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = () => rej(r.error); r.readAsDataURL(b); }); }
  function dataURLToBlob(u) { const bin = atob(u.split(",")[1]); const a = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) a[i] = bin.charCodeAt(i); return new Blob([a], { type: "image/jpeg" }); }

  async function buildImage(file, when, f, place, note) {
    const bmp = await createImageBitmap(file, { imageOrientation: "from-image" });
    const scale = Math.min(1, CFG.maxImageSize / Math.max(bmp.width, bmp.height));
    const W = Math.round(bmp.width * scale), H = Math.round(bmp.height * scale);
    const tz = new Intl.DateTimeFormat("en-AU", { timeZoneName: "short" }).formatToParts(when).find((p) => p.type === "timeZoneName");
    const lines = [
      pad(when.getDate()) + "/" + pad(when.getMonth() + 1) + "/" + when.getFullYear() + "  " + pad(when.getHours()) + ":" + pad(when.getMinutes()) + ":" + pad(when.getSeconds()) + (tz ? " " + tz.value : ""),
      f ? "GPS " + fmtCoord(f.lat) + ", " + fmtCoord(f.lon) + "  (±" + Math.round(f.acc) + " m)" : "GPS unavailable"
    ];
    if (place) lines.push(place);
    if (note) lines.push(note);
    const fs = Math.max(18, Math.round(W * 0.03)), lh = Math.round(fs * 1.35), padv = Math.round(fs * 0.6);
    const band = lines.length * lh + padv * 2;
    const c = document.createElement("canvas");
    c.width = W; c.height = H + band;                    // stamp goes BELOW the photo so it never covers receipt text
    const g = c.getContext("2d");
    g.fillStyle = "#000"; g.fillRect(0, 0, c.width, c.height);
    g.drawImage(bmp, 0, 0, W, H);
    bmp.close && bmp.close();
    g.fillStyle = "#fff"; g.textBaseline = "top";
    g.font = "600 " + fs + "px system-ui, Roboto, Arial, sans-serif";
    lines.forEach((t, i) => {
      let s = t;
      while (g.measureText(s).width > W - padv * 2 && s.length > 4) s = s.slice(0, -2) + "…";
      g.fillText(s, padv, H + padv + i * lh);
    });
    let dataUrl = c.toDataURL("image/jpeg", CFG.jpegQuality);
    // EXIF so SharePoint/Photos show the date and map location
    try {
      const zeroth = {}, exif = {}, gps = {};
      const exifDate = when.getFullYear() + ":" + pad(when.getMonth() + 1) + ":" + pad(when.getDate()) + " " + pad(when.getHours()) + ":" + pad(when.getMinutes()) + ":" + pad(when.getSeconds());
      zeroth[piexif.ImageIFD.ImageDescription] = ascii("Receipt " + lines.join(" | "));
      zeroth[piexif.ImageIFD.Software] = "Receipt Capture";
      zeroth[piexif.ImageIFD.DateTime] = exifDate;
      exif[piexif.ExifIFD.DateTimeOriginal] = exifDate;
      exif[piexif.ExifIFD.DateTimeDigitized] = exifDate;
      if (f) {
        gps[piexif.GPSIFD.GPSVersionID] = [2, 3, 0, 0];
        gps[piexif.GPSIFD.GPSLatitudeRef] = f.lat < 0 ? "S" : "N";
        gps[piexif.GPSIFD.GPSLatitude] = piexif.GPSHelper.degToDmsRational(Math.abs(f.lat));
        gps[piexif.GPSIFD.GPSLongitudeRef] = f.lon < 0 ? "W" : "E";
        gps[piexif.GPSIFD.GPSLongitude] = piexif.GPSHelper.degToDmsRational(Math.abs(f.lon));
        gps[piexif.GPSIFD.GPSHPositioningError] = [Math.round(f.acc * 100), 100];
      }
      dataUrl = piexif.insert(piexif.dump({ "0th": zeroth, Exif: exif, GPS: gps }), dataUrl);
    } catch (e) { console.warn("EXIF write failed", e); }
    return { blob: dataURLToBlob(dataUrl), dataUrl };
  }

  // ---------- capture flow ----------
  async function onPhoto(file) {
    if (!file) return;
    const when = TEST && TEST.now ? new Date(TEST.now) : new Date();
    $("shoot").disabled = true;
    setMsg("Stamping photo…");
    try {
      const f = await freshFix();
      const note = cleanNote($("note").value);
      const place = await placeName(f);
      const folder = folderFor(when);
      const { blob, dataUrl } = await buildImage(file, when, f, place, note);
      const filename = when.getFullYear() + "-" + pad(when.getMonth() + 1) + "-" + pad(when.getDate()) + " " + pad(when.getHours()) + pad(when.getMinutes()) + (note ? " " + note : " Receipt") + ".jpg";
      const item = { id: Date.now(), filename, folder, blob, status: "pending", error: "", created: when.toISOString(), gps: f ? { lat: f.lat, lon: f.lon, acc: f.acc } : null };
      await qPut(item);
      $("preview").src = dataUrl; $("preview").style.display = "block";
      $("note").value = "";
      setMsg("Saved on phone, uploading…");
      await renderQueue();
      processQueue();
    } catch (e) {
      setMsg("Could not process photo: " + (e.message || e), "err");
    } finally {
      $("shoot").disabled = false;
      $("file").value = "";
    }
  }

  // ---------- start ----------
  async function start() {
    $("shoot").onclick = () => $("file").click();
    $("file").onchange = (e) => onPhoto(e.target.files[0]);
    $("retry").onclick = () => processQueue();
    window.addEventListener("online", () => processQueue());
    document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") { recheckAccount(); processQueue(); } });
    setInterval(() => processQueue(), 60000);
    gpsStatus(); startGps();
    if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
    if ("serviceWorker" in navigator && !TEST) navigator.serviceWorker.register("sw.js").catch(() => {});
    await renderQueue();
    await initAuth();
    if (/signin=lost/.test(location.search) && !account) setMsg("Sign-in didn't come back to the app. Tap Sign in again.", "err");
    if (account) processQueue();
  }
  window.ReceiptApp = { folderFor, processQueue, onPhoto, qAll, cacheClear };
  start();
})();
