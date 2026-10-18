/* SRL Reply Tracker - Outlook side panel */
(function () {
  "use strict";
  const CFG = Object.assign({ daysBack: 60, inDaysBack: 30, redAfter: 3, weekend: [5, 6], alsoTrack: [], ignore: [] }, window.TRACKER_CONFIG || {});
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const SCOPES = ["Mail.Read", "User.Read"];
  const DEMO = /[?&]demo=1/.test(location.search);
  const $ = id => document.getElementById(id);
  const esc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const fmt = d => d ? new Date(d).toLocaleDateString(undefined, { day: "2-digit", month: "short" }) + " " + new Date(d).toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" }) : "";
  const state = { tab: "out", data: null, updated: null };
  let pca = null;

  // ---------- small persistence (per viewer, optional) ----------
  const store = {
    get(k) { try { return JSON.parse(localStorage.getItem("srlrt-" + k)); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem("srlrt-" + k, JSON.stringify(v)); } catch (e) { } }
  };

  function setStatus(t) { $("status").textContent = t; }
  function progress(p) { $("bar").style.width = (p || 0) + "%"; }
  function showError(msg) { $("error").innerHTML = msg ? `<div class="err">${esc(msg)}</div>` : ""; }

  // ---------- sign-in (Nested App Authentication) ----------
  async function initAuth() {
    if (!CFG.clientId || /PASTE/.test(CFG.clientId)) throw new Error("Setup not finished: open config.js and paste your Application (client) ID.");
    if (!window.msal || !msal.createNestablePublicClientApplication) throw new Error("Sign-in library did not load (msal-browser.min.js missing).");
    pca = await msal.createNestablePublicClientApplication({
      auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/organizations" },
      cache: { cacheLocation: "localStorage" }
    });
  }
  async function getToken(interactive) {
    const req = { scopes: SCOPES };
    try {
      const acct = pca.getActiveAccount() || pca.getAllAccounts()[0];
      const r = await pca.acquireTokenSilent(acct ? Object.assign({ account: acct }, req) : req);
      return r.accessToken;
    } catch (e) {
      if (!interactive) throw Object.assign(new Error("needs-signin"), { needsSignin: true });
      const r = await pca.acquireTokenPopup(req);
      if (r.account) pca.setActiveAccount(r.account);
      return r.accessToken;
    }
  }

  // ---------- Microsoft Graph ----------
  let token = null;
  async function g(url) {
    for (let i = 0; i < 6; i++) {
      const res = await fetch(url.startsWith("http") ? url : GRAPH + url, {
        headers: { Authorization: "Bearer " + token, Prefer: 'outlook.body-content-type="html"' }
      });
      if (res.ok) return res.json();
      if (res.status === 401 && i === 0) { token = await getToken(false); continue; }
      if (res.status === 429 || res.status >= 500) { await new Promise(r => setTimeout(r, (+res.headers.get("Retry-After") || 4) * 1000)); continue; }
      let msg = res.status + " " + res.statusText;
      try { const j = await res.json(); if (j.error) msg = j.error.code + ": " + j.error.message; } catch (e) { }
      throw new Error(msg);
    }
    throw new Error("Microsoft 365 is busy, try Refresh in a minute.");
  }
  async function gAll(url, max, stopWhen) {
    const all = [];
    while (url && all.length < (max || 20000)) {
      const r = await g(url);
      const v = r.value || [];
      all.push(...v);
      if (stopWhen && v.length && stopWhen(v[v.length - 1])) break;
      url = r["@odata.nextLink"];
    }
    return all;
  }
  const iso = d => new Date(d).toISOString().replace(/\.\d{3}Z$/, "Z");

  async function loadFromGraph() {
    const now = new Date();
    const cutRecv = new Date(now - Math.max(CFG.daysBack, CFG.inDaysBack) * 864e5);
    const cutIn = new Date(now - CFG.inDaysBack * 864e5);

    setStatus("Reading your profile…"); progress(5);
    const meRaw = await g("/me?$select=displayName,mail,userPrincipalName,otherMails,proxyAddresses");
    const addrs = new Set([meRaw.mail, meRaw.userPrincipalName, ...(meRaw.otherMails || [])].filter(Boolean).map(s => s.toLowerCase()));
    (meRaw.proxyAddresses || []).forEach(p => { const m = /^smtp:(.+)$/i.exec(p); if (m) addrs.add(m[1].toLowerCase()); });
    const me = { name: meRaw.displayName || "", addrs };

    setStatus("Reading your sent emails…"); progress(15);
    const sent = await gAll(`/me/mailFolders/sentitems/messages?$filter=sentDateTime ge ${iso(cutRecv)}&$orderby=sentDateTime desc&$select=id,subject,sentDateTime,conversationId,toRecipients,ccRecipients,uniqueBody&$top=50`);

    setStatus(`Reading received emails… (${sent.length} sent checked)`); progress(45);
    const received = await gAll(`/me/messages?$filter=receivedDateTime ge ${iso(cutRecv)}&$orderby=receivedDateTime desc&$select=id,subject,receivedDateTime,from,sender,conversationId,toRecipients,ccRecipients,isDraft&$top=250`);

    setStatus("Finding emails where you were @mentioned…"); progress(75);
    let mentionedMe = [];
    try {
      mentionedMe = await gAll(`https://graph.microsoft.com/beta/me/messages?$filter=mentionsPreview/isMentioned eq true&$select=id,subject,receivedDateTime,from,sender,conversationId,isDraft,uniqueBody&$top=50`,
        1000, last => new Date(last.receivedDateTime) < cutIn);
    } catch (e) {
      // fallback: open emails addressed to me and look for an @mention of me
      const cands = received.filter(m => new Date(m.receivedDateTime) >= cutIn &&
        [...(m.toRecipients || []), ...(m.ccRecipients || [])].some(r => r.emailAddress && addrs.has((r.emailAddress.address || "").toLowerCase()))).slice(0, 150);
      const outArr = [];
      for (let i = 0; i < cands.length; i += 4) {
        const part = await Promise.all(cands.slice(i, i + 4).map(m => g(`/me/messages/${encodeURIComponent(m.id)}?$select=uniqueBody`).then(r => ({ ...m, uniqueBody: r.uniqueBody })).catch(() => null)));
        for (const m of part) if (m && TrackerCore.getMentions(TrackerCore.ownPart(m.uniqueBody && m.uniqueBody.content)).some(x => (x.email && addrs.has(x.email.toLowerCase())) || TrackerCore.normName(x.name) === TrackerCore.normName(me.name))) outArr.push(m);
        progress(75 + Math.round(20 * i / Math.max(1, cands.length)));
      }
      mentionedMe = outArr;
    }
    progress(100);
    const data = TrackerCore.build(me, sent, received, mentionedMe, CFG, now);
    data.user = me.name;
    return data;
  }

  // ---------- demo data (open taskpane.html?demo=1 in a browser to preview) ----------
  function demoData() {
    const h = x => new Date(Date.now() - x * 36e5);
    return {
      user: "Demo", out: [
        { dir: "out", person: "Saad Iqbal", email: "saad.iqbal@srlksa.com", subject: "Aramco FCL shipment Collection LSR#710000 Shipment#773606", firstSent: h(1.5), lastSent: h(1.5), reminders: 0, replied: false },
        { dir: "out", person: "Saad Iqbal", email: "saad.iqbal@srlksa.com", subject: "Request for FROs - August 2026", firstSent: h(150), lastSent: h(50), reminders: 2, replied: false },
        { dir: "out", person: "Adeel Sajid", email: "adeel.sajid@srlksa.com", subject: "Pending Invoices", firstSent: h(52), lastSent: h(52), reminders: 0, replied: false },
        { dir: "out", person: "Ricardo Dayata - DSV", email: "ricardo.dayata@dsv.com", subject: "Port Shuttling Request / Inspection Required", firstSent: h(72), lastSent: h(72), reminders: 0, replied: true, replyAt: h(40), replyFrom: "Ricardo Dayata - DSV" }
      ], in: [
        { dir: "in", person: "Shamnab Abdulazeez", email: "shamnab@example.com", subject: "Export DG full operations Jeddah loc", firstSent: h(98), lastSent: h(98), reminders: 0, replied: false }
      ]
    };
  }

  // ---------- render ----------
  function revive(d) {
    for (const k of ["out", "in"]) for (const r of d[k] || []) for (const f of ["firstSent", "lastSent", "replyAt"]) if (r[f]) r[f] = new Date(r[f]);
    return d;
  }
  function render() {
    const d = state.data; if (!d) return;
    $("main").hidden = false;
    const work = $("workdays").checked, red = CFG.redAfter;
    const list = (d[state.tab] || []).map(r => Object.assign({}, r, { days: TrackerCore.dayCount(r.firstSent, r.replied ? r.replyAt : new Date(), work, CFG.weekend) }));
    $("nOut").textContent = (d.out || []).filter(r => !r.replied).length;
    $("nIn").textContent = (d.in || []).filter(r => !r.replied).length;

    const open = list.filter(r => !r.replied), done = list.filter(r => r.replied);
    const late = open.filter(r => r.days >= red).length;
    const avg = done.length ? (done.reduce((s, r) => s + r.days, 0) / done.length).toFixed(1) : "–";
    $("kpis").innerHTML = `
      <div class="kpi ${open.length ? "bad" : "ok"}"><div class="l">No reply</div><div class="v">${open.length}</div></div>
      <div class="kpi ${late ? "warn" : ""}"><div class="l">${red}+ days</div><div class="v">${late}</div></div>
      <div class="kpi ok"><div class="l">Avg reply</div><div class="v">${avg}<span style="font-size:12px"> d</span></div></div>`;

    // person filter options
    const ps = $("personSel"), cur = ps.value;
    const people = [...new Map(list.map(r => [(r.email || r.person || "").toLowerCase(), r.person])).entries()].sort((a, b) => a[1].localeCompare(b[1]));
    ps.innerHTML = `<option value="">Everyone</option>` + people.map(([k, n]) => `<option value="${esc(k)}">${esc(n)}</option>`).join("");
    ps.value = people.some(p => p[0] === cur) ? cur : "";

    const q = $("q").value.trim().toLowerCase(), st = $("statusSel").value, pf = ps.value;
    const rows = list.filter(r => (st === "all" || (st === "open" ? !r.replied : r.replied))
      && (!pf || (r.email || r.person || "").toLowerCase() === pf)
      && (!q || (r.subject + " " + r.person + " " + r.email).toLowerCase().includes(q)))
      .sort((a, b) => (a.replied - b.replied) || (b.days - a.days));

    $("list").innerHTML = rows.length ? rows.map((r, i) => {
      const cls = r.replied ? "done" : r.days >= red ? "d2" : r.days >= Math.max(1, red - 1) ? "d1" : "";
      const meta = r.replied
        ? `Sent ${fmt(r.firstSent)} · replied ${fmt(r.replyAt)}`
        : `Sent ${fmt(r.firstSent)}${r.reminders ? ` · ${r.reminders} follow-up${r.reminders > 1 ? "s" : ""}, last ${fmt(r.lastSent)}` : ""}`;
      return `<div class="card ${cls}" data-i="${i}" title="Click to open this email">
        <div class="h"><div class="who">${esc(r.person)}</div><div class="days">${r.replied ? "replied " : ""}${r.days}d</div></div>
        <div class="subj">${esc(r.subject)}</div>
        <div class="meta">${meta}</div></div>`;
    }).join("") : `<div class="empty">${st === "open" ? "Everyone has replied. 👍" : "Nothing here."}</div>`;
    $("list").querySelectorAll(".card").forEach(c => c.onclick = () => openItem(rows[+c.dataset.i]));
  }
  function openItem(r) {
    if (!r || !r.openId || !window.Office || !Office.context || !Office.context.mailbox) return;
    try {
      const ewsId = Office.context.mailbox.convertToEwsId(r.openId, Office.MailboxEnums.RestVersion.v2_0);
      Office.context.mailbox.displayMessageForm(ewsId);
    } catch (e) { }
  }

  // ---------- main ----------
  async function refresh(interactive) {
    showError("");
    $("refresh").disabled = true;
    try {
      if (DEMO) { state.data = revive(demoData()); state.updated = new Date(); }
      else {
        if (!pca) await initAuth();
        try { token = await getToken(interactive); }
        catch (e) {
          if (e.needsSignin) { $("signin").hidden = false; setStatus("Not signed in"); return; }
          throw e;
        }
        $("signin").hidden = true;
        state.data = await loadFromGraph();
        state.updated = new Date();
        store.set("cache", { data: state.data, updated: state.updated });
      }
      setStatus(`${state.data.user ? state.data.user + " · " : ""}updated ${fmt(state.updated)}`);
      render();
    } catch (e) {
      let msg = e.message || String(e);
      if (/AADSTS65001|consent|admin/i.test(msg)) msg += "\n\nSRL IT may need to approve this app (Mail.Read) in Microsoft Entra.";
      if (/AADSTS50011|redirect/i.test(msg)) msg += "\n\nCheck the redirect URI in your app registration: brk-multihub://<your github.io domain>";
      showError(msg);
      setStatus("Could not refresh");
    } finally { $("refresh").disabled = false; setTimeout(() => progress(0), 600); }
  }

  function wire() {
    document.querySelectorAll(".tab").forEach(t => t.onclick = () => {
      state.tab = t.dataset.tab;
      document.querySelectorAll(".tab").forEach(x => x.setAttribute("aria-selected", x === t));
      render();
    });
    ["q", "statusSel", "personSel"].forEach(id => $(id).addEventListener("input", render));
    $("workdays").checked = !!store.get("workdays");
    $("workdays").onchange = () => { store.set("workdays", $("workdays").checked); render(); };
    $("refresh").onclick = () => refresh(true);
    $("signinBtn").onclick = () => refresh(true);
    const c = store.get("cache");
    if (c && c.data && !DEMO) { state.data = revive(c.data); state.updated = new Date(c.updated); setStatus("Showing last result from " + fmt(state.updated) + " · refreshing…"); render(); }
  }

  function start() { wire(); refresh(false); }
  if (window.Office && Office.onReady) Office.onReady(() => start()); else document.addEventListener("DOMContentLoaded", start);
})();
