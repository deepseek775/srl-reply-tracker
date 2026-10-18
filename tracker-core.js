/* SRL Reply Tracker - matching logic (no Outlook / network code here) */
(function (root) {
  "use strict";

  function normName(s) {
    if (!s) return "";
    s = String(s).replace(/\s+-\s+.*$/, "").replace(/\(.*?\)/g, "");
    return s.toLowerCase().replace(/[^a-z؀-ۿ]/g, "");
  }
  function samePerson(nameA, mailA, nameB, mailB) {
    if (mailA && mailB && mailA.toLowerCase() === mailB.toLowerCase()) return true;
    const a = normName(nameA), b = normName(nameB);
    return a.length >= 4 && b.length >= 4 && (a === b || a.startsWith(b) || b.startsWith(a));
  }
  function topicKey(t) {
    if (!t) return "";
    return String(t).replace(/^\s*((re|fw|fwd|aw|wg|tr|rv)\s*(\[\d+\])?\s*:\s*)+/i, "").toLowerCase().replace(/\s+/g, " ").trim();
  }
  // keep only what the sender wrote (cut quoted thread)
  function ownPart(html) {
    if (!html) return "";
    const markers = ['id="divRplyFwdMsg"', "id=divRplyFwdMsg", 'id="appendonsend"', "id=appendonsend",
      "border-top:solid #E1E1E1", "border-top: solid #E1E1E1", "border-top:solid #B5C4DF",
      'id="mail-editor-reference-message-container"', "-----Original Message-----", "<hr"];
    const low = html.toLowerCase();
    let cut = html.length;
    for (const m of markers) { const i = low.indexOf(m.toLowerCase()); if (i >= 0 && i < cut) cut = i; }
    return html.slice(0, cut);
  }
  function decode(s) {
    return s.replace(/&nbsp;/gi, " ").replace(/&amp;/gi, "&").replace(/&lt;/gi, "<").replace(/&gt;/gi, ">")
      .replace(/&quot;/gi, '"').replace(/&#39;/gi, "'").replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n));
  }
  // Outlook @mentions = links whose text starts with "@" pointing to mailto:person
  function getMentions(html) {
    const found = new Map();
    if (!html) return [];
    const rx = /<a\b([^>]*)>([\s\S]*?)<\/a>/gi;
    let m;
    while ((m = rx.exec(html))) {
      let text = decode(m[2].replace(/<[^>]+>/g, "")).replace(/[ ​]/g, " ").trim();
      if (!text.startsWith("@")) continue;
      const name = text.replace(/^@+/, "").trim();
      let mail = "";
      const mm = /href\s*=\s*["']?mailto:([^"'\s>?]+)/i.exec(m[1]);
      if (mm) { try { mail = decodeURIComponent(mm[1]); } catch (e) { mail = mm[1]; } }
      const key = mail ? mail.toLowerCase() : normName(name);
      if (key && !found.has(key)) found.set(key, { name, email: mail, via: "@mention" });
    }
    return [...found.values()];
  }
  const addr = r => (r && r.emailAddress && r.emailAddress.address) || "";
  const aname = r => (r && r.emailAddress && r.emailAddress.name) || "";
  function index(map, key, obj) { if (!key) return; (map[key] = map[key] || []).push(obj); }

  /**
   * me:       { name, addrs:Set<string lower> }
   * sent:     Graph messages from Sent Items (with uniqueBody)
   * received: Graph messages (headers) from all folders
   * mentionedMe: Graph messages that @mention me (with uniqueBody optional)
   */
  function build(me, sent, received, mentionedMe, cfg, now) {
    now = now || new Date();
    const cutOut = new Date(now - cfg.daysBack * 864e5);
    const cutIn = new Date(now - cfg.inDaysBack * 864e5);
    const ignore = new Set((cfg.ignore || []).map(s => s.toLowerCase()));
    const also = new Set((cfg.alsoTrack || []).map(s => s.toLowerCase()));

    const recvByConv = {}, recvByTopic = {};
    for (const m of received) {
      if (m.isDraft) continue;
      const email = addr(m.from) || addr(m.sender), name = aname(m.from) || aname(m.sender);
      if (email && (me.addrs.has(email.toLowerCase()) || ignore.has(email.toLowerCase()))) continue;
      const o = { id: m.id, name, email, time: new Date(m.receivedDateTime), conv: m.conversationId || "", topic: topicKey(m.subject) };
      index(recvByConv, o.conv, o); index(recvByTopic, o.topic, o);
    }

    const sentByConv = {}, sentByTopic = {}, requests = [];
    for (const m of sent) {
      const t = new Date(m.sentDateTime), conv = m.conversationId || "", topic = topicKey(m.subject);
      index(sentByConv, conv, { time: t }); index(sentByTopic, topic, { time: t });
      if (t < cutOut) continue;
      const people = getMentions(ownPart(m.uniqueBody && m.uniqueBody.content));
      const rcpts = [...(m.toRecipients || []), ...(m.ccRecipients || [])];
      for (const p of people) if (!p.email) { const r = rcpts.find(r => samePerson(p.name, "", aname(r), "")); if (r) p.email = addr(r); }
      if (also.size) for (const r of (m.toRecipients || [])) {
        const sm = addr(r).toLowerCase();
        if (sm && also.has(sm) && !people.some(p => samePerson(p.name, p.email, aname(r), sm))) people.push({ name: aname(r), email: addr(r), via: "To" });
      }
      for (const p of people) {
        const pe = (p.email || "").toLowerCase();
        if (pe && (me.addrs.has(pe) || ignore.has(pe))) continue;
        requests.push({ person: p.name, email: p.email, subject: m.subject || "", sent: t, conv, topic, id: m.id, via: p.via });
      }
    }

    function findReply(req) {
      const c = [...(recvByConv[req.conv] || []), ...(recvByTopic[req.topic] || [])];
      let best = null;
      for (const x of c) {
        if (x.time <= req.sent) continue;
        if (!samePerson(req.person, req.email, x.name, x.email)) continue;
        if (!best || x.time < best.time) best = x;
      }
      return best;
    }

    // group: same conversation + same person; follow-ups answered by the same reply are merged
    const groups = {};
    for (const r of requests) {
      const k = (r.conv || r.topic) + "|" + (r.email ? r.email.toLowerCase() : normName(r.person));
      (groups[k] = groups[k] || []).push(r);
    }
    const out = [];
    for (const k in groups) {
      const reqs = groups[k].sort((a, b) => a.sent - b.sent);
      const byReply = new Map();
      for (const r of reqs) {
        const rep = findReply(r);
        const key = rep ? rep.time.toISOString() : "OPEN";
        if (!byReply.has(key)) byReply.set(key, { reqs: [], reply: rep });
        byReply.get(key).reqs.push(r);
      }
      for (const set of byReply.values()) {
        const first = set.reqs[0], last = set.reqs[set.reqs.length - 1];
        out.push({
          dir: "out", person: first.person, email: first.email, subject: last.subject,
          firstSent: first.sent, lastSent: last.sent, reminders: set.reqs.length - 1,
          replied: !!set.reply, replyAt: set.reply ? set.reply.time : null, replyFrom: set.reply ? set.reply.name : null,
          openId: set.reply ? set.reply.id : last.id, via: first.via
        });
      }
    }

    // incoming: they @mentioned me
    const inn = [];
    const seen = new Map();
    for (const m of mentionedMe) {
      const t = new Date(m.receivedDateTime);
      if (t < cutIn || m.isDraft) continue;
      const email = addr(m.from) || addr(m.sender), name = aname(m.from) || aname(m.sender);
      if (email && (me.addrs.has(email.toLowerCase()) || ignore.has(email.toLowerCase()))) continue;
      if (m.uniqueBody && m.uniqueBody.content) {        // double-check the mention is in THEIR new text
        const ms = getMentions(ownPart(m.uniqueBody.content));
        const meNorm = normName(me.name);
        if (ms.length && !ms.some(x => (x.email && me.addrs.has(x.email.toLowerCase())) || (!x.email && normName(x.name) === meNorm))) continue;
      }
      const conv = m.conversationId || "", topic = topicKey(m.subject);
      const mine = [...(sentByConv[conv] || []), ...(sentByTopic[topic] || [])].filter(s => s.time > t).sort((a, b) => a.time - b.time);
      const rep = mine[0] ? mine[0].time : null;
      const key = topic + "|" + (email || name).toLowerCase() + "|" + (rep ? rep.toISOString() : "OPEN");
      if (seen.has(key)) { const r = seen.get(key); r.reminders++; if (t > r.lastSent) { r.lastSent = t; r.openId = m.id; } if (t < r.firstSent) r.firstSent = t; continue; }
      const row = { dir: "in", person: name, email, subject: m.subject || "", firstSent: t, lastSent: t, reminders: 0,
        replied: !!rep, replyAt: rep, replyFrom: rep ? "You" : null, openId: m.id, via: "@mention" };
      seen.set(key, row); inn.push(row);
    }
    return { out, in: inn };
  }

  function dayCount(a, b, workOnly, weekend) {
    const s = new Date(a), e = new Date(b);
    if (!workOnly) return Math.max(0, Math.floor((e - s) / 864e5));
    const we = new Set(weekend || [5, 6]);
    let n = 0; const d = new Date(s); d.setHours(0, 0, 0, 0); const end = new Date(e); end.setHours(0, 0, 0, 0);
    while (d < end) { d.setDate(d.getDate() + 1); if (!we.has(d.getDay())) n++; }
    return n;
  }

  const api = { normName, samePerson, topicKey, ownPart, getMentions, build, dayCount };
  if (typeof module !== "undefined" && module.exports) module.exports = api; else root.TrackerCore = api;
})(typeof window !== "undefined" ? window : globalThis);
