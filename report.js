/*
 * 답사보고서 워드 변환 엔진
 * 입력 데이터(JSON) → MBC씨앤아이 답사보고서 양식(.docx)
 * 브라우저(window.docx)와 Node(require('docx')) 양쪽에서 동작
 */
(function (root) {
  "use strict";

  var NAVY = "1F3864", LIGHT = "E7EDF6", BORDER = "9AA5B8", RED = "C00000";
  var DARK = "111111", GRAY = "333333", META = "666666";
  var FONT = { ascii: "맑은 고딕", eastAsia: "맑은 고딕", hAnsi: "맑은 고딕", cs: "맑은 고딕" };
  var FULL = 9638; // A4, 여백 2cm 기준 본문 폭(DXA)
  var WEEK = ["일", "월", "화", "수", "목", "금", "토"];

  /* ---------- 날짜 ---------- */
  function parseDate(s) {
    if (!s) return null;
    var m = String(s).match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (!m) return null;
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }
  function korDate(s) {
    var d = parseDate(s);
    if (!d) return s || "";
    return d.getFullYear() + "년 " + (d.getMonth() + 1) + "월 " + d.getDate() + "일(" + WEEK[d.getDay()] + ")";
  }
  function dotDate(s) {
    var d = parseDate(s);
    if (!d) return s || "";
    var p = function (n) { return n < 10 ? "0" + n : "" + n; };
    return d.getFullYear() + ". " + (d.getMonth() + 1) + ". " + p(d.getDate()) + ".";
  }
  function rangeDate(a, b) {
    if (!a) return "";
    if (!b || a === b) return korDate(a);
    var da = parseDate(a), db = parseDate(b);
    if (!da || !db) return korDate(a) + " ~ " + korDate(b);
    if (da.getFullYear() === db.getFullYear() && da.getMonth() === db.getMonth())
      return korDate(a) + " ~ " + db.getDate() + "일(" + WEEK[db.getDay()] + ")";
    if (da.getFullYear() === db.getFullYear())
      return korDate(a) + " ~ " + (db.getMonth() + 1) + "월 " + db.getDate() + "일(" + WEEK[db.getDay()] + ")";
    return korDate(a) + " ~ " + korDate(b);
  }

  /* ---------- 텍스트 ---------- */
  function lines(v) {
    if (Array.isArray(v)) return v.map(function (x) { return String(x).trim(); }).filter(Boolean);
    return String(v || "").split(/\r?\n/).map(function (x) { return x.replace(/^\s*[-•·*]\s+/, "").trim(); }).filter(Boolean);
  }
  function has(v) { return lines(v).length > 0; }

  /* 중계 형태: ["생방송","SNG"] + "카메라 3대" → "생방송·SNG 방송 / 카메라 3대"
     (예전 저장본의 broadcastForm 한 줄도 그대로 사용) */
  function formText(data) {
    var types = Array.isArray(data.broadcastTypes) ? data.broadcastTypes.filter(Boolean) : [];
    var extra = String(data.broadcastExtra || "").trim();
    if (!types.length && !extra) return data.broadcastForm || "";
    return [types.length ? types.join("·") + " 방송" : "", extra].filter(Boolean).join(" / ");
  }
  function netText(data) {
    return (Array.isArray(data.networks) ? data.networks.filter(Boolean) : []).join(" · ");
  }

  function decodeImage(dataUrl) {
    var b64 = String(dataUrl).split(",").pop();
    var bin = atob(b64), out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }
  function imgType(dataUrl) {
    var m = String(dataUrl).match(/^data:image\/(\w+)/);
    var t = m ? m[1].toLowerCase() : "jpg";
    return t === "jpeg" ? "jpg" : t;
  }

  function build(D, data) {
    var P = D.Paragraph, T = D.TextRun, A = D.AlignmentType, W = D.WidthType;

    /* **강조** → 빨간 굵은 글씨 */
    function rich(text, base) {
      base = base || {};
      var parts = String(text || "").split(/\*\*(.+?)\*\*/g);
      var runs = [];
      parts.forEach(function (s, i) {
        if (!s) return;
        var o = { text: s };
        Object.keys(base).forEach(function (k) { o[k] = base[k]; });
        if (i % 2 === 1) { o.bold = true; o.color = RED; }
        runs.push(new T(o));
      });
      if (!runs.length) runs.push(new T(Object.assign({ text: "" }, base)));
      return runs;
    }
    function para(text, opt) {
      opt = opt || {};
      return new P({
        alignment: opt.align,
        spacing: opt.spacing,
        keepNext: opt.keepNext,
        border: opt.border,
        numbering: opt.bullet ? { reference: "dot", level: 0 } : undefined,
        children: opt.runs || rich(text, opt.run),
      });
    }

    var bd = { style: D.BorderStyle.SINGLE, size: 4, color: BORDER };
    var borders = { top: bd, bottom: bd, left: bd, right: bd };
    function cell(children, o) {
      o = o || {};
      var m = o.margins || [70, 110];
      return new D.TableCell({
        width: { size: o.width, type: W.DXA },
        columnSpan: o.span,
        borders: borders,
        shading: o.fill ? { type: D.ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
        margins: { top: m[0], bottom: m[0], left: m[1], right: m[1] },
        verticalAlign: D.VerticalAlign.CENTER,
        children: children,
      });
    }
    function table(widths, rows, o) {
      o = o || {};
      return new D.Table({
        width: { size: widths.reduce(function (a, b) { return a + b; }, 0), type: W.DXA },
        columnWidths: widths,
        layout: D.TableLayoutType.FIXED,
        alignment: o.center ? A.CENTER : undefined,
        indent: o.center ? undefined : { size: 110, type: W.DXA },
        rows: rows,
      });
    }
    function labelCell(text, width, size, margins) {
      return cell([para(text, { align: A.CENTER, run: { bold: true, size: size } })], { width: width, fill: LIGHT, margins: margins });
    }
    function heading(text, opt) {
      opt = opt || {};
      return new P({
        keepNext: true,
        border: opt.noLine ? undefined : { bottom: { style: D.BorderStyle.SINGLE, size: 8, space: 2, color: NAVY } },
        spacing: { before: 200, after: opt.noLine ? 20 : 100 },
        children: [new T({ text: text, bold: true, color: NAVY, size: 23 })],
      });
    }
    function image(dataUrl, w, h) {
      return new D.ImageRun({ type: imgType(dataUrl), data: decodeImage(dataUrl), transformation: { width: w, height: h } });
    }

    var company = data.company || "MBC씨앤아이";
    var place = data.place || {};
    var title = data.title || ((place.name || "") + " 촬영(중계) 위치 답사");
    var members = [company + " " + (data.members || "")].concat(lines(data.partners));
    var whenText = korDate(data.surveyDate) + (data.surveyTime ? " " + data.surveyTime : "");
    var kids = [];

    /* ================= 표지 ================= */
    kids.push(new P({ children: [] }));
    kids.push(new P({ alignment: A.CENTER, spacing: { after: 200 }, children: [new T({ text: company, bold: true, color: NAVY, size: 28 })] }));
    kids.push(new P({
      alignment: A.CENTER, spacing: { after: 300 },
      border: { bottom: { style: D.BorderStyle.SINGLE, size: 18, space: 12, color: NAVY } },
      children: [new T({ text: "답사보고서", bold: true, color: DARK, size: 64 })],
    }));
    kids.push(new P({ alignment: A.CENTER, spacing: { before: 200, after: 2400 }, children: [new T({ text: title, color: GRAY, size: 30 })] }));
    var coverRows = [
      ["일   시", [korDate(data.surveyDate)]],
      ["장   소", [place.name || ""]],
      ["답사 인원", members],
    ].map(function (r) {
      return new D.TableRow({
        children: [
          labelCell(r[0], 1700, 21, [110, 110]),
          cell(r[1].map(function (t) { return para(t, { run: { size: 21 } }); }), { width: 4800, margins: [110, 110] }),
        ],
      });
    });
    kids.push(table([1700, 4800], coverRows, { center: true }));
    kids.push(new P({ spacing: { after: 40 }, children: [new D.PageBreak()] }));

    /* ================= 본문 머리 ================= */
    kids.push(new P({ spacing: { after: 60 }, children: [new T({ text: title + "보고서", bold: true, color: DARK, size: 26 })] }));
    kids.push(new P({ spacing: { after: 60 }, children: [new T({ text: company + " | " + dotDate(data.reportDate || data.surveyDate), color: META, size: 17 })] }));

    var n = 0;
    function sec(t, o) { n += 1; kids.push(heading(n + ". " + t, o)); }

    /* ================= 1. 답사 개요 ================= */
    sec("답사 개요");
    var placeRuns = rich((place.name || "") + (place.address ? " (" + place.address + ")" : ""));
    if (place.lat && place.lng) {
      placeRuns.push(new T({ text: "  " }));
      placeRuns.push(new D.ExternalHyperlink({
        link: "https://map.kakao.com/link/map/" + encodeURIComponent(place.name || "답사지") + "," + place.lat + "," + place.lng,
        children: [new T({ text: "카카오맵", color: "1F5FBF", underline: {}, size: 17 })],
      }));
    }
    var contacts = (data.contacts || []).filter(function (c) { return c && (c.name || c.org); }).map(function (c) {
      var who = [c.org, c.name, c.title].filter(Boolean).join(" ");
      var how = [c.phone, c.email].filter(Boolean).join(" / ");
      return who + (how ? " (" + how + ")" : "");
    });
    var ov = [
      ["일시", [whenText]],
      ["장소", null, placeRuns],
      ["답사 인원", [members.join(" / ")]],
      ["현장 담당자", contacts],
      ["프로그램", [data.program]],
      ["방송일", [rangeDate(data.broadcastStart, data.broadcastEnd)]],
      ["목적", [data.purpose]],
      ["중계 형태", [formText(data)]],
      ["사용 망", [netText(data)]],
    ];
    var ovRows = [];
    ov.forEach(function (r) {
      var body;
      if (r[2]) body = [new P({ children: r[2] })];
      else {
        var ls = (r[1] || []).filter(function (x) { return x && String(x).trim(); });
        if (!ls.length) return;
        body = ls.map(function (t) { return para(t); });
      }
      ovRows.push(new D.TableRow({ children: [labelCell(r[0], 1500), cell(body, { width: 8138 })] }));
    });
    kids.push(table([1500, 8138], ovRows));

    /* ================= 2. 후보지별 검토 ================= */
    var cands = (data.candidates || []).filter(function (c) { return c && (c.name || has(c.location) || has(c.pros) || has(c.cons)); });
    function candLabel(c, i) { return c.label || (i + 1) + "안"; }
    if (cands.length) {
      sec(cands.length > 1 ? "후보지별 검토" : "현장 검토");
      if (data.candidateSummary) kids.push(para(data.candidateSummary, { spacing: { after: 80 }, run: { size: 18 } }));
      var hdr = ["구분", "위치·환경", "장점", "고려사항"], cw = [1250, 2350, 2700, 3338];
      var rows = [new D.TableRow({
        tableHeader: true,
        children: hdr.map(function (h, i) {
          return cell([para(h, { align: A.CENTER, run: { bold: true, color: "FFFFFF" } })], { width: cw[i], fill: NAVY });
        }),
      })];
      cands.forEach(function (c, i) {
        var small = { size: 18 };
        var bul = function (v) { var ls = lines(v); return ls.length ? ls.map(function (t) { return para(t, { bullet: true, run: small }); }) : [para("")]; };
        var plain = function (v) { var ls = lines(v); return ls.length ? ls.map(function (t) { return para(t, { run: small }); }) : [para("")]; };
        rows.push(new D.TableRow({
          cantSplit: true,
          children: [
            cell([
              para(candLabel(c, i), { align: A.CENTER, run: { bold: true, color: NAVY } }),
              para((c.name || "") + (c.planned ? "(예정)" : ""), { align: A.CENTER, run: small }),
            ], { width: cw[0], fill: LIGHT }),
            cell(plain(c.location), { width: cw[1] }),
            cell(bul(c.pros), { width: cw[2] }),
            cell(bul(c.cons), { width: cw[3] }),
          ],
        }));
      });
      kids.push(table(cw, rows));
    }

    /* ================= 기술 검토 사항 ================= */
    var tech = (data.tech || []).filter(function (t) { return t && has(t.items); });
    if (tech.length) {
      sec("기술 검토 사항");
      kids.push(table([1500, 8138], tech.map(function (t) {
        return new D.TableRow({
          cantSplit: true,
          children: [labelCell(t.category || "", 1500), cell(lines(t.items).map(function (x) { return para(x, { bullet: true }); }), { width: 8138 })],
        });
      })));
    }

    /* ================= 종합 의견 ================= */
    if (has(data.opinion)) {
      sec("종합 의견");
      lines(data.opinion).forEach(function (t) { kids.push(para(t, { spacing: { after: 80 } })); });
    }

    /* ================= 위치도·배치도 ================= */
    var maps = (data.diagrams || []).filter(function (d) { return d && d.data; });
    if (maps.length) {
      sec("위치도·배치도", { noLine: true });
      var mrows = [];
      maps.forEach(function (d) {
        var maxW = 610, maxH = 760;
        var w = d.w || 4, h = d.h || 3, s = Math.min(maxW / w, maxH / h);
        mrows.push(new D.TableRow({ cantSplit: true, children: [cell([para(d.caption || "배치도", { align: A.CENTER, keepNext: true, run: { bold: true, color: "FFFFFF" } })], { width: FULL, fill: NAVY, margins: [40, 110] })] }));
        mrows.push(new D.TableRow({ cantSplit: true, children: [cell([new P({ alignment: A.CENTER, children: [image(d.data, Math.round(w * s), Math.round(h * s))] })], { width: FULL, margins: [60, 110] })] }));
      });
      kids.push(table([FULL], mrows));
    }

    /* ================= 현장 사진 ================= */
    var groups = [];
    cands.forEach(function (c, i) {
      var ph = (c.photos || []).filter(function (p) { return p && p.data; });
      if (!ph.length) return;
      var lab = candLabel(c, i) + (c.planned ? "(예정)" : "");
      groups.push({ title: lab + " – " + (c.photoTitle || c.name || ""), photos: ph });
    });
    var extra = (data.extraPhotos || []).filter(function (p) { return p && p.data; });
    if (extra.length) groups.push({ title: data.extraPhotoTitle || "기타 현장 사진", photos: extra });

    if (groups.length) {
      sec("현장 사진", { noLine: true });
      var HALF = 4819, num = 0, prow = [];
      var band = function (t, w, span) {
        return cell([para(t, { align: A.CENTER, keepNext: true, run: { bold: true, color: "FFFFFF" } })], { width: w, span: span, fill: NAVY, margins: [40, 110] });
      };
      var photoCell = function (p) {
        if (!p) return cell([para("")], { width: HALF, margins: [60, 110] });
        num += 1;
        // 가로 사진은 310×232(4:3) 칸에 꽉 차게, 세로 사진은 높이에 맞춰 전체가 보이게
        var pw = 310, ph = 232;
        if (p.w && p.h && p.h > p.w) { pw = Math.round(232 * p.w / p.h); }
        return cell([
          new P({ alignment: A.CENTER, spacing: { after: 40 }, children: [image(p.data, pw, ph)] }),
          para("사진 " + num + ". " + (p.caption || ""), { align: A.CENTER, run: { size: 17 } }),
        ], { width: HALF, margins: [60, 110] });
      };
      var i = 0;
      while (i < groups.length) {
        var g = groups[i], nx = groups[i + 1];
        if (g.photos.length === 1 && nx && nx.photos.length === 1) {
          // 사진 1장짜리 후보지 두 곳은 한 줄에 나란히
          prow.push(new D.TableRow({ cantSplit: true, children: [band(g.title, HALF), band(nx.title, HALF)] }));
          prow.push(new D.TableRow({ cantSplit: true, children: [photoCell(g.photos[0]), photoCell(nx.photos[0])] }));
          i += 2;
          continue;
        }
        prow.push(new D.TableRow({ cantSplit: true, children: [band(g.title, FULL, 2)] }));
        for (var k = 0; k < g.photos.length; k += 2) {
          prow.push(new D.TableRow({ cantSplit: true, children: [photoCell(g.photos[k]), photoCell(g.photos[k + 1])] }));
        }
        i += 1;
      }
      kids.push(table([HALF, HALF], prow));
    }

    return new D.Document({
      creator: company,
      title: title + "보고서",
      styles: { default: { document: { run: { font: FONT, size: 19 } } } },
      numbering: {
        config: [{
          reference: "dot",
          levels: [{
            level: 0, format: D.LevelFormat.BULLET, text: "•", alignment: A.LEFT,
            style: { paragraph: { indent: { left: 260, hanging: 200 } } },
          }],
        }],
      },
      sections: [{
        properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } } },
        children: kids,
      }],
    });
  }

  function fileName(data) {
    var t = (data.title || ((data.place && data.place.name) || "") + " 촬영(중계) 위치 답사").replace(/[\\/:*?"<>|]/g, "");
    var d = (data.surveyDate || "").replace(/-/g, "");
    return "답사보고서_" + t + (d ? "_" + d : "") + ".docx";
  }

  var api = { build: build, fileName: fileName, korDate: korDate, dotDate: dotDate, rangeDate: rangeDate };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.ReportEngine = api;
})(typeof window !== "undefined" ? window : this);
