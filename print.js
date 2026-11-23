/*
 * 답사보고서 PDF(인쇄)용 화면 만들기
 * 입력 데이터(JSON) → 워드 양식과 같은 모양의 HTML
 * 폰의 '인쇄 → PDF로 저장' 기능으로 PDF 파일을 만든다 (별도 라이브러리·글꼴 없이 폰 글꼴 사용)
 * 날짜 함수는 report.js(ReportEngine)의 것을 그대로 씀
 */
(function (root) {
  "use strict";

  /** @param {string} s */
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
    });
  }
  /** **강조** → 빨간 굵은 글씨 @param {string} s @returns {string} */
  function rich(s) {
    return esc(s).replace(/\*\*(.+?)\*\*/g, '<b class="em">$1</b>');
  }
  /** 한 줄 = 한 항목 (report.js와 같은 규칙) @returns {string[]} */
  function lines(v) {
    if (Array.isArray(v)) return v.map(function (x) { return String(x).trim(); }).filter(Boolean);
    return String(v || "").split(/\r?\n/).map(function (x) { return x.replace(/^\s*[-•·*]\s+/, "").trim(); }).filter(Boolean);
  }
  function has(v) { return lines(v).length > 0; }
  function formText(d) {
    var t = Array.isArray(d.broadcastTypes) ? d.broadcastTypes.filter(Boolean) : [];
    var x = String(d.broadcastExtra || "").trim();
    if (!t.length && !x) return d.broadcastForm || "";
    return [t.length ? t.join("·") + " 방송" : "", x].filter(Boolean).join(" / ");
  }
  function netText(d) { return (Array.isArray(d.networks) ? d.networks.filter(Boolean) : []).join(" · "); }
  function ul(v) {
    var ls = lines(v);
    return ls.length ? "<ul>" + ls.map(function (t) { return "<li>" + rich(t) + "</li>"; }).join("") + "</ul>" : "";
  }
  function ps(v) { return lines(v).map(function (t) { return "<p>" + rich(t) + "</p>"; }).join(""); }

  var CSS = [
    "@page{size:A4;margin:20mm}",
    ".rp{font-family:'Malgun Gothic','맑은 고딕','Apple SD Gothic Neo','Noto Sans KR','Noto Sans CJK KR',sans-serif;font-size:9.5pt;line-height:1.45;color:#000;",
    "-webkit-print-color-adjust:exact;print-color-adjust:exact;width:170mm;margin:0 auto;background:#fff}",
    ".rp *{box-sizing:border-box}",
    ".rp p{margin:0}",
    ".rp .em{color:#C00000;font-weight:700}",
    ".rp a{color:#1F5FBF}",
    /* 표지 */
    ".rp .cover{text-align:center;padding-top:12mm;break-after:page;page-break-after:always}",
    ".rp .cover .co{color:#1F3864;font-weight:700;font-size:14pt;margin-bottom:10pt}",
    ".rp .cover .big{font-size:32pt;font-weight:700;color:#111;padding-bottom:6pt;border-bottom:2.25pt solid #1F3864;margin-bottom:15pt}",
    ".rp .cover .csub{font-size:15pt;color:#333;margin:10pt 0 120pt}",
    ".rp .cover table{margin:0 auto;width:325pt;text-align:left}",
    ".rp .cover td{font-size:10.5pt;padding:5.5pt}",
    ".rp .cover th{font-size:10.5pt;width:85pt;padding:5.5pt}",
    /* 본문 머리 */
    ".rp .h0{font-size:13pt;font-weight:700;color:#111;margin-bottom:3pt}",
    ".rp .meta{font-size:8.5pt;color:#666;margin-bottom:3pt}",
    ".rp h2{font-size:11.5pt;color:#1F3864;margin:10pt 0 5pt;padding-bottom:1pt;border-bottom:1pt solid #1F3864;break-after:avoid;page-break-after:avoid}",
    ".rp h2.nl{border-bottom:0;margin-bottom:1pt}",
    /* 표 */
    ".rp table{border-collapse:collapse;width:100%;table-layout:fixed}",
    ".rp td,.rp th{border:.5pt solid #9AA5B8;padding:3.5pt 5.5pt;vertical-align:middle;word-break:keep-all;overflow-wrap:anywhere}",
    ".rp th{background:#E7EDF6;font-weight:700;text-align:center}",
    ".rp thead th{background:#1F3864;color:#fff}",
    ".rp tr{break-inside:avoid;page-break-inside:avoid}",
    ".rp .ov th{width:15.6%}",
    ".rp .cand td{font-size:9pt}",
    ".rp .cand .lab{background:#E7EDF6;text-align:center}",
    ".rp .cand .lab b{display:block;color:#1F3864;font-size:9.5pt}",
    ".rp ul{margin:0;padding-left:10pt}",
    ".rp li{margin:0;padding-left:0}",
    ".rp .summary{font-size:9pt;margin-bottom:4pt}",
    ".rp .op p{margin-bottom:4pt}",
    /* 사진·배치도 */
    ".rp .band{background:#1F3864;color:#fff;font-weight:700;text-align:center;padding:2pt 5.5pt;border:.5pt solid #9AA5B8}",
    ".rp .blk{break-inside:avoid;page-break-inside:avoid}",
    ".rp .row2{display:flex}",
    ".rp .row2>div{flex:1 1 0;min-width:0}",
    ".rp .cellp{border:.5pt solid #9AA5B8;border-top:0;padding:3pt 5.5pt;text-align:center}",
    ".rp .row2>div+div .cellp,.rp .row2>div+div.band{border-left:0}",
    ".rp .blk>.row2:first-child{margin-top:-.5pt}",
    ".rp .blk>.row2:first-child .cellp{border-top:.5pt solid #9AA5B8}",
    ".rp .pic{height:174pt;display:flex;align-items:center;justify-content:center;margin-bottom:2pt}",
    ".rp .pic img{max-width:100%;max-height:100%;display:block}",
    ".rp .cap{font-size:8.5pt}",
    ".rp .dg img{max-width:100%;max-height:200mm;display:block;margin:0 auto}"
  ].join("\n");

  /**
   * @param {object} data 입력 데이터
   * @returns {{html:string, css:string, title:string}}
   */
  function build(data) {
    var R = root.ReportEngine;
    var company = data.company || "MBC씨앤아이";
    var place = data.place || {};
    var title = data.title || ((place.name || "") + " 촬영(중계) 위치 답사");
    var members = [company + " " + (data.members || "")].concat(lines(data.partners));
    var when = R.korDate(data.surveyDate) + (data.surveyTime ? " " + data.surveyTime : "");
    var h = [];

    /* 표지 */
    h.push('<section class="cover"><div class="co">' + esc(company) + '</div><div class="big">답사보고서</div>' +
      '<div class="csub">' + esc(title) + "</div><table>" +
      "<tr><th>일&nbsp;&nbsp;&nbsp;시</th><td>" + esc(R.korDate(data.surveyDate)) + "</td></tr>" +
      "<tr><th>장&nbsp;&nbsp;&nbsp;소</th><td>" + esc(place.name || "") + "</td></tr>" +
      "<tr><th>답사 인원</th><td>" + members.map(function (m) { return "<p>" + esc(m) + "</p>"; }).join("") + "</td></tr>" +
      "</table></section>");

    /* 본문 머리 */
    h.push('<div class="h0">' + esc(title) + '보고서</div><div class="meta">' + esc(company) + " | " + esc(R.dotDate(data.reportDate || data.surveyDate)) + "</div>");
    var n = 0;
    function sec(t, noLine) { n += 1; h.push('<h2' + (noLine ? ' class="nl"' : "") + ">" + n + ". " + esc(t) + "</h2>"); }

    /* 답사 개요 */
    sec("답사 개요");
    var placeHtml = rich((place.name || "") + (place.address ? " (" + place.address + ")" : ""));
    if (place.lat && place.lng) {
      placeHtml += ' <a href="https://map.kakao.com/link/map/' + encodeURIComponent(place.name || "답사지") + "," + place.lat + "," + place.lng + '">카카오맵</a>';
    }
    var contacts = (data.contacts || []).filter(function (c) { return c && (c.name || c.org); }).map(function (c) {
      var who = [c.org, c.name, c.title].filter(Boolean).join(" ");
      var how = [c.phone, c.email].filter(Boolean).join(" / ");
      return who + (how ? " (" + how + ")" : "");
    });
    var rows = [
      ["일시", [when]], ["장소", null, placeHtml], ["답사 인원", [members.join(" / ")]], ["현장 담당자", contacts],
      ["프로그램", [data.program]], ["방송일", [R.rangeDate(data.broadcastStart, data.broadcastEnd)]],
      ["목적", [data.purpose]], ["중계 형태", [formText(data)]], ["사용 망", [netText(data)]]
    ];
    h.push('<table class="ov">' + rows.map(function (r) {
      var body = r[2] ? "<p>" + r[2] + "</p>" : (r[1] || []).filter(function (x) { return x && String(x).trim(); }).map(function (t) { return "<p>" + rich(t) + "</p>"; }).join("");
      return body ? "<tr><th>" + r[0] + "</th><td>" + body + "</td></tr>" : "";
    }).join("") + "</table>");

    /* 후보지 */
    var cands = (data.candidates || []).filter(function (c) { return c && (c.name || has(c.location) || has(c.pros) || has(c.cons)); });
    function lab(c, i) { return c.label || (i + 1) + "안"; }
    if (cands.length) {
      sec(cands.length > 1 ? "후보지별 검토" : "현장 검토");
      if (data.candidateSummary) h.push('<p class="summary">' + rich(data.candidateSummary) + "</p>");
      h.push('<table class="cand"><colgroup><col style="width:13%"><col style="width:24.4%"><col style="width:28%"><col style="width:34.6%"></colgroup>' +
        "<thead><tr><th>구분</th><th>위치·환경</th><th>장점</th><th>고려사항</th></tr></thead><tbody>" +
        cands.map(function (c, i) {
          return '<tr><td class="lab"><b>' + esc(lab(c, i)) + "</b>" + esc((c.name || "") + (c.planned ? "(예정)" : "")) + "</td>" +
            "<td>" + ps(c.location) + "</td><td>" + ul(c.pros) + "</td><td>" + ul(c.cons) + "</td></tr>";
        }).join("") + "</tbody></table>");
    }

    /* 기술 검토 사항 */
    var tech = (data.tech || []).filter(function (t) { return t && has(t.items); });
    if (tech.length) {
      sec("기술 검토 사항");
      h.push('<table class="ov">' + tech.map(function (t) { return "<tr><th>" + esc(t.category || "") + "</th><td>" + ul(t.items) + "</td></tr>"; }).join("") + "</table>");
    }

    /* 종합 의견 */
    if (has(data.opinion)) { sec("종합 의견"); h.push('<div class="op">' + ps(data.opinion) + "</div>"); }

    /* 위치도·배치도 */
    var maps = (data.diagrams || []).filter(function (d) { return d && d.data; });
    if (maps.length) {
      sec("위치도·배치도", true);
      maps.forEach(function (d) {
        h.push('<div class="blk dg"><div class="band">' + esc(d.caption || "배치도") + '</div><div class="cellp"><img src="' + d.data + '" alt=""></div></div>');
      });
    }

    /* 현장 사진 — 워드와 같은 묶음 규칙 (사진 1장짜리 후보지 2곳은 나란히) */
    var groups = [];
    cands.forEach(function (c, i) {
      var ph = (c.photos || []).filter(function (p) { return p && p.data; });
      if (ph.length) groups.push({ title: lab(c, i) + (c.planned ? "(예정)" : "") + " – " + (c.photoTitle || c.name || ""), photos: ph });
    });
    var extra = (data.extraPhotos || []).filter(function (p) { return p && p.data; });
    if (extra.length) groups.push({ title: data.extraPhotoTitle || "기타 현장 사진", photos: extra });
    if (groups.length) {
      sec("현장 사진", true);
      var num = 0;
      var cell = function (p) {
        if (!p) return '<div><div class="cellp"><div class="pic"></div><div class="cap">&nbsp;</div></div></div>';
        num += 1;
        return '<div><div class="cellp"><div class="pic"><img src="' + p.data + '" alt=""></div><div class="cap">사진 ' + num + ". " + esc(p.caption || "") + "</div></div></div>";
      };
      var i = 0;
      while (i < groups.length) {
        var g = groups[i], nx = groups[i + 1];
        if (g.photos.length === 1 && nx && nx.photos.length === 1) {
          h.push('<div class="blk"><div class="row2"><div class="band">' + esc(g.title) + '</div><div class="band">' + esc(nx.title) + "</div></div>" +
            '<div class="row2">' + cell(g.photos[0]) + cell(nx.photos[0]) + "</div></div>");
          i += 2; continue;
        }
        for (var k = 0; k < g.photos.length; k += 2) {
          // 제목줄은 첫 줄 사진과 한 덩어리로 묶어 페이지 끝에 혼자 남지 않게
          h.push('<div class="blk">' + (k === 0 ? '<div class="band">' + esc(g.title) + "</div>" : "") +
            '<div class="row2">' + cell(g.photos[k]) + cell(g.photos[k + 1]) + "</div></div>");
        }
        i += 1;
      }
    }

    return { html: '<div class="rp">' + h.join("") + "</div>", css: CSS, title: R.fileName(data).replace(/\.docx$/, "") };
  }

  var api = { build: build };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.ReportPrint = api;
})(typeof window !== "undefined" ? window : this);
