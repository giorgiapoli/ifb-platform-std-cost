"""
Tabelle HTML ridimensionabili: l'utente trascina l'angolo in basso a destra per cambiare larghezza e altezza della tabella,
e il bordo destro di ogni intestazione per cambiare la larghezza della colonna. Le misure si ricordano (localStorage del
browser) e vengono riapplicate a ogni aggiornamento della pagina, tabella per tabella.

Le tabelle si riconoscono dalle classi CSS 'tab-h' e 'tab-det' e dal testo delle intestazioni. Lo script viene inserito
nella pagina principale (non nel riquadro del componente), cosi' continua a funzionare quando Streamlit ridisegna la pagina.
"""
import streamlit.components.v1 as components

_JS = r"""
(function () {
  var P = window.parent, D = P.document;
  if (D.getElementById("stdc-resize-script")) return;
  var corpo = function () {
    if (window.__stdcResize) return; window.__stdcResize = true;
    var KEY = "stdc_tabelle_v1";
    var leggi = function () { try { return JSON.parse(localStorage.getItem(KEY) || "{}"); } catch (e) { return {}; } };
    var scrivi = function (o) { try { localStorage.setItem(KEY, JSON.stringify(o)); } catch (e) {} };
    var chiave = function (box) {
      return Array.prototype.map.call(box.querySelectorAll("thead th"), function (t) { return (t.innerText || "").trim(); }).join("|").slice(0, 300);
    };
    var colgroup = function (box) {
      var table = box.querySelector("table"), ths = box.querySelectorAll("thead th");
      var cg = table.querySelector("colgroup");
      if (!cg) {
        cg = document.createElement("colgroup");
        Array.prototype.forEach.call(ths, function (th) {
          var c = document.createElement("col"); c.style.width = th.offsetWidth + "px"; cg.appendChild(c);
        });
        table.insertBefore(cg, table.firstChild);
      }
      table.style.tableLayout = "fixed";
      var tot = 0; Array.prototype.forEach.call(cg.children, function (c) { tot += parseFloat(c.style.width) || 0; });
      table.style.width = tot + "px"; table.style.maxWidth = "none";
      return cg;
    };
    var prepara = function (box) {
      if (box.__stdc) return; box.__stdc = true;
      var k = chiave(box); if (!k) return;
      var tutti = leggi(), st = tutti[k] || {};
      if (st.w) box.style.width = st.w + "px";
      if (st.h) box.style.height = st.h + "px";
      var ths = box.querySelectorAll("thead th");
      if (st.cols && st.cols.length === ths.length) {
        var cg = colgroup(box);
        st.cols.forEach(function (w, i) { cg.children[i].style.width = w + "px"; });
        colgroup(box);
      }
      // maniglie per cambiare la larghezza delle colonne
      Array.prototype.forEach.call(ths, function (th, i) {
        th.style.position = "sticky"; th.style.overflow = "visible";
        var m = document.createElement("div");
        m.style.cssText = "position:absolute;top:0;right:-3px;width:7px;height:100%;cursor:col-resize;z-index:5;user-select:none";
        m.onmouseenter = function () { m.style.background = "rgba(15,118,110,.35)"; };
        m.onmouseleave = function () { m.style.background = "transparent"; };
        m.onmousedown = function (ev) {
          ev.preventDefault(); ev.stopPropagation();
          var cg2 = colgroup(box), col = cg2.children[i], x0 = ev.clientX, w0 = parseFloat(col.style.width) || th.offsetWidth;
          var muovi = function (e) { col.style.width = Math.max(30, w0 + e.clientX - x0) + "px"; colgroup(box); };
          var rilascia = function () {
            document.removeEventListener("mousemove", muovi); document.removeEventListener("mouseup", rilascia);
            var t = leggi(), s = t[k] || {};
            s.cols = Array.prototype.map.call(cg2.children, function (c) { return parseFloat(c.style.width) || 0; });
            t[k] = s; scrivi(t);
          };
          document.addEventListener("mousemove", muovi); document.addEventListener("mouseup", rilascia);
        };
        th.appendChild(m);
      });
      // dimensione dell'area (trascinando l'angolo in basso a destra)
      var primo = true;
      new ResizeObserver(function () {
        if (primo) { primo = false; return; }
        var t = leggi(), s = t[k] || {};
        s.w = Math.round(box.offsetWidth); s.h = Math.round(box.offsetHeight); t[k] = s; scrivi(t);
      }).observe(box);
    };
    var scansiona = function () {
      Array.prototype.forEach.call(document.querySelectorAll(".tab-h, .tab-det"), prepara);
    };
    new MutationObserver(function () { clearTimeout(window.__stdcT); window.__stdcT = setTimeout(scansiona, 60); })
      .observe(document.body, { childList: true, subtree: true });
    scansiona();
  };
  var s = D.createElement("script"); s.id = "stdc-resize-script";
  s.text = "(" + corpo.toString() + ")();";
  D.head.appendChild(s);
})();
"""


def inietta():
    """Da chiamare una volta per esecuzione, in cima alla pagina: non occupa spazio."""
    components.html("<script>" + _JS + "</script>", height=0)
