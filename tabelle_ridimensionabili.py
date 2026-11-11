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
      var larghezzaPagina = function () { var p = box.parentElement; return (p && p.clientWidth) || window.innerWidth; };
      box.style.maxWidth = "100%";                       // mai piu' larga dello spazio disponibile (cambio di schermo)
      if (st.wf) box.style.width = Math.min(larghezzaPagina(), Math.round(st.wf * larghezzaPagina())) + "px";
      else if (st.w) box.style.width = Math.min(larghezzaPagina(), st.w) + "px";
      if (st.h) box.style.height = st.h + "px";
      var ths = box.querySelectorAll("thead th");
      var n = ths.length;
      var ord = (st.ord && st.ord.length === n) ? st.ord.slice() : Array.apply(null, Array(n)).map(function (_, i) { return i; });
      // sposta la colonna da posizione i a posizione j in intestazione, righe e colgroup
      var muoviColonna = function (i, j) {
        if (i === j) return;
        var righe = box.querySelectorAll("tr");
        Array.prototype.forEach.call(righe, function (tr) {
          var c = Array.prototype.slice.call(tr.children);
          if (c.length !== n) return;
          tr.insertBefore(c[i], j > i ? c[j].nextSibling : c[j]);
        });
        var cg0 = box.querySelector("colgroup");
        if (cg0 && cg0.children.length === n) {
          var cc = Array.prototype.slice.call(cg0.children);
          cg0.insertBefore(cc[i], j > i ? cc[j].nextSibling : cc[j]);
        }
      };
      if (st.ord && st.ord.length === n) {
        // riapplico l'ordine salvato: porto ogni colonna originale al suo posto
        var attuale = Array.apply(null, Array(n)).map(function (_, i) { return i; });
        ord.forEach(function (orig, pos) {
          var da = attuale.indexOf(orig);
          if (da !== pos) { muoviColonna(da, pos); attuale.splice(pos, 0, attuale.splice(da, 1)[0]); }
        });
        ths = box.querySelectorAll("thead th");
      }
      if (st.cols && st.cols.length === ths.length) {
        var cg = colgroup(box);
        st.cols.forEach(function (w, i) { cg.children[i].style.width = w + "px"; });
        colgroup(box);
      }
      // maniglie per cambiare la larghezza delle colonne
      Array.prototype.forEach.call(ths, function (th) {
        th.style.position = "sticky"; th.style.overflow = "visible";
        var pos = function () { return Array.prototype.indexOf.call(th.parentNode.children, th); };
        // trascinando un'intestazione su un'altra, la colonna si sposta (l'ordine si ricorda)
        th.draggable = true;
        th.ondragstart = function (ev) { ev.dataTransfer.setData("text/plain", String(pos())); ev.dataTransfer.effectAllowed = "move"; th.style.opacity = ".5"; };
        th.ondragend = function () { th.style.opacity = ""; };
        th.ondragover = function (ev) { ev.preventDefault(); th.style.boxShadow = "inset 0 -3px 0 #0F766E"; };
        th.ondragleave = function () { th.style.boxShadow = ""; };
        th.ondrop = function (ev) {
          ev.preventDefault(); th.style.boxShadow = "";
          var da = parseInt(ev.dataTransfer.getData("text/plain"), 10), a = pos();
          if (isNaN(da) || da === a) return;
          muoviColonna(da, a);
          ord.splice(a, 0, ord.splice(da, 1)[0]);
          var t = leggi(), s = t[k] || {};
          if (s.cols && s.cols.length === n) s.cols.splice(a, 0, s.cols.splice(da, 1)[0]);
          s.ord = ord.slice(); t[k] = s; scrivi(t);
        };
        var m = document.createElement("div");
        m.style.cssText = "position:absolute;top:0;right:-3px;width:7px;height:100%;cursor:col-resize;z-index:5;user-select:none";
        m.onmouseenter = function () { m.style.background = "rgba(15,118,110,.35)"; };
        m.onmouseleave = function () { m.style.background = "transparent"; };
        m.onmousedown = function (ev) {
          ev.preventDefault(); ev.stopPropagation();
          var cg2 = colgroup(box), col = cg2.children[pos()], x0 = ev.clientX, w0 = parseFloat(col.style.width) || th.offsetWidth;
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
        s.w = Math.round(box.offsetWidth); s.h = Math.round(box.offsetHeight);
        s.wf = Math.min(1, box.offsetWidth / ((box.parentElement && box.parentElement.clientWidth) || window.innerWidth));
        t[k] = s; scrivi(t);
      }).observe(box);
    };
    var scansiona = function () {
      Array.prototype.forEach.call(document.querySelectorAll(".tab-h, .tab-det"), prepara);
    };
    new MutationObserver(function () { clearTimeout(window.__stdcT); window.__stdcT = setTimeout(scansiona, 60); })
      .observe(document.body, { childList: true, subtree: true });
    // clic sul codice articolo di una tabella con data-apri: scrive il codice nel campo nascosto di Streamlit (apre lo storico)
    document.addEventListener("click", function (e) {
      var td = e.target.closest ? e.target.closest("td.col0") : null;
      if (!td) return;
      var box = td.closest("[data-apri]"); if (!box) return;
      var inp = document.querySelector(".st-key-" + box.getAttribute("data-apri") + " input"); if (!inp) return;
      var set = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value").set;
      inp.focus(); set.call(inp, td.innerText.trim() + "|" + Date.now());
      inp.dispatchEvent(new Event("input", { bubbles: true }));
      ["keydown", "keypress", "keyup"].forEach(function (t) {          // invio + uscita dal campo: Streamlit conferma il valore
        inp.dispatchEvent(new KeyboardEvent(t, { key: "Enter", code: "Enter", keyCode: 13, which: 13, charCode: 13, bubbles: true }));
      });
      inp.dispatchEvent(new FocusEvent("focusout", { bubbles: true })); inp.blur();
    }, true);
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
