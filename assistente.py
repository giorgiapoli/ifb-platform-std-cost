"""
Suggerimento «assistente»: passando il mouse (o cliccando) su un campo compare un fumetto con un avatar che dice a cosa serve
il campo e cosa scrivere. Solo grafica: non cambia come funziona l'app. Resta fermo sotto il campo e si chiude con la X.

Uso: dare una `key` al widget e chiamare `bolla(key, ...)` subito DOPO il widget.
"""
import streamlit as st
import streamlit.components.v1 as components

# Il fumetto compare sotto il campo passandoci sopra (dopo un attimo) e STA FERMO li': non segue il mouse. Si chiude con la X
# (e allora non ricompare finche' non si ricarica la pagina); se ci si sposta resta ancora qualche secondo e poi sparisce da solo;
# sparisce anche scrivendo. Si puo' spegnere del tutto dalla barra laterale.
_JS = r"""
(function () {
  var P = window.parent, D = P.document;
  if (D.getElementById("stdc-ai-v3")) return;
  var corpo = function () {
    if (window.__stdcAi3) return; window.__stdcAi3 = true;
    var cur = null, att = null, tA = null, tH = null, chiusi = {};
    var nascondi = function () { clearTimeout(tH); if (att) { att.classList.remove("ai-on"); att = null; } };
    document.addEventListener("mouseover", function (e) {
      var w = e.target.closest ? e.target.closest('[class*="st-key-"]') : null;
      if (w === cur) return;
      clearTimeout(tA); cur = w;
      if (!w) return;
      var b = w.nextElementSibling ? w.nextElementSibling.querySelector(".ai-bolla") : null;
      if (!b) return;
      var id = Array.prototype.find.call(w.classList, function (c) { return c.indexOf("st-key-") === 0; });
      if (!id || chiusi[id]) return;
      tA = setTimeout(function () {
        if (att && att !== b) { att.classList.remove("ai-on"); }
        clearTimeout(tH);
        b.setAttribute("data-id", id); b.classList.add("ai-on"); att = b;
      }, 250);
    }, true);
    document.addEventListener("mouseout", function (e) {
      if (cur && !(e.relatedTarget && cur.contains(e.relatedTarget))) {
        clearTimeout(tA); cur = null;
        if (att) { clearTimeout(tH); tH = setTimeout(nascondi, 6000); }       // resta qualche secondo, poi sparisce
      }
    }, true);
    document.addEventListener("click", function (e) {
      if (e.target.classList && e.target.classList.contains("ai-x")) {
        var b = e.target.closest(".ai-bolla");
        if (b) { chiusi[b.getAttribute("data-id")] = true; b.classList.remove("ai-on"); if (att === b) att = null; }
        e.stopPropagation(); e.preventDefault();
      }
    }, true);
    document.addEventListener("keydown", nascondi, true);
  };
  var s = D.createElement("script"); s.id = "stdc-ai-v3";
  s.text = "(" + corpo.toString() + ")();";
  D.head.appendChild(s);
})();
"""


def inietta():
    """Da chiamare una volta per esecuzione, in cima alla pagina: non occupa spazio."""
    components.html("<script>" + _JS + "</script>", height=0)


_AVATAR = (
    '<svg class="ai-avatar" viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">'
    '<defs><linearGradient id="aig" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#14B8A6"/>'
    '<stop offset="1" stop-color="#0F766E"/></linearGradient></defs>'
    '<circle cx="24" cy="24" r="23" fill="url(#aig)"/>'
    '<g class="ai-occhi" fill="#fff"><ellipse cx="17" cy="21" rx="3" ry="3.6"/><ellipse cx="31" cy="21" rx="3" ry="3.6"/></g>'
    '<path d="M15 30 Q24 38 33 30" fill="none" stroke="#fff" stroke-width="2.6" stroke-linecap="round"/>'
    '<circle cx="11" cy="28" r="3" fill="#fff" opacity=".18"/><circle cx="37" cy="28" r="3" fill="#fff" opacity=".18"/>'
    '</svg>')


def bolla(chiave, titolo, testo, esempi=(), larghezza=620, destra=False):
    """Fumetto che compare vicino al puntatore del mouse passando sopra al widget con la `chiave` indicata (va chiamata subito dopo il widget)."""
    if not st.session_state.get("assistente_on", True):          # spento dalla barra laterale
        return
    chip = "".join(f'<span class="ai-chip">{e}</span>' for e in esempi)
    righe = f'<div class="ai-esempi">{chip}</div>' if chip else ""
    k = f".st-key-{chiave}"
    lato = "right:14px !important;left:auto !important" if destra else "left:14px !important;right:auto !important"
    freccia = "right:34px" if destra else "left:34px"
    css = f"""<style>
.ai-bolla.ai-on{{opacity:1;visibility:visible}}
.ai-wrap{{position:relative;height:0;margin:-6px 0 0 0}}
.ai-bolla{{position:absolute;top:2px !important;{lato};z-index:60;width:min({min(larghezza, 400)}px,90vw);display:flex;gap:12px;align-items:flex-start;
  background:#fff;border:1px solid #B2DDD6;border-radius:14px;padding:12px 30px 12px 12px;pointer-events:none;
  box-shadow:0 10px 28px rgba(15,118,110,.18),0 2px 6px rgba(16,24,40,.08);
  opacity:0;visibility:hidden;transition:opacity .2s ease,visibility .2s}}
.ai-bolla:before{{content:"";position:absolute;top:-7px;{freccia};width:12px;height:12px;background:#fff;
  border-left:1px solid #B2DDD6;border-top:1px solid #B2DDD6;transform:rotate(45deg)}}
.ai-x{{position:absolute;top:5px;right:7px;font-size:1.25rem;line-height:1;color:#98A2B3;cursor:pointer;pointer-events:auto;
  padding:1px 7px;border-radius:6px}}
.ai-x:hover{{background:#F2F4F7;color:#344054}}
.ai-avatar{{flex:0 0 44px;width:44px;height:44px;filter:drop-shadow(0 2px 4px rgba(15,118,110,.35))}}
.ai-occhi{{transform-origin:24px 21px;animation:ai-ammicca 4.5s infinite}}
@keyframes ai-ammicca{{0%,92%,100%{{transform:scaleY(1)}}96%{{transform:scaleY(.1)}}}}
.ai-nome{{font-size:.72rem;font-weight:700;letter-spacing:.06em;color:#0F766E;text-transform:uppercase}}
.ai-titolo{{font-size:.98rem;font-weight:700;color:#14213D;margin:1px 0 3px 0}}
.ai-testo{{font-size:.88rem;line-height:1.4;color:#344054}}
.ai-esempi{{margin-top:7px}}
.ai-chip{{display:inline-block;background:#F0FDFA;color:#0F766E;border:1px solid #99F6E4;border-radius:999px;padding:1px 10px;
  font-size:.78rem;margin:0 6px 4px 0}}
</style>"""
    html = (f'<div class="ai-wrap"><div class="ai-bolla"><span class="ai-x" title="Chiudi">&times;</span>{_AVATAR}<div><div class="ai-nome">Assistente Standard Cost</div>'
            f'<div class="ai-titolo">{titolo}</div><div class="ai-testo">{testo}</div>{righe}</div></div></div>')
    st.markdown(css + html, unsafe_allow_html=True)
