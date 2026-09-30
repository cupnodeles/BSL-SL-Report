# src/chaos.py
# Chaos mode — Cloud-safe port of bpi-automation-sl ChaosButton/useChaos.
# Streamlit Community Cloud allows <style>/<div> via unsafe_allow_html
# but NOT <script>, autoplay audio, or Vite-bundled GIFs. So this is:
# session_state toggle + CSS keyframes + emoji floaters. Silent by
# default; visuals only; pipeline state untouched.

import streamlit as st

MEMES = ["🐶", "🚗", "✨", "👾", "🌈", "⚡", "💥", "🛸"]


def is_chaos() -> bool:
    return bool(st.session_state.get("chaos", False))


def toggle_chaos() -> None:
    st.session_state["chaos"] = not is_chaos()
    if st.session_state["chaos"]:
        try:
            st.balloons()
        except Exception:
            pass


def inject_chaos_css() -> None:
    """Injects strobe/rumble/scatter styles + overlay divs. No-op when off."""
    if not is_chaos():
        return
    spots = []
    for i, emoji in enumerate(MEMES * 3):
        top = (i * 37) % 85
        spots.append(
            f'<span class="chaos-meme" style="top:{top}vh;'
            f'animation-duration:{3 + (i % 4)}s;'
            f'animation-delay:-{(i * 0.4):.1f}s;">{emoji}</span>'
        )
    st.markdown(
        """
<style>
.chaos-strobe {
  position: fixed; inset: 0; z-index: 45; pointer-events: none;
  opacity: 0.5; background: #ff0000;
  animation: chaos-strobe 0.5s steps(1) infinite;
}
@keyframes chaos-strobe {
  0% { background: #ff0000; } 16% { background: #ffea00; }
  33% { background: #00ff00; } 50% { background: #00e5ff; }
  66% { background: #0033ff; } 83% { background: #ff00cc; }
  100% { background: #ff0000; }
}
[data-testid="stAppViewContainer"] {
  animation: chaos-rumble 0.3s steps(2) infinite;
}
@keyframes chaos-rumble {
  0% { transform: translate(0,0); } 25% { transform: translate(-6px,4px); }
  50% { transform: translate(6px,-6px); } 75% { transform: translate(-4px,-4px); }
  100% { transform: translate(4px,6px); }
}
.px-panel, .px-title {
  animation: chaos-hue 0.7s linear infinite;
}
@keyframes chaos-hue {
  from { filter: hue-rotate(0deg) saturate(3); }
  to { filter: hue-rotate(360deg) saturate(3); }
}
.px-panel {
  animation: chaos-hue 0.7s linear infinite, chaos-wobble 0.9s ease-in-out infinite;
}
@keyframes chaos-wobble {
  0%, 100% { transform: rotate(-1deg); }
  50% { transform: rotate(1deg) translateY(-2px); }
}
.chaos-meme-layer {
  position: fixed; inset: 0; z-index: 44; pointer-events: none; overflow: hidden;
}
.chaos-meme {
  position: absolute; left: 0; font-size: 2.8rem; pointer-events: none;
  animation-name: chaos-meme-fly; animation-timing-function: linear;
  animation-iteration-count: infinite;
}
@keyframes chaos-meme-fly {
  from { transform: translate(-15vw, 20vh) rotate(-20deg); }
  to { transform: translate(110vw, -4vh) rotate(340deg); }
}
.chaos-btn-on button, .chaos-btn-on > div {
  background: linear-gradient(90deg,#ff0000,#ffea00,#00ff00,#00e5ff,#0033ff,#ff00cc,#ff0000) !important;
  background-size: 600% 100% !important;
  animation: chaos-btn-slide 0.5s linear infinite;
  color: #0b101d !important;
}
@keyframes chaos-btn-slide {
  from { background-position: 0% 0; } to { background-position: 100% 0; }
}
@media (prefers-reduced-motion: reduce) {
  .chaos-strobe, [data-testid="stAppViewContainer"],
  .px-panel, .chaos-meme { animation: none; }
}
</style>
<div class="chaos-strobe" aria-hidden="true"></div>
<div class="chaos-meme-layer" aria-hidden="true">"""
        + "".join(spots)
        + "</div>",
        unsafe_allow_html=True,
    )
