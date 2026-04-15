import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import matplotlib.patheffects as pe
from PIL import Image, ImageDraw, ImageOps
import io
import base64
import random
import time
import math

# ─────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="🐍 Snakes & Ladders",
    page_icon="🐍",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────
#  GLOBAL CSS  (retro arcade vibe)
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Press+Start+2P&family=Nunito:wght@400;700;900&display=swap');

html, body, [class*="css"] {
    font-family: 'Nunito', sans-serif;
    background: #0d0d1a;
    color: #f0e6ff;
}

.stApp {
    background: radial-gradient(ellipse at top, #1a0a2e 0%, #0d0d1a 60%);
}

h1, h2, h3 {
    font-family: 'Press Start 2P', monospace !important;
    color: #ffe94d !important;
    text-shadow: 0 0 20px #ffaa00, 0 0 40px #ff6600;
}

.pixel-title {
    font-family: 'Press Start 2P', monospace;
    font-size: 1.6rem;
    color: #ffe94d;
    text-shadow: 3px 3px 0px #cc5500, 0 0 30px #ffaa00;
    text-align: center;
    margin-bottom: 0.3rem;
    line-height: 2.2rem;
}

.subtitle {
    font-family: 'Nunito', sans-serif;
    font-size: 1.1rem;
    color: #b8a0ff;
    text-align: center;
    margin-bottom: 1.5rem;
}

.player-card {
    background: linear-gradient(135deg, #1e1040 0%, #2d1b69 100%);
    border: 2px solid #5533aa;
    border-radius: 16px;
    padding: 1.2rem;
    text-align: center;
    transition: all 0.3s ease;
    box-shadow: 0 4px 20px rgba(85,51,170,0.3);
}

.player-card:hover {
    border-color: #ffe94d;
    box-shadow: 0 0 25px rgba(255,233,77,0.4);
    transform: translateY(-3px);
}

.player-card.active-player {
    border: 3px solid #ffe94d !important;
    box-shadow: 0 0 35px rgba(255,233,77,0.7) !important;
    background: linear-gradient(135deg, #2d2000 0%, #4a3500 100%) !important;
    animation: pulse-border 1.5s infinite;
}

@keyframes pulse-border {
    0%, 100% { box-shadow: 0 0 20px rgba(255,233,77,0.6); }
    50% { box-shadow: 0 0 45px rgba(255,233,77,1); }
}

.player-name {
    font-family: 'Press Start 2P', monospace;
    font-size: 0.55rem;
    color: #ffe94d;
    margin-top: 0.5rem;
    word-break: break-word;
}

.player-pos {
    font-size: 0.8rem;
    color: #b8a0ff;
    margin-top: 0.3rem;
}

.dice-display {
    font-size: 4rem;
    text-align: center;
    animation: dice-spin 0.5s ease-out;
    filter: drop-shadow(0 0 12px #ffe94d);
}

@keyframes dice-spin {
    0% { transform: rotate(0deg) scale(0.5); opacity: 0; }
    60% { transform: rotate(20deg) scale(1.2); opacity: 1; }
    100% { transform: rotate(0deg) scale(1); opacity: 1; }
}

.event-box {
    background: linear-gradient(135deg, #1a0a2e, #2d1b69);
    border-left: 4px solid #ffe94d;
    border-radius: 8px;
    padding: 0.8rem 1rem;
    margin: 0.5rem 0;
    font-size: 0.85rem;
    color: #f0e6ff;
}

.event-snake {
    border-left-color: #ff4444 !important;
    background: linear-gradient(135deg, #2e0a0a, #4a1515) !important;
}

.event-ladder {
    border-left-color: #44ff88 !important;
    background: linear-gradient(135deg, #0a2e15, #154a25) !important;
}

.event-win {
    border-left-color: #ffe94d !important;
    background: linear-gradient(135deg, #2e2a00, #4a4400) !important;
    font-size: 1rem !important;
    font-weight: 900 !important;
}

.stButton > button {
    font-family: 'Press Start 2P', monospace !important;
    font-size: 0.6rem !important;
    background: linear-gradient(135deg, #5533aa, #7744cc) !important;
    color: #ffe94d !important;
    border: 2px solid #ffe94d !important;
    border-radius: 8px !important;
    padding: 0.7rem 1.2rem !important;
    transition: all 0.2s !important;
    box-shadow: 0 4px 15px rgba(85,51,170,0.5) !important;
    text-shadow: none !important;
}

.stButton > button:hover {
    background: linear-gradient(135deg, #7744cc, #9955ee) !important;
    box-shadow: 0 0 25px rgba(255,233,77,0.6) !important;
    transform: translateY(-2px) !important;
}

.roll-btn > button {
    font-size: 0.7rem !important;
    padding: 1rem 2rem !important;
    background: linear-gradient(135deg, #cc4400, #ff6600) !important;
    border-color: #ffaa00 !important;
    color: #fff !important;
    box-shadow: 0 6px 25px rgba(255,102,0,0.6) !important;
    width: 100% !important;
}

.upload-zone {
    background: linear-gradient(135deg, #1a1040, #2d1b69);
    border: 2px dashed #5533aa;
    border-radius: 16px;
    padding: 1.5rem;
    text-align: center;
}

.win-banner {
    font-family: 'Press Start 2P', monospace;
    font-size: 1.2rem;
    color: #ffe94d;
    text-align: center;
    text-shadow: 0 0 30px #ffaa00, 3px 3px 0 #cc5500;
    animation: win-flash 0.8s infinite;
    padding: 1.5rem;
    background: linear-gradient(135deg, #2e2a00, #4a4000);
    border: 3px solid #ffe94d;
    border-radius: 16px;
    margin: 1rem 0;
}

@keyframes win-flash {
    0%, 100% { text-shadow: 0 0 20px #ffaa00; }
    50% { text-shadow: 0 0 50px #ffe94d, 0 0 80px #ff6600; }
}

.log-entry { padding: 0.3rem 0; border-bottom: 1px solid #2d1b69; font-size: 0.82rem; }

/* hide streamlit default elements */
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 1.5rem !important; }

div[data-testid="stImage"] img {
    border-radius: 50% !important;
    border: 3px solid #5533aa !important;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  SOUND SYSTEM  (base64 HTML audio)
# ─────────────────────────────────────────────
def play_sound(sound_type: str):
    """Generate and play sounds using Web Audio API via HTML."""
    sounds = {
        "dice": """
            var ctx=new AudioContext();
            for(var i=0;i<6;i++){
                var o=ctx.createOscillator(),g=ctx.createGain();
                o.connect(g);g.connect(ctx.destination);
                o.frequency.value=300+Math.random()*400;
                o.type='square';
                g.gain.setValueAtTime(0.15,ctx.currentTime+i*0.07);
                g.gain.exponentialRampToValueAtTime(0.001,ctx.currentTime+i*0.07+0.06);
                o.start(ctx.currentTime+i*0.07);
                o.stop(ctx.currentTime+i*0.07+0.06);
            }
        """,
        "snake": """
            var ctx=new AudioContext();
            var o=ctx.createOscillator(),g=ctx.createGain();
            o.connect(g);g.connect(ctx.destination);
            o.type='sawtooth';
            o.frequency.setValueAtTime(600,ctx.currentTime);
            o.frequency.exponentialRampToValueAtTime(80,ctx.currentTime+0.8);
            g.gain.setValueAtTime(0.3,ctx.currentTime);
            g.gain.exponentialRampToValueAtTime(0.001,ctx.currentTime+0.8);
            o.start(ctx.currentTime);o.stop(ctx.currentTime+0.8);
        """,
        "ladder": """
            var ctx=new AudioContext();
            var notes=[261,329,392,523,659];
            notes.forEach(function(freq,i){
                var o=ctx.createOscillator(),g=ctx.createGain();
                o.connect(g);g.connect(ctx.destination);
                o.frequency.value=freq;o.type='sine';
                g.gain.setValueAtTime(0.2,ctx.currentTime+i*0.1);
                g.gain.exponentialRampToValueAtTime(0.001,ctx.currentTime+i*0.1+0.15);
                o.start(ctx.currentTime+i*0.1);o.stop(ctx.currentTime+i*0.1+0.15);
            });
        """,
        "win": """
            var ctx=new AudioContext();
            var melody=[523,659,784,1047,784,1047,1319];
            melody.forEach(function(freq,i){
                var o=ctx.createOscillator(),g=ctx.createGain();
                o.connect(g);g.connect(ctx.destination);
                o.frequency.value=freq;o.type='sine';
                g.gain.setValueAtTime(0.25,ctx.currentTime+i*0.12);
                g.gain.exponentialRampToValueAtTime(0.001,ctx.currentTime+i*0.12+0.2);
                o.start(ctx.currentTime+i*0.12);o.stop(ctx.currentTime+i*0.12+0.25);
            });
        """,
        "move": """
            var ctx=new AudioContext();
            var o=ctx.createOscillator(),g=ctx.createGain();
            o.connect(g);g.connect(ctx.destination);
            o.frequency.value=440;o.type='sine';
            g.gain.setValueAtTime(0.1,ctx.currentTime);
            g.gain.exponentialRampToValueAtTime(0.001,ctx.currentTime+0.1);
            o.start(ctx.currentTime);o.stop(ctx.currentTime+0.1);
        """,
    }
    js = sounds.get(sound_type, sounds["move"])
    st.markdown(f"<script>{js}</script>", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  GAME CONSTANTS
# ─────────────────────────────────────────────
SNAKES = {16:6, 47:26, 49:11, 56:53, 62:19, 64:60, 87:24, 93:73, 95:75, 99:78}
LADDERS = {4:14, 9:31, 20:38, 28:84, 40:59, 51:67, 63:81, 71:91}
PLAYER_COLORS = ["#FF6B6B","#4ECDC4","#FFE66D","#A8E6CF"]
DICE_FACES = ["⚀","⚁","⚂","⚃","⚄","⚅"]

# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────
def circular_crop(img: Image.Image, size=80) -> Image.Image:
    img = img.convert("RGBA").resize((size, size), Image.LANCZOS)
    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).ellipse((0,0,size,size), fill=255)
    result = Image.new("RGBA", (size, size), (0,0,0,0))
    result.paste(img, mask=mask)
    return result

def img_to_base64(img: Image.Image) -> str:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode()

def cell_to_xy(cell: int):
    """Convert cell number (1-100) to board (col, row) with boustrophedon."""
    cell -= 1
    row = cell // 10
    col = cell % 10
    if row % 2 == 1:
        col = 9 - col
    return col, row

# ─────────────────────────────────────────────
#  BOARD DRAWING
# ─────────────────────────────────────────────
def draw_board(positions, avatars, names, colors, current_player):
    fig, ax = plt.subplots(figsize=(9, 9))
    fig.patch.set_facecolor("#0d0d1a")
    ax.set_facecolor("#0d0d1a")
    ax.set_xlim(-0.5, 9.5)
    ax.set_ylim(-0.5, 9.5)
    ax.set_aspect('equal')
    ax.axis('off')

    # ── Cell colors
    safe_cells = {1, 8, 13, 24, 38, 50, 68, 72, 100}
    for num in range(1, 101):
        cx, cy = cell_to_xy(num)
        if num == 100:
            color = "#ffe94d"
        elif num in safe_cells:
            color = "#2d4a1e"
        elif (cx + cy) % 2 == 0:
            color = "#1e1040"
        else:
            color = "#2d1b69"

        rect = FancyBboxPatch((cx - 0.48, cy - 0.48), 0.96, 0.96,
                              boxstyle="round,pad=0.04", linewidth=0.8,
                              edgecolor="#3d2b79", facecolor=color)
        ax.add_patch(rect)

        txt_color = "#ffe94d" if num == 100 else "#7a6aaa" if num not in safe_cells else "#88ff88"
        ax.text(cx, cy + 0.28, str(num), ha='center', va='center',
                fontsize=5.5, color=txt_color, fontweight='bold',
                fontfamily='monospace')

    # ── Snakes
    for head, tail in SNAKES.items():
        hx, hy = cell_to_xy(head)
        tx, ty = cell_to_xy(tail)
        mid_x = (hx + tx) / 2 + random.choice([-1, 1]) * 0.8
        mid_y = (hy + ty) / 2
        xs = np.linspace(hx, tx, 60)
        ys = []
        for t in np.linspace(0, 1, 60):
            bx = (1-t)**2 * hx + 2*(1-t)*t * mid_x + t**2 * tx
            by_ = (1-t)**2 * hy + 2*(1-t)*t * mid_y + t**2 * ty
            wiggle = 0.18 * math.sin(t * math.pi * 5)
            ys.append(by_ + wiggle)

        bx_pts = [(1-t)**2*hx + 2*(1-t)*t*mid_x + t**2*tx for t in np.linspace(0,1,60)]
        ax.plot(bx_pts, ys, color="#ff4444", linewidth=3.5, alpha=0.85,
                solid_capstyle='round', zorder=3)
        ax.plot(bx_pts, ys, color="#ff8888", linewidth=1.2, alpha=0.5, zorder=4)
        # snake head circle
        ax.add_patch(plt.Circle((hx, hy), 0.22, color="#ff2222", zorder=5))
        ax.text(hx, hy, "🐍", ha='center', va='center', fontsize=7, zorder=6)

    # ── Ladders
    for bottom, top in LADDERS.items():
        bx, by = cell_to_xy(bottom)
        tx, ty = cell_to_xy(top)
        offset = 0.12
        ax.plot([bx - offset, tx - offset], [by, ty], color="#44ff88",
                linewidth=2.5, alpha=0.8, zorder=3)
        ax.plot([bx + offset, tx + offset], [by, ty], color="#44ff88",
                linewidth=2.5, alpha=0.8, zorder=3)
        rungs = 6
        for i in range(rungs):
            t = i / (rungs - 1)
            rx = bx - offset + t * (tx - bx)
            ry = by + t * (ty - by)
            ax.plot([rx - offset, rx + offset], [ry, ry], color="#88ffaa",
                    linewidth=1.8, alpha=0.7, zorder=4)
        ax.text(bx, by - 0.28, "🪜", ha='center', va='center', fontsize=7, zorder=6)

    # ── Player tokens
    token_offsets = [(-0.18, 0.12), (0.18, 0.12), (-0.18, -0.12), (0.18, -0.12)]
    for i, (pos, color, name) in enumerate(zip(positions, colors, names)):
        if pos < 1: continue
        cx, cy = cell_to_xy(pos)
        ox, oy = token_offsets[i]
        is_current = (i == current_player)
        ring_color = "#ffe94d" if is_current else color
        ring_w = 2.5 if is_current else 1.5
        circ = plt.Circle((cx + ox, cy + oy), 0.16, color=color,
                          zorder=8, linewidth=ring_w, edgecolor=ring_color)
        ax.add_patch(circ)
        initial = name[0].upper() if name else "?"
        ax.text(cx + ox, cy + oy, initial, ha='center', va='center',
                fontsize=6, color='#0d0d1a', fontweight='bold', zorder=9)
        if is_current:
            glow = plt.Circle((cx + ox, cy + oy), 0.22, color=color,
                             alpha=0.3, zorder=7)
            ax.add_patch(glow)

    # ── Border glow
    for spine in ['top','bottom','left','right']:
        ax.spines[spine].set_visible(False)
    border = FancyBboxPatch((-0.5, -0.5), 10, 10, boxstyle="round,pad=0.1",
                           linewidth=3, edgecolor="#5533aa", facecolor="none",
                           zorder=10)
    ax.add_patch(border)

    plt.tight_layout(pad=0.2)
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=130, bbox_inches='tight',
                facecolor="#0d0d1a", edgecolor='none')
    plt.close(fig)
    buf.seek(0)
    return buf

# ─────────────────────────────────────────────
#  SESSION STATE INIT
# ─────────────────────────────────────────────
def init_state():
    defaults = {
        "phase": "upload",       # upload → select → play → gameover
        "photos": [None]*4,
        "names": ["Player 1","Player 2","Player 3","Player 4"],
        "chosen_player": None,
        "positions": [0]*4,
        "current_turn": 0,
        "log": [],
        "last_dice": None,
        "last_event": None,
        "winner": None,
        "board_needs_update": True,
        "board_img": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()
S = st.session_state

# ─────────────────────────────────────────────
#  PHASE 1 — UPLOAD
# ─────────────────────────────────────────────
if S.phase == "upload":
    st.markdown('<div class="pixel-title">🐍 SNAKES & LADDERS 🪜</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitle">Upload 4 player photos to begin the chaos!</div>', unsafe_allow_html=True)

    cols = st.columns(4)
    all_uploaded = True
    for i, col in enumerate(cols):
        with col:
            st.markdown(f'<div class="upload-zone">', unsafe_allow_html=True)
            label = "👤 YOU" if i == 0 else f"🎮 Friend {i}"
            st.markdown(f"**{label}**")
            name = st.text_input(f"Name", value=S.names[i], key=f"name_{i}",
                                 label_visibility="collapsed",
                                 placeholder=f"Player {i+1} name")
            S.names[i] = name
            uploaded = st.file_uploader(f"Photo {i+1}", type=["jpg","jpeg","png"],
                                        key=f"upload_{i}", label_visibility="collapsed")
            if uploaded:
                img = Image.open(uploaded)
                S.photos[i] = img
                cropped = circular_crop(img, 100)
                st.image(cropped, use_container_width=False, width=100)
            elif S.photos[i]:
                cropped = circular_crop(S.photos[i], 100)
                st.image(cropped, use_container_width=False, width=100)
            else:
                st.markdown("📷 *Upload photo*")
                all_uploaded = False
            st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    _, mid, _ = st.columns([1,2,1])
    with mid:
        if all_uploaded:
            if st.button("⚡ NEXT — PICK YOUR PLAYER ⚡"):
                S.phase = "select"
                st.rerun()
        else:
            st.info("📸 Upload all 4 photos to continue")

# ─────────────────────────────────────────────
#  PHASE 2 — SELECT PLAYER
# ─────────────────────────────────────────────
elif S.phase == "select":
    st.markdown('<div class="pixel-title">👾 WHO ARE YOU?</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitle">Pick your character — the others become SNAKES on the board!</div>', unsafe_allow_html=True)

    cols = st.columns(4)
    for i, col in enumerate(cols):
        with col:
            cropped = circular_crop(S.photos[i], 120)
            st.image(cropped, use_container_width=False, width=120)
            st.markdown(f'<div class="player-name">{S.names[i]}</div>', unsafe_allow_html=True)
            if st.button(f"PICK ME!", key=f"pick_{i}"):
                S.chosen_player = i
                S.positions = [0, 0, 0, 0]
                S.current_turn = i   # chosen player goes first
                S.log = [f"🎮 {S.names[i]} chosen as player! Others are snakes... 🐍"]
                S.phase = "play"
                S.board_needs_update = True
                st.rerun()

# ─────────────────────────────────────────────
#  PHASE 3 — PLAY
# ─────────────────────────────────────────────
elif S.phase == "play":
    # title row
    st.markdown('<div class="pixel-title" style="font-size:1rem;">🐍 SNAKES & LADDERS 🪜</div>',
                unsafe_allow_html=True)

    board_col, ui_col = st.columns([3, 2], gap="medium")

    # ── Render board (cached until move happens)
    with board_col:
        if S.board_needs_update or S.board_img is None:
            buf = draw_board(S.positions, S.photos, S.names,
                            PLAYER_COLORS, S.current_turn)
            S.board_img = buf.read()
            S.board_needs_update = False
        st.image(S.board_img, use_container_width=True)

    with ui_col:
        # ── Player cards
        cols2 = st.columns(2)
        for i in range(4):
            col = cols2[i % 2]
            with col:
                is_active = (i == S.current_turn)
                is_chosen = (i == S.chosen_player)
                card_class = "player-card active-player" if is_active else "player-card"
                cropped = circular_crop(S.photos[i], 70)
                img_b64 = img_to_base64(cropped)
                role_tag = "🎮" if is_chosen else "🐍"
                pos_txt = f"Cell {S.positions[i]}" if S.positions[i] > 0 else "Start"
                turn_indicator = "▶ YOUR TURN" if is_active and is_chosen else ("▶ AI" if is_active else "")

                st.markdown(f"""
                <div class="{card_class}">
                    <img src="data:image/png;base64,{img_b64}" width="55"
                         style="border-radius:50%;border:3px solid {PLAYER_COLORS[i]};"/>
                    <div class="player-name">{role_tag} {S.names[i][:10]}</div>
                    <div class="player-pos">{pos_txt}</div>
                    {"<div style='color:#ffe94d;font-size:0.6rem;margin-top:4px;'>"+turn_indicator+"</div>" if turn_indicator else ""}
                </div>
                """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Dice + event display
        if S.last_dice:
            st.markdown(f'<div class="dice-display">{DICE_FACES[S.last_dice-1]}</div>',
                       unsafe_allow_html=True)
            st.markdown(f'<div style="text-align:center;color:#b8a0ff;font-size:0.85rem;">Rolled: {S.last_dice}</div>',
                       unsafe_allow_html=True)

        if S.last_event:
            ev_class = "event-box"
            if "snake" in S.last_event.lower() or "bitten" in S.last_event.lower():
                ev_class += " event-snake"
            elif "ladder" in S.last_event.lower() or "climbed" in S.last_event.lower():
                ev_class += " event-ladder"
            elif "win" in S.last_event.lower() or "🎉" in S.last_event:
                ev_class += " event-win"
            st.markdown(f'<div class="{ev_class}">{S.last_event}</div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Roll button (only show for chosen player's turn)
        is_my_turn = (S.current_turn == S.chosen_player)

        if is_my_turn:
            st.markdown('<div class="roll-btn">', unsafe_allow_html=True)
            roll_pressed = st.button("🎲 ROLL THE DICE!")
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            roll_pressed = False
            ai_btn = st.button(f"▶ Roll for {S.names[S.current_turn]}")
            if ai_btn:
                roll_pressed = True

        # ── GAME LOGIC
        if roll_pressed:
            dice = random.randint(1, 6)
            S.last_dice = dice
            cp = S.current_turn
            old_pos = S.positions[cp]
            new_pos = old_pos + dice
            event_msg = ""

            play_sound("dice")

            if new_pos > 100:
                event_msg = f"🚫 {S.names[cp]} needs {100 - old_pos} to win! No move."
                new_pos = old_pos
            else:
                S.positions[cp] = new_pos

                if new_pos == 100:
                    event_msg = f"🎉 {S.names[cp]} WON THE GAME! 🏆"
                    S.last_event = event_msg
                    S.log.insert(0, event_msg)
                    S.winner = cp
                    S.board_needs_update = True
                    play_sound("win")
                    S.phase = "gameover"
                    st.rerun()

                elif new_pos in SNAKES:
                    snake_tail = SNAKES[new_pos]
                    event_msg = f"🐍 OH NO! {S.names[cp]} bitten by snake at {new_pos}! Slides to {snake_tail}!"
                    S.positions[cp] = snake_tail
                    play_sound("snake")

                elif new_pos in LADDERS:
                    ladder_top = LADDERS[new_pos]
                    event_msg = f"🪜 LUCKY! {S.names[cp]} climbed ladder from {new_pos} to {ladder_top}!"
                    S.positions[cp] = ladder_top
                    play_sound("ladder")

                else:
                    event_msg = f"🎲 {S.names[cp]} rolled {dice} → moved to cell {new_pos}"
                    play_sound("move")

            S.last_event = event_msg
            S.log.insert(0, f"Turn {len(S.log)+1}: {event_msg}")
            if len(S.log) > 20:
                S.log = S.log[:20]

            # next turn
            S.current_turn = (S.current_turn + 1) % 4
            S.board_needs_update = True
            st.rerun()

        # ── Game log
        st.markdown("---")
        st.markdown("**📜 Game Log**")
        log_html = "".join([f'<div class="log-entry">{e}</div>' for e in S.log[:8]])
        st.markdown(f'<div style="max-height:180px;overflow-y:auto;">{log_html}</div>',
                   unsafe_allow_html=True)

        # ── Reset button
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 RESTART GAME"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

# ─────────────────────────────────────────────
#  PHASE 4 — GAME OVER
# ─────────────────────────────────────────────
elif S.phase == "gameover":
    st.markdown('<div class="pixel-title">🏆 GAME OVER! 🏆</div>', unsafe_allow_html=True)

    w = S.winner
    if w is not None and S.photos[w]:
        _, mid, _ = st.columns([1.5, 1, 1.5])
        with mid:
            cropped = circular_crop(S.photos[w], 180)
            st.image(cropped, use_container_width=False, width=180)

    st.markdown(f"""
    <div class="win-banner">
        🎉 {S.names[w] if w is not None else "Someone"} WINS! 🎉<br>
        <span style="font-size:0.7rem;color:#b8a0ff;">Reached Cell 100!</span>
    </div>
    """, unsafe_allow_html=True)

    # final board
    if S.board_img:
        _, mid, _ = st.columns([0.5, 3, 0.5])
        with mid:
            st.image(S.board_img, use_container_width=True)

    st.markdown("---")
    st.markdown("**📜 Final Game Log**")
    for entry in S.log[:15]:
        st.markdown(f'<div class="log-entry">{entry}</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    _, mid, _ = st.columns([1, 2, 1])
    with mid:
        if st.button("🎮 PLAY AGAIN WITH SAME PLAYERS"):
            S.positions = [0]*4
            S.current_turn = S.chosen_player
            S.log = [f"🔄 New game! {S.names[S.chosen_player]} goes first!"]
            S.last_dice = None
            S.last_event = None
            S.winner = None
            S.board_needs_update = True
            S.board_img = None
            S.phase = "play"
            st.rerun()

        if st.button("🔄 FULL RESTART"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()
