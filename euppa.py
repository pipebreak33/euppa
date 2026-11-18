#!/usr/bin/env python3
"""euppa – allt i Entra ID med några knapptryck. Windows, macOS och Linux.

  euppa                    meny
  euppa tenants            tenants du loggat in i
  euppa tenant NAMN        växla tenant
  euppa ny                 lägg till en tenant för hand (oftast onödigt – inloggning sparar den)
  euppa logga-in [--scopes A,B] [--enhetskod]   logga in i Microsoft Graph
  euppa åtgärder [--json]  lista åtgärderna (--json: allt Ruppa behöver)
  euppa kör ÅTGÄRD [Namn=värde …] [--ja] [--csv|--excel|--json]
                           kör en åtgärd (frågar efter det som saknas). Ändrande åtgärder kräver --ja
                           utan terminal. --csv/--excel sparar resultatet i ~/Documents/Euppa, --json skriver ut det
  euppa användare [UPN …]  välj användare (flerval) och gör något med dem
  euppa uppslag TYP [sök] [--alla] [--json]     sök i tenanten: user, group, device, app, sp, role, ca, sku …
  euppa kom-ihåg på|av     spara inloggningar mellan körningar (nyckelringen)
  euppa logga-ut           logga ut från Microsoft Graph
  euppa hej

Allt går via Microsoft Graph (Invoke-MgGraphRequest) i PowerShell 7, med rätt scopes per åtgärd.
Bara modulen Microsoft.Graph.Authentication behövs. Läsande åtgärder körs direkt, ändrande visar
kommandot och kräver ett ja.

Flaggor: -n/--torr visar kommandon utan att köra något · --no-anim stänger av animationer.
Data: ~/.euppa (kunder.json, config.json – inga hemligheter). Enbart standardbiblioteket, Python 3.8+.
"""
from __future__ import annotations

import base64
import json
import os
import queue
import random
import re
import shutil
import signal
import subprocess
import sys
import tempfile
import threading
import time
from datetime import datetime
from pathlib import Path

VERSION = "1.0.0"
HOME = Path.home()
DIR = Path(os.environ.get("EUPPA_DIR") or HOME / ".euppa")
os.environ["EUPPA_DIR"] = str(DIR)
CUSTOMERS = DIR / "kunder.json"
CONFIG = DIR / "config.json"
REPORTS = HOME / "Documents" / "Euppa"
IS_WIN = os.name == "nt"
IS_MAC = sys.platform == "darwin"
if IS_WIN:
    _PF = os.environ.get("ProgramFiles", r"C:\Program Files")
    _LA = os.environ.get("LOCALAPPDATA", str(HOME / "AppData" / "Local"))
    EXTRA_PATH = [os.path.join(_PF, "PowerShell", "7"), os.path.join(_LA, "Microsoft", "WindowsApps")]
else:
    EXTRA_PATH = ["/opt/homebrew/bin", "/home/linuxbrew/.linuxbrew/bin", "/usr/local/bin",
                  "/usr/local/microsoft/powershell/7", "/snap/bin", "/usr/bin", "/bin"]
os.environ["PATH"] = os.pathsep.join(dict.fromkeys(os.environ.get("PATH", "").split(os.pathsep) + EXTRA_PATH))


def _win_console():
    """UTF-8 + ANSI-färger i Windows-konsolen (Windows Terminal, conhost, VS Code)."""
    for stream in (sys.stdout, sys.stderr, sys.stdin):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    try:
        import ctypes
        k = ctypes.windll.kernel32
        k.SetConsoleOutputCP(65001)
        k.SetConsoleCP(65001)
        h = k.GetStdHandle(-11)
        mode = ctypes.c_uint32()
        if k.GetConsoleMode(h, ctypes.byref(mode)):
            k.SetConsoleMode(h, mode.value | 0x0004)       # ENABLE_VIRTUAL_TERMINAL_PROCESSING
    except Exception:
        pass


if IS_WIN:
    _win_console()


def pwsh_exe():
    """PowerShell 7: pwsh, annars förhandsversionen (Homebrew-casken powershell@preview heter pwsh-preview)."""
    return shutil.which("pwsh") or shutil.which("pwsh-preview") or "pwsh"


def ps_args(script, *flags):
    """pwsh med -EncodedCommand: inga citat-/escape-problem på varken Windows eller Unix."""
    enc = base64.b64encode(script.encode("utf-16-le")).decode()
    return [pwsh_exe(), "-NoLogo", "-NoProfile", *flags, "-EncodedCommand", enc]


def detached():
    """Arbetsprocessen ska inte få Ctrl+C från terminalen – Euppa avbryter den själv."""
    return {"creationflags": 0x00000200} if IS_WIN else {"start_new_session": True}


def copy_text(s):
    for cmd in (["pbcopy"], ["clip"], ["wl-copy"], ["xclip", "-selection", "clipboard"]):
        exe = shutil.which(cmd[0])
        if exe:
            try:
                subprocess.run([exe] + cmd[1:], input=s, text=True, timeout=5)
                return True
            except Exception:
                return False
    return False


def open_file(path):
    try:
        if IS_WIN:
            os.startfile(str(path))                              # noqa – finns bara på Windows
        elif IS_MAC:
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen([shutil.which("xdg-open") or "xdg-open", str(path)])
    except Exception:
        out(dim(f"  Öppna själv: {path}"))


# ───────────────────────────── terminal ─────────────────────────────
TTY = sys.stdout.isatty()
ITTY = sys.stdin.isatty()
COLOR = TTY and "NO_COLOR" not in os.environ
ANIM = TTY and COLOR and not os.environ.get("EUPPA_NO_ANIM") and not os.environ.get("CI")
DRY = False
EXPORT_NOW = None  # euppa kör … --excel/--csv/--json: exportera resultatet direkt
DIRTY = False      # något har skrivits ut som användaren bör hinna läsa innan skärmen töms


def c(code, s):
    return f"\033[{code}m{s}\033[0m" if COLOR else s


def green(s): return c("38;5;42", s)
def yellow(s): return c("38;5;214", s)
def red(s): return c("38;5;203", s)
def dim(s): return c("2", s)
def bold(s): return c("1", s)
def accent(s): return c("38;5;205", s)
def cyan(s): return c("38;5;80", s)


ANSI = re.compile(r"\x1b\[[0-9;?]*[a-zA-Z]")


def vlen(s):
    return len(ANSI.sub("", s))


def pad(s, n):
    return s + " " * max(0, n - vlen(s))


def cut(s, n):
    s = ANSI.sub("", s).replace("\t", " ").rstrip()
    return s if len(s) <= n else s[: max(0, n - 1)] + "…"


def width():
    return max(20, shutil.get_terminal_size((100, 30)).columns)


def height():
    return max(10, shutil.get_terminal_size((100, 30)).lines)


def fit(s, w=None):
    """Klipp en (färgad) rad så att den aldrig radbryts i terminalen."""
    w = max(8, (w or width()) - 1)
    if vlen(s) <= w:
        return s
    res, n, i = [], 0, 0
    while i < len(s) and n < w - 1:
        m = ANSI.match(s, i)
        if m:
            res.append(m.group(0))
            i = m.end()
            continue
        res.append(s[i])
        n += 1
        i += 1
    return "".join(res) + ("\033[0m" if COLOR else "") + "…"


def out(s=""):
    print(fit(s) if s else "")


def wrap(msg, w, maxl=4):
    lines, cur = [], ""
    for word in ANSI.sub("", msg).split():
        if cur and len(cur) + 1 + len(word) > w:
            lines.append(cur)
            cur = word
        else:
            cur = (cur + " " + word).strip()
    if cur:
        lines.append(cur)
    if len(lines) > maxl:
        lines = lines[:maxl]
        lines[-1] = cut(lines[-1] + " …", w)
    return [cut(l, w) for l in lines] or [""]


def hard_wrap(s, w):
    """Bryt en lång rad i bitar om w tecken (för kommandon som måste synas i sin helhet)."""
    s = s.replace("\t", "    ")
    if len(s) <= w:
        return [s]
    parts = [s[:w]]
    s = s[w:]
    while s:
        parts.append(s[: w - 2])
        s = s[w - 2:]
    return parts


class Live:
    """Ritar om ett block rader på plats, även efter att fönstret ändrat storlek."""

    def __init__(self):
        self.prev = []

    def _up(self):
        w = width()
        return sum(max(1, -(-n // w)) for n in self.prev)

    def draw(self, lines):
        w = width()
        lines = [fit(l, w) for l in lines]
        sys.stdout.write("\r" + "\033[F" * self._up() + "\033[J" + "".join(l + "\n" for l in lines))
        sys.stdout.flush()
        self.prev = [vlen(l) for l in lines]

    def clear(self):
        sys.stdout.write("\r" + "\033[F" * self._up() + "\033[J")
        sys.stdout.flush()
        self.prev = []


# ───────────────────────────── Euppa, maskoten ─────────────────────────────
# Euppa är ett tjockt E. Ansiktet bor i den övre gluggen, mellan översta och mellersta armen.
EYES = {"happy": " ◕ ◕ ", "blink": " ─ ─ ", "left": "◕ ◕  ", "right": "  ◕ ◕", "wait": " ◕ ◕ ",
        "sad": " ◕ ◕ ", "proud": " ^ ^ ", "wow": " ◕ ◕ ", "worry": " ◔ ◔ ", "sleep": " ─ ─ "}
MOUTH = {"happy": "  ‿  ", "blink": "  ‿  ", "left": " ‿   ", "right": "   ‿ ", "wait": "  ─  ",
         "sad": "  ⁀  ", "proud": "  ◡  ", "wow": "  o  ", "worry": "  ~  ", "sleep": "  ᶻ  "}
ART = ["█████████", "██{e}  ", "██{m}  ", "███████  ", "██       ", "█████████"]
SHADE = ["38;5;218", "38;5;212", "38;5;205", "38;5;199", "38;5;163", "38;5;126"]
SAD_SHADE = ["38;5;181", "38;5;175", "38;5;139", "38;5;133", "38;5;96", "38;5;59"]
ART_W = 9
LOGO = "E›"
JOKES = [
    "Jag heter Euppa. E som i Entra. Eller Effektiv. Eller Ett E med ögon.",
    "Villkorlig åtkomst är som en dörrvakt. Report-only är dörrvakten som bara antecknar.",
    "Azure AD heter Entra ID nu. Jag hette aldrig något annat, så jag slapp byta.",
    "Det finns två sorters admins: de som har ett break-glass-konto, och de som kommer att skaffa ett.",
    "En hemlighet som går ut i morgon är en incident i övermorgon. Jag håller koll.",
    "Global Admin åt alla? Nej. PIM åt alla som behöver, när de behöver.",
    "Invoke-MgGraphRequest: för att en modul räcker när man kan REST.",
    "Jag kryssar gärna i 200 användare åt dig. Jag frågar bara en gång innan jag ändrar.",
]


def euppa_art(mood="happy"):
    e, m = EYES.get(mood, EYES["happy"]), MOUTH.get(mood, MOUTH["happy"])
    sh = SHADE if mood != "sad" else SAD_SHADE
    res = []
    for k, row in enumerate(ART):
        if "{e}" in row or "{m}" in row:
            inner = e if "{e}" in row else m
            res.append(c(sh[k], "██") + bold(inner) + "  ")
        else:
            res.append(c(sh[k], row))
    return res


def frame(msg="", mood="happy", lift=0, top="", sub=""):
    """Maskot + pratbubbla. Under 44 kolumner döljs maskoten."""
    w = width()
    if w < 44:
        lines = []
        if msg:
            for k, l in enumerate(wrap(msg, w - 6, 4)):
                lines.append((f"  {accent(bold(LOGO))} " if k == 0 else "     ") + l)
        if sub:
            lines.append("     " + dim(cut(sub, w - 7)))
        return lines
    rows = [""] * 7
    for k, r in enumerate(euppa_art(mood)):
        rows[k + 1 - lift] = r
    if top and not rows[0]:
        rows[0] = top
    bub = []
    if msg:
        lines = wrap(msg, max(14, min(60, w - 20)), 4)
        iw = max(len(l) for l in lines)
        bub = [dim("╭" + "─" * (iw + 2) + "╮")] + \
              [dim("│ ") + l + " " * (iw - len(l)) + dim(" │") for l in lines] + \
              [dim("╰" + "─" * (iw + 2) + "╯")]
    res = []
    for k in range(7):
        line = "  " + pad(rows[k], ART_W + 1)
        if 0 <= k - 1 < len(bub):
            line += " " + (accent("◂ ") if k == 2 else "  ") + bub[k - 1]
        elif sub and k - 1 == len(bub):
            line += "    " + dim(cut(sub, w - ART_W - 9))
        res.append(fit(line, w))
    while res and not ANSI.sub("", res[-1]).strip():
        res.pop()
    while res and not ANSI.sub("", res[0]).strip():
        res.pop(0)
    return res


def say(msg, mood="happy", sub=""):
    global DIRTY
    DIRTY = True
    if TTY:
        print("\n".join(frame(msg, mood, sub=sub)))
    else:
        print(f"  {LOGO} {msg}")


def idle_mood(t):
    """Blinkar då och då, tittar runt ibland."""
    if t % 4.1 < 0.16:
        return "blink"
    phase = int(t / 2.3) % 7
    return {2: "left", 5: "right"}.get(phase, "happy")


def celebrate(msg):
    global DIRTY
    DIRTY = True
    if not ANIM:
        return say(msg, "proud")
    cols = ["38;5;205", "38;5;214", "38;5;80", "38;5;218", "38;5;227"]
    live = Live()
    sys.stdout.write("\033[?25l")
    try:
        for f in range(14):
            top = "".join(c(random.choice(cols), random.choice("✦✧⋆·*")) if random.random() < .35 else " "
                          for _ in range(ART_W + 1))
            live.draw(frame(msg, "proud" if f % 5 else "blink", lift=f % 2, top=top))
            time.sleep(0.09)
        live.draw(frame(msg, "proud"))
    finally:
        sys.stdout.write("\033[?25h")
        sys.stdout.flush()


# ───────────────────────────── input ─────────────────────────────
class Cancel(Exception):
    pass


class Keys:
    """Rå tangentläsning. Windows: msvcrt. macOS/Linux: termios + select. read() ger None vid timeout,
    så menyerna kan animera och rita om när fönstret ändrat storlek (även där SIGWINCH saknas)."""

    def __enter__(self):
        if IS_WIN:
            import msvcrt
            self.m = msvcrt
        else:
            import termios
            import tty
            self.fd = sys.stdin.fileno()
            self.old = termios.tcgetattr(self.fd)
            tty.setcbreak(self.fd)
        sys.stdout.write("\033[?25l")
        sys.stdout.flush()
        return self

    def __exit__(self, *a):
        if not IS_WIN:
            import termios
            termios.tcsetattr(self.fd, termios.TCSADRAIN, self.old)
        sys.stdout.write("\033[?25h")
        sys.stdout.flush()

    def read(self, timeout=0.12):
        if IS_WIN:
            end = time.time() + timeout
            while time.time() < end:
                if self.m.kbhit():
                    ch = self.m.getwch()
                    if ch in ("\x00", "\xe0"):
                        return {"H": "\x1b[A", "P": "\x1b[B", "K": "\x1b[D", "M": "\x1b[C",
                                "I": "\x1b[5~", "Q": "\x1b[6~"}.get(self.m.getwch(), "")
                    return ch
                time.sleep(0.02)
            return None
        import select
        r, _, _ = select.select([self.fd], [], [], timeout)
        return os.read(self.fd, 8).decode(errors="ignore") if r else None


def new_screen():
    """Töm skärmen inför en meny. Har något just skrivits ut får man läsa klart först."""
    global DIRTY
    if DIRTY:
        pause()
    DIRTY = False
    if TTY:
        sys.stdout.write("\033[H\033[2J")
        sys.stdout.flush()


def fuzzy_score(query, text, loose=True):
    """Fuzzy-match: alla ord i frågan måste finnas i texten, som delsträng (bäst) eller som bokstäver i ordning.
    Högre = bättre. None = ingen träff."""
    text = ANSI.sub("", text).lower()
    total = 0
    for word in query.lower().split():
        i = text.find(word)
        if i >= 0:
            start = i == 0 or not text[i - 1].isalnum()
            total += 100 + (40 if start else 0) - min(i, 30)
            continue
        if len(word) < 3 or not loose:       # korta sökord / långa texter: bara som delsträng
            return None
        pos, gaps, last = 0, 0, -1
        for ch in word:
            j = text.find(ch, pos)
            if j < 0:
                return None
            if last >= 0:
                gaps += j - last - 1
            last, pos = j, j + 1
        if gaps > 2 * len(word):             # för utspritt för att vara det man menade
            return None
        total += max(1, 60 - gaps * 3)
    return total


def fuzzy_rank(query, items, key):
    if not query.strip():
        return list(items)
    scored = [(fuzzy_score(query, key(x)), n, x) for n, x in enumerate(items)]
    return [x for s, n, x in sorted([y for y in scored if y[0] is not None], key=lambda y: (-y[0], y[1]))]


def choose(rows, title="", msg="", mood="happy", sub="", status=None, start=None, footer=None, screen=False,
           search_rows=None):
    """Piltangent-meny. rows: ("head", text) eller dict(label, hint, value, badge, key).
    Siffra väljer direkt, q/esc avbryter (→ None). Börja skriva för att fuzzy-söka (search_rows = sök i en
    större lista, t.ex. alla åtgärder). Maskoten blinkar medan den väntar.
    screen=True: menyn ritas på en tom skärm (navigeringsmenyer), annars där markören står (frågor i ett flöde)."""
    if not ITTY:
        raise Cancel()
    if screen:
        new_screen()
    hot = {r["key"] for r in rows if not isinstance(r, tuple) and r.get("key")}
    st = {"q": "", "pos": 0, "top": 0, "mood": None}

    def view():
        if st["q"]:
            pool = [r for r in (search_rows if search_rows is not None else rows) if not isinstance(r, tuple)]
            pool = [r for r in pool if r.get("value") is not None]
            q = st["q"]
            scored = []
            for n, r in enumerate(pool):
                a = fuzzy_score(q, r["label"])
                b = fuzzy_score(q, r["label"] + " " + ANSI.sub("", r.get("hint") or "") + " " + (r.get("search") or ""),
                                loose=False)
                if a is None and r.get("search"):
                    a = fuzzy_score(q, r["search"])
                sc = max(a * 2 if a is not None else -1, b if b is not None else -1)
                if sc >= 0:
                    scored.append((-sc, n, r))
            return [r for _, _, r in sorted(scored, key=lambda x: (x[0], x[1]))]
        return rows

    def numbering(vr):
        sel = [k for k, r in enumerate(vr) if not isinstance(r, tuple)]
        keys, n = {}, 0
        for k in sel:
            r = vr[k]
            if r.get("key") and not st["q"]:
                keys[r["key"]] = k
            else:
                n += 1
                if n <= 9:
                    keys[str(n)] = k
        return sel, keys

    vr = view()
    sel, keys = numbering(vr)
    if not sel:
        return None
    st["pos"] = sel[0]
    if start is not None:
        for k in sel:
            if vr[k]["value"] == start:
                st["pos"] = k
    live = Live()
    winch = {"flag": False}
    t0 = time.time()

    def draw(mood_now):
        w = width()
        lines = []
        if msg and not st["q"]:
            lines += frame(msg, mood_now, sub=sub)
        if status:
            lines += [""] + status()
        lines.append("")
        if st["q"]:
            lines.append("  " + accent("⌕ ") + bold(st["q"]) + accent("▏") + dim(f"  {len(sel)} träffar"))
        elif title:
            lines.append("  " + accent("▍") + bold(title))
        narrow = w < 58
        if not sel:
            lines.append("     " + dim("Inga träffar. Backsteg ändrar, esc rensar."))
        labels = [vr[k]["label"] for k in sel] or [""]
        lw = min(max(len(x) for x in labels) + 2, max(10, w - 14 if narrow else min(38, w // 2)))
        room = max(3, height() - len(lines) - 5)     # ↑/↓ fler, tomrad och hjälptext
        pos = st["pos"]
        top = min(max(st["top"], pos - room + 1), pos)
        top = max(0, min(top, max(0, len(vr) - room)))
        st["top"] = top
        if top > 0:
            lines.append(dim("     ↑ fler"))
        num = {k: n for n, k in keys.items()}
        anybadge = any(vr[j].get("badge") for j in sel)
        for k in range(top, min(len(vr), top + room)):
            r = vr[k]
            if isinstance(r, tuple):
                lines.append("   " + dim(cut(r[1].upper(), w - 5)))
                continue
            on = k == pos
            key = num.get(k, " ")
            label = cut(r["label"], lw - 1)
            label = bold(accent(label)) if on else label
            badge = r.get("badge") or ""
            if narrow:
                line = f"  {accent('❯') if on else ' '}{dim(key)} " + (pad(label, lw) + badge if anybadge else label)
            else:
                line = f"   {accent('❯') if on else ' '} {dim(key)}  " + pad(label, lw) + (pad(badge, 8) if anybadge else "")
                if r.get("hint") and w >= 64:
                    line += dim(ANSI.sub("", r["hint"]))
            lines.append(line)
        if top + room < len(vr):
            lines.append(dim("     ↓ fler"))
        lines.append("")
        if st["q"]:
            help_ = "↑/↓ · enter väljer · backsteg · esc rensar" if w >= 50 else "↑↓ enter · esc"
        else:
            help_ = ((footer or "↑/↓ · enter · siffra väljer · q tillbaka") + " · skriv för att söka") if w >= 72 else \
                (footer or "↑/↓ · enter · q tillbaka") if w >= 50 else "↑↓ enter · q"
        lines.append(dim("   " + help_))
        live.draw(lines)

    def refilter():
        nonlocal vr, sel, keys
        vr = view()
        sel, keys = numbering(vr)
        st["pos"] = sel[0] if sel else 0
        st["top"] = 0

    def on_winch(*_):
        winch["flag"] = True

    old_winch = signal.signal(signal.SIGWINCH, on_winch) if hasattr(signal, "SIGWINCH") else None
    size = shutil.get_terminal_size((100, 30))
    choice = None
    picked = None
    kb = Keys()
    try:
        kb.__enter__()
        draw(mood)
        while True:
            ch = kb.read(0.12)
            if ch is None:
                now = shutil.get_terminal_size((100, 30))
                if now != size:              # fönstret ändrat storlek (Windows har ingen SIGWINCH)
                    size = now
                    winch["flag"] = True
                m = mood
                if ANIM and not st["q"]:
                    im = idle_mood(time.time() - t0)
                    m = im if mood == "happy" else ("blink" if im == "blink" else mood)
                if winch["flag"] or m != st["mood"]:
                    winch["flag"] = False
                    st["mood"] = m
                    draw(m)
                continue
            idx = sel.index(st["pos"]) if st["pos"] in sel else 0
            if ch in ("\x1b[A", "\x1bOA"):
                if sel:
                    st["pos"] = sel[(idx - 1) % len(sel)]
            elif ch in ("\x1b[B", "\x1bOB"):
                if sel:
                    st["pos"] = sel[(idx + 1) % len(sel)]
            elif ch in ("\r", "\n"):
                if sel:
                    picked = vr[st["pos"]]
                    choice = picked["value"]
                    break
            elif ch in ("\x7f", "\x08"):
                if st["q"]:
                    st["q"] = st["q"][:-1]
                    refilter()
            elif ch == "\x1b":
                if st["q"]:
                    st["q"] = ""
                    refilter()
                else:
                    break
            elif ch in ("\x03", "\x04"):
                break
            elif not st["q"] and ch.isdigit() and ch in keys:
                picked = vr[keys[ch]]
                choice = picked["value"]
                break
            elif not st["q"] and ch in hot:
                picked = vr[keys[ch]]
                choice = picked["value"]
                break
            elif not st["q"] and ch == "q":
                break
            elif ch and not ch.startswith("\x1b") and ch.isprintable() and (st["q"] or not ch.startswith(" ")):
                st["q"] += ch                    # även inklistrat / snabbt skrivet (flera tecken på en gång)
                refilter()
            draw(st["mood"] or mood)
    except KeyboardInterrupt:
        choice = None
    finally:
        if old_winch is not None:
            signal.signal(signal.SIGWINCH, old_winch)
        kb.__exit__()
        live.clear()
        sys.stdout.flush()
    if choice is not None and not screen:
        out(f"   {accent('❯')} {bold(picked['label'])}")
    return choice


def ask_text(question, default="", required=False, check=None, secret=False):
    """Fråga efter en text. Enter = standardvärdet. Ctrl+C avbryter flödet."""
    if not ITTY:
        if required and not default:
            raise Cancel()
        return default
    w = width()
    for k, l in enumerate(wrap(question, w - 6, 6)):
        out(("  " + accent("?") + " " if k == 0 else "    ") + bold(l))
    if default:
        out("    " + dim("standard: " + cut(default, w - 16) + "  (enter behåller)"))
    while True:
        try:
            if secret:
                import getpass
                v = getpass.getpass("    › ")
            else:
                v = input("    " + accent("›") + " ")
        except (EOFError, KeyboardInterrupt):
            print()
            raise Cancel()
        v = v.strip()
        if not v:
            v = default
        if not v and required:
            out("    " + yellow("Det här behövs. Ctrl+C avbryter."))
            continue
        if check and v:
            err = check(v)
            if err:
                out("    " + yellow(err))
                continue
        return v


def confirm(question, default=False):
    if not ITTY:
        return default
    opts = "[J/n]" if default else "[j/N]"
    try:
        a = input(fit(f"  {accent('?')} {bold(question)} {dim(opts)}") + " ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        print()
        return False
    if not a:
        return default
    return a in ("j", "ja", "y", "yes")


def pause():
    """Vänta på enter (eller esc/q) innan skärmen töms. Läser tangenter rått så att piltangenter inte syns som ^[[A."""
    if ITTY:
        print()
        sys.stdout.write(dim("   enter = tillbaka till menyn "))
        sys.stdout.flush()
        kb = Keys()
        try:
            kb.__enter__()
            while True:
                ch = kb.read(0.5)
                if ch is None:
                    continue
                if ch in ("\r", "\n", "\x1b", "q", "\x03", "\x04", " "):
                    break
        except KeyboardInterrupt:
            pass
        finally:
            kb.__exit__()
            print()


def pick_many(items, title, chosen=None):
    """Lista med kryssrutor och fuzzy-sök. items: [(värde, namn, detalj)]. -> [värden] (tom = avbrutet).
    Skriv för att söka · mellanslag/tab kryssar · ctrl+a kryssar alla som syns · enter klar · esc rensar/avbryter."""
    if not ITTY:
        raise Cancel()
    chosen = list(chosen or [])
    st = {"q": "", "pos": 0, "top": 0}

    def view():
        return fuzzy_rank(st["q"], items, lambda x: f"{x[1]} {x[0]} {x[2]}") if st["q"].strip() else list(items)

    vis = view()
    live = Live()

    def draw():
        w = width()
        lines = [""]
        if st["q"]:
            lines.append("  " + accent("⌕ ") + bold(st["q"]) + accent("▏") + dim(f"  {len(vis)} träffar"))
        else:
            lines.append("  " + accent("▍") + bold(title) + dim(f"  {len(items)} st · skriv för att söka"))
        lines.append("  " + (green(f"{len(chosen)} valda: ") + cut(", ".join(
            next((i[1] for i in items if i[0] == v), v) for v in chosen), w - 16) if chosen else dim("Inga valda än.")))
        lines.append("")
        room = max(3, height() - len(lines) - 4)
        pos = min(st["pos"], max(0, len(vis) - 1))
        top = min(max(st["top"], pos - room + 1), pos)
        st["top"], st["pos"] = top, pos
        if not vis:
            lines.append("     " + dim("Inga träffar. Backsteg ändrar, esc rensar."))
        if top > 0:
            lines.append(dim("     ↑ fler"))
        nw = min(max([len(v[1]) for v in vis[top:top + room]] or [10]) + 2, max(12, w // 2 - 6))
        for k in range(top, min(len(vis), top + room)):
            v = vis[k]
            on = k == pos
            mark = green("◉") if v[0] in chosen else dim("○")
            name = cut(v[1] or v[0], nw - 1)
            line = f"   {accent('❯') if on else ' '} {mark} " + pad(bold(accent(name)) if on else name, nw)
            if w >= 64:
                line += dim(v[2] or v[0])
            lines.append(line)
        if top + room < len(vis):
            lines.append(dim("     ↓ fler"))
        lines.append("")
        lines.append(dim("   " + ("mellanslag/tab kryssar · ctrl+a alla som syns · enter klar · esc avbryter"
                                  if w >= 80 else "tab kryssar · enter klar · esc")))
        live.draw(lines)

    kb = Keys()
    done = False
    try:
        kb.__enter__()
        draw()
        size = shutil.get_terminal_size((100, 30))
        while True:
            ch = kb.read(0.12)
            if ch is None:
                if shutil.get_terminal_size((100, 30)) != size:
                    size = shutil.get_terminal_size((100, 30))
                    draw()
                continue
            if ch in ("\x1b[A", "\x1bOA"):
                st["pos"] = max(0, st["pos"] - 1)
            elif ch in ("\x1b[B", "\x1bOB"):
                st["pos"] = min(max(0, len(vis) - 1), st["pos"] + 1)
            elif ch in ("\x1b[5~",):
                st["pos"] = max(0, st["pos"] - 10)
            elif ch in ("\x1b[6~",):
                st["pos"] = min(max(0, len(vis) - 1), st["pos"] + 10)
            elif ch == "\t" or (ch == " " and not st["q"]):
                if vis:
                    v = vis[st["pos"]][0]
                    chosen.remove(v) if v in chosen else chosen.append(v)
                    st["pos"] = min(len(vis) - 1, st["pos"] + 1)
            elif ch == "\x01":
                for v in vis:
                    if v[0] not in chosen:
                        chosen.append(v[0])
            elif ch in ("\r", "\n"):
                if not chosen and vis:
                    chosen.append(vis[st["pos"]][0])
                done = True
                break
            elif ch in ("\x7f", "\x08"):
                st["q"] = st["q"][:-1]
                vis, st["pos"], st["top"] = view(), 0, 0
            elif ch == "\x1b":
                if st["q"]:
                    st["q"] = ""
                    vis, st["pos"], st["top"] = view(), 0, 0
                else:
                    break
            elif ch in ("\x03", "\x04"):
                break
            elif ch.isprintable() and not ch.startswith("\x1b"):
                st["q"] += ch
                vis, st["pos"], st["top"] = view(), 0, 0
            draw()
    except KeyboardInterrupt:
        done = False
    finally:
        kb.__exit__()
        live.clear()
    return chosen if done else []


# ───────────────────────────── lagring ─────────────────────────────
GUID = re.compile(r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")


def private(path):
    try:
        os.chmod(path, 0o700 if Path(path).is_dir() else 0o600)
    except OSError:
        pass


def load_json(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def save_json(path, data):
    DIR.mkdir(parents=True, exist_ok=True)
    private(DIR)
    tmp = Path(str(path) + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=1, ensure_ascii=False)
    private(tmp)
    os.replace(tmp, path)


def config():
    return load_json(CONFIG, {})


def set_config(**kw):
    cfg = config()
    cfg.update(kw)
    save_json(CONFIG, cfg)


def tenants():
    """Tenants du loggat in i. Första gången lånas Puppas lista (samma kunder, samma konton)."""
    d = load_json(CUSTOMERS, None)
    if d is None:
        puppa = load_json(HOME / ".puppa" / "kunder.json", {}).get("tenants") or []
        ts = [dict(name=t["name"], tenant=t.get("tenant", ""), tenant_id=t.get("tenant_id", ""), upn=t.get("upn", ""),
                   auth="device" if t.get("auth") == "device" else "interactive",
                   created=t.get("created", ""), last=t.get("last", ""))
              for t in puppa if t.get("name")]
        if ts:
            save_json(CUSTOMERS, {"tenants": ts})
        return ts
    return d.get("tenants", [])


def save_tenants(ts):
    save_json(CUSTOMERS, {"tenants": ts})


def find_tenant(name):
    if not name:
        return None
    n = name.lower()
    ts = tenants()
    for t in ts:
        if n in (t["name"].lower(), t.get("tenant", "").lower(), t.get("tenant_id", "").lower()):
            return t
    hits = [t for t in ts if t["name"].lower().startswith(n) or t.get("tenant", "").lower().startswith(n)]
    return hits[0] if len(hits) == 1 else None


def active_tenant():
    return find_tenant(config().get("active"))


def update_tenant(name, **kw):
    ts = tenants()
    for t in ts:
        if t["name"] == name:
            t.update(kw)
    save_tenants(ts)


def slug(s):
    s = s.lower().replace("å", "a").replace("ä", "a").replace("ö", "o")
    return re.sub(r"[^a-z0-9]+", "-", s).strip("-") or "tenant"


def unique_name(base, ts):
    names = {t["name"].lower() for t in ts}
    name, k = base, 2
    while name.lower() in names:
        name, k = f"{base} {k}", k + 1
    return name


def tenant_domain(t):
    """'contoso' / 'contoso.onmicrosoft.com' / 'contoso.se' / en GUID -> det Connect-MgGraph -TenantId tar."""
    t = re.sub(r"^https?://", "", t.strip().lower()).split("/")[0]
    if GUID.match(t) or "." in t:
        return t
    return t + ".onmicrosoft.com"


def remember_login(who):
    """Efter en lyckad inloggning: spara/uppdatera tenanten (domän, id, konto) och gör den aktiv."""
    tid, upn, domain, org, default = (list(who) + ["", "", "", "", ""])[:5]
    if not domain and "@" in upn:
        domain = upn.split("@", 1)[1]
    ts = tenants()
    t = next((x for x in ts if tid and x.get("tenant_id", "").lower() == tid.lower()), None) or \
        next((x for x in ts if domain and x.get("tenant", "").lower() == domain.lower()), None)
    new = t is None
    now = datetime.now().isoformat(timespec="seconds")
    if new:
        base = org or (domain.split(".")[0].capitalize() if domain else "Tenant")
        t = dict(name=unique_name(base, ts), tenant=domain or tid, tenant_id=tid, upn=upn,
                 auth=config().get("auth", "interactive"), created=now)
        ts.append(t)
    else:
        if tid:
            t["tenant_id"] = tid
        if upn:
            t["upn"] = upn
        cur = t.get("tenant", "")
        if domain and (not cur or GUID.match(cur) or
                       (domain.endswith(".onmicrosoft.com") and not cur.endswith(".onmicrosoft.com"))):
            t["tenant"] = domain
    if default:
        t["domain"] = default
    t["last"] = now
    save_tenants(ts)
    set_config(active=t["name"])
    return t, new


# ───────────────────────────── PowerShell ─────────────────────────────
def psq(v):
    """PowerShell-sträng med enkla citattecken. PowerShell räknar även ‘ ’ ‚ ‛ som apostrofer – dubbla dem också."""
    return "'" + re.sub(r"(['‘’‚‛])", r"\1\1", str(v)) + "'"


R = "\x1e"
# Hjälpfunktioner som finns i varje körning. Allt mot Graph går via Graph (Invoke-MgGraphRequest),
# så att bara modulen Microsoft.Graph.Authentication behövs.
EU_LIB = r"""
$global:EuH = @{ 'ocp-client-name' = 'Euppa'; 'ocp-client-version' = '1.0' }
function global:Graph {
  param([Parameter(Position = 0)][string]$Uri, [string]$Method = 'GET', $Body, [switch]$All, [switch]$Eventual,
        [switch]$Beta, [hashtable]$Headers)
  if ($Uri -notmatch '^https://') { $Uri = 'https://graph.microsoft.com/' + $(if ($Beta) { 'beta' } else { 'v1.0' }) + '/' + $Uri.TrimStart('/') }
  $h = @{}
  if ($Headers) { foreach ($k in $Headers.Keys) { $h[$k] = $Headers[$k] } }
  if ($Eventual) { $h['ConsistencyLevel'] = 'eventual' }
  $p = @{ Method = $Method; Uri = $Uri; OutputType = 'PSObject'; ErrorAction = 'Stop' }
  if ($h.Count) { $p.Headers = $h }
  if ($null -ne $Body) { $p.Body = (ConvertTo-Json -InputObject $Body -Depth 20 -Compress); $p.ContentType = 'application/json' }
  $r = Invoke-MgGraphRequest @p
  if ($Method -ne 'GET') { return $r }
  if ($null -ne $r -and $r.PSObject.Properties.Name -contains 'value') {
    $r.value
    while ($All -and $r.'@odata.nextLink') { $p.Uri = $r.'@odata.nextLink'; $r = Invoke-MgGraphRequest @p; $r.value }
  } else { $r }
}
function global:Enc([string]$s) { [uri]::EscapeDataString($s) }
function global:Q([string]$s) { $s -replace "'", "''" }
function global:Seg([string]$s) { $s -replace '%', '%25' -replace '#', '%23' -replace '\?', '%3F' -replace '/', '%2F' }
function global:Items([string]$s) { @($s -split '[,;\r\n]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
function global:IsGuid([string]$s) { $s -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$' }
function global:Msg($e) {
  $t = "$($e.ErrorDetails.Message) $($e.Exception.Message) $e"
  $m = [regex]::Match($t, '"message"\s*:\s*"((?:[^"\\]|\\.)*)"')
  if ($m.Success) { return ($m.Groups[1].Value -replace '\\"', '"') }
  ("$($e.Exception.Message)" -split "`n")[0]
}
function global:D($x) { if (-not $x) { return '' }; try { ([datetime]$x).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } catch { "$x" } }
function global:Ago($x) { if (-not $x) { return $null }; try { [int]((Get-Date) - ([datetime]$x)).TotalDays } catch { $null } }
function global:Since([double]$hours) { (Get-Date).ToUniversalTime().AddHours(-$hours).ToString('yyyy-MM-ddTHH:mm:ssZ') }
function global:JaNej($b) { if ($b) { 'ja' } else { 'nej' } }
function global:Pick($v, $n, $d) { Write-Host ('EUPPA-PICK|{0}|{1}|{2}' -f $v, ("$n" -replace '\|', '/'), ("$d" -replace '\|', '/')) }
function global:Done($obj, [string]$what) { [pscustomobject][ordered]@{ Objekt = "$obj"; Resultat = $what } }
function global:Fail($obj, $err) {
  $m = if ($err -is [string]) { $err } else { Msg $err }
  Write-Error "${obj}: $m"
  [pscustomobject][ordered]@{ Objekt = "$obj"; Resultat = "Fel: $m" }
}
function global:Me { if (-not $global:EuMe) { $global:EuMe = Graph 'me?$select=id,userPrincipalName,displayName' }; $global:EuMe }
function global:NewPw {
  $r = [System.Security.Cryptography.RandomNumberGenerator]
  $up = 'ABCDEFGHJKLMNPQRSTUVWXYZ'; $lo = 'abcdefghijkmnopqrstuvwxyz'; $di = '23456789'; $sy = '!#%+-?'
  $s = (1..3 | ForEach-Object { $up[$r::GetInt32($up.Length)] }) + (1..7 | ForEach-Object { $lo[$r::GetInt32($lo.Length)] }) +
       $sy[$r::GetInt32($sy.Length)] + (1..3 | ForEach-Object { $di[$r::GetInt32($di.Length)] })
  -join $s
}
function global:UserType($o) {
  switch -Wildcard ($o.'@odata.type') { '*user' { 'användare' } '*group' { 'grupp' } '*servicePrincipal' { 'app' }
    '*device' { 'enhet' } '*directoryRole' { 'roll' } '*administrativeUnit' { 'adminenhet' } default { "$($o.'@odata.type')" -replace '#microsoft.graph.', '' } }
}
function global:GroupType($g) {
  if ($g.groupTypes -contains 'Unified') { 'Microsoft 365' } elseif ($g.securityEnabled -and $g.mailEnabled) { 'e-postaktiverad säkerhet' }
  elseif ($g.securityEnabled) { 'säkerhet' } else { 'distribution' }
}
$global:EuUserSel = 'id,displayName,userPrincipalName,mail,accountEnabled,usageLocation,userType'
function global:Users([string]$s) {
  foreach ($x in (Items $s)) {
    try { Graph ('users/' + (Seg $x) + '?$select=' + $global:EuUserSel) }
    catch {
      $err = $_
      $hit = @()
      if ($x -notmatch '@') { try { $hit = @(Graph ('users?$select=' + $global:EuUserSel + '&$filter=' + (Enc "displayName eq '$(Q $x)'"))) } catch { } }
      if ($hit.Count -eq 1) { $hit[0] } else { Write-Error ("Hittar ingen användare '{0}': {1}" -f $x, (Msg $err)) }
    }
  }
}
function global:Groups([string]$s) {
  foreach ($x in (Items $s)) {
    $sel = '$select=id,displayName,mail,groupTypes,securityEnabled,mailEnabled,membershipRule,membershipRuleProcessingState'
    $hit = @()
    try {
      if (IsGuid $x) { $hit = @(Graph ("groups/$x" + '?' + $sel)) }
      else { $hit = @(Graph ('groups?' + $sel + '&$filter=' + (Enc "mail eq '$(Q $x)' or displayName eq '$(Q $x)'"))) }
    } catch { }
    if ($hit.Count -eq 1) { $hit[0] }
    elseif ($hit.Count -gt 1) { Write-Error "Flera grupper heter '$x'. Välj med id i stället (euppa uppslag group $x)." }
    else { Write-Error "Hittar ingen grupp '$x'." }
  }
}
function global:Devices([string]$s) {
  foreach ($x in (Items $s)) {
    $sel = '$select=id,deviceId,displayName,accountEnabled,operatingSystem,operatingSystemVersion,trustType'
    $hit = @()
    try {
      if (IsGuid $x) {
        $hit = @(Graph ('devices?' + $sel + '&$filter=' + (Enc "deviceId eq '$x'")))
        if (-not $hit.Count) { $hit = @(Graph ("devices/$x" + '?' + $sel)) }
      } else { $hit = @(Graph ('devices?' + $sel + '&$filter=' + (Enc "displayName eq '$(Q $x)'"))) }
    } catch { }
    if ($hit.Count -gt 1) { Write-Warning "$($hit.Count) enheter heter '$x' – tar alla." }
    if ($hit.Count) { $hit } else { Write-Error "Hittar ingen enhet '$x'." }
  }
}
function global:Apps([string]$s) {
  foreach ($x in (Items $s)) {
    $hit = @()
    try {
      if (IsGuid $x) { $hit = @(Graph ('applications?$filter=' + (Enc "appId eq '$x'"))); if (-not $hit.Count) { $hit = @(Graph "applications/$x") } }
      else { $hit = @(Graph ('applications?$filter=' + (Enc "displayName eq '$(Q $x)'"))) }
    } catch { }
    if ($hit.Count -eq 1) { $hit[0] } elseif ($hit.Count -gt 1) { Write-Error "Flera appar heter '$x'. Använd AppId." }
    else { Write-Error "Hittar ingen appregistrering '$x'." }
  }
}
function global:Sps([string]$s) {
  foreach ($x in (Items $s)) {
    $hit = @()
    try {
      if (IsGuid $x) { $hit = @(Graph ('servicePrincipals?$filter=' + (Enc "appId eq '$x'"))); if (-not $hit.Count) { $hit = @(Graph "servicePrincipals/$x") } }
      else { $hit = @(Graph ('servicePrincipals?$filter=' + (Enc "displayName eq '$(Q $x)'"))) }
    } catch { }
    if ($hit.Count -eq 1) { $hit[0] } elseif ($hit.Count -gt 1) { Write-Error "Flera företagsappar heter '$x'. Använd AppId." }
    else { Write-Error "Hittar ingen företagsapp '$x'." }
  }
}
function global:SpByAppId([string]$appId) {
  if (-not $global:EuSp) { $global:EuSp = @{} }
  if (-not $global:EuSp.ContainsKey($appId)) { $global:EuSp[$appId] = @(Graph ('servicePrincipals?$filter=' + (Enc "appId eq '$appId'")))[0] }
  $global:EuSp[$appId]
}
function global:Roles([string]$s) {
  foreach ($x in (Items $s)) {
    $hit = @()
    try {
      if (IsGuid $x) { $hit = @(Graph "roleManagement/directory/roleDefinitions/$x") }
      else { $hit = @(Graph ('roleManagement/directory/roleDefinitions?$filter=' + (Enc "displayName eq '$(Q $x)'"))) }
    } catch { }
    if ($hit.Count) { $hit[0] } else { Write-Error "Hittar ingen roll '$x'." }
  }
}
function global:Sku([string]$s) {
  $all = @(Graph 'subscribedSkus')
  $hit = @($all | Where-Object { $_.skuPartNumber -eq $s -or $_.skuId -eq $s })
  if (-not $hit.Count) { throw "Hittar ingen licens '$s' i tenanten. Finns: $(($all.skuPartNumber) -join ', ')" }
  $hit[0]
}
function global:CaPolicies([string]$s) {
  $all = $null
  foreach ($x in (Items $s)) {
    if (IsGuid $x) { try { Graph "identity/conditionalAccess/policies/$x"; continue } catch { } }
    if ($null -eq $all) { $all = @(Graph 'identity/conditionalAccess/policies') }
    $hit = @($all | Where-Object { $_.displayName -eq $x })
    if ($hit.Count) { $hit[0] } else { Write-Error "Hittar ingen policy '$x'." }
  }
}
function global:Ref($id) { @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$id" } }
"""

WORKER_PS = r"""
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'
$PSStyle.OutputRendering = 'PlainText'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
[Console]::InputEncoding = [Text.Encoding]::UTF8
$global:EuppaLast = @()
$R = [char]0x1e
function global:Out-Euppa([string]$kind, [string]$text) { [Console]::Out.WriteLine("$R$kind " + ($text -replace "`r?`n", ' ')); [Console]::Out.Flush() }
""" + EU_LIB + r"""
Out-Euppa 'D' '1 0'
while ($true) {
  $buf = [System.Text.StringBuilder]::new()
  while ($true) {
    $l = [Console]::In.ReadLine()
    if ($null -eq $l) { exit 0 }
    if ($l -eq "${R}END") { break }
    [void]$buf.AppendLine($l)
  }
  $text = $buf.ToString()
  $nl = $text.IndexOf("`n")
  $w = [int]$text.Substring(0, $nl).Trim()
  $code = $text.Substring($nl + 1)
  $st = @{ ok = $true }
  $data = [System.Collections.Generic.List[object]]::new()
  try {
    $sb = [scriptblock]::Create($code)
    & $sb 2>&1 3>&1 | ForEach-Object {
      if ($_ -is [System.Management.Automation.ErrorRecord]) { $st.ok = $false; Out-Euppa 'E' $_.ToString() }
      elseif ($_ -is [System.Management.Automation.WarningRecord]) { Out-Euppa 'W' $_.Message }
      elseif ($null -ne $_) { $data.Add($_) }
    }
  } catch { $st.ok = $false; Out-Euppa 'E' (Msg $_) }
  if ($data.Count -and -not ($data.Count -eq 1 -and $data[0] -is [bool])) {
    $global:EuppaLast = $data.ToArray()
    try { $data | Out-String -Stream -Width $w | ForEach-Object { [Console]::Out.WriteLine("${R}O " + $_) } }
    catch { Out-Euppa 'E' $_.ToString() }
  }
  Out-Euppa 'D' ("{0} {1}" -f [int]$st.ok, $data.Count)
}
"""


class Worker:
    """En långlivad pwsh-process: Graph-inloggningen lever kvar mellan åtgärderna.
    conn = (tenantnamn, [scopes]) när vi är inloggade."""

    def __init__(self):
        self.p = None
        self.q = queue.Queue()
        self.conn = None
        self.lock = threading.Lock()
        self.ready = threading.Event()

    def start(self):
        with self.lock:
            if self.p and self.p.poll() is None:
                return
            self.conn = None
            self.ready.clear()
            DIR.mkdir(parents=True, exist_ok=True)
            private(DIR)
            self.q = queue.Queue()
            self.p = subprocess.Popen(ps_args(WORKER_PS, "-NonInteractive"),
                                      stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                      text=True, encoding="utf-8", errors="replace", bufsize=1, **detached())
            q, p = self.q, self.p

            def reader():
                for line in p.stdout:
                    q.put(line.rstrip("\n"))
                q.put(None)
            threading.Thread(target=reader, daemon=True).start()
        while True:                  # första D-raden = redo
            ln = self.q.get()
            if ln is None or ln.startswith(R + "D"):
                break
        self.ready.set()

    def warm(self):
        threading.Thread(target=self.start, daemon=True).start()

    def stop(self):
        if self.p and self.p.poll() is None:
            try:
                self.p.stdin.close()
                self.p.wait(3)
            except Exception:
                self.p.kill()
        self.p = None
        self.conn = None

    def run(self, code, on_event):
        """Kör kod; on_event(kind, text) får H(ost)/W/E/O-rader. -> (ok, antal objekt)."""
        if not self.p or self.p.poll() is not None:
            self.start()
        self.ready.wait(120)
        payload = f"{max(30, width() - 6)}\n{code}\n{R}END\n"
        try:
            self.p.stdin.write(payload)
            self.p.stdin.flush()
        except (BrokenPipeError, OSError, AttributeError):
            on_event("E", "PowerShell-processen har stängts. Försök igen.")
            self.p = None
            return False, 0
        while True:
            ln = self.q.get()
            if ln is None:
                on_event("E", "PowerShell-processen avslutades oväntat.")
                self.p = None
                return False, 0
            if ln.startswith(R):
                kind, _, text = ln[1:].partition(" ")
                if kind == "D":
                    ok, _, n = text.partition(" ")
                    return ok == "1", int(n or 0)
                on_event(kind, text)
            elif ln.startswith(("#< CLIXML", "<Objs ")):
                continue                 # förloppsdata som pwsh skriver till stderr (CLIXML) – inget att visa
            else:
                on_event("H", ln)


WORKER = Worker()


class Runner:
    """Visar maskoten som väntar (blinkar, tittar runt) och de senaste raderna medan något körs."""
    DEVICE = re.compile(r"(https?://\S*(?:devicelogin|deviceauth)\S*).*?\b(?:code|kod)\b\s+([A-Z0-9-]{6,12})", re.I)

    def __init__(self, title):
        self.title = title
        self.lines = []      # (kind, text)
        self.tail = []
        self.msg = None
        self.stop = threading.Event()
        self.lock = threading.Lock()
        self.live = Live()
        self.t0 = time.time()
        self.th = threading.Thread(target=self.loop, daemon=True)

    def event(self, kind, text):
        hidden = text.startswith(("EUPPA-WHO|", "EUPPA-PICK|"))
        with self.lock:
            self.lines.append((kind, text))
            if kind in ("H", "W", "E") and text.strip() and not hidden:
                self.tail = (self.tail + [text])[-3:]
                m = self.DEVICE.search(text)
                if m:
                    code = m.group(2)
                    copied = copy_text(code)
                    self.msg = f"Logga in på {m.group(1).split('//')[-1]} med koden {code}" + \
                        (" (kopierad)." if copied else ".")
                    if not ANIM:
                        out(f"  {LOGO} {self.msg}")
        if not ANIM and kind == "H" and not hidden and not self.DEVICE.search(text):
            out("    " + dim(text))

    def frame_lines(self, mood, lift=0):
        el = int(time.time() - self.t0)
        msg = self.msg or f"{self.title} … {el}s"
        w = width()
        lines = frame(msg, mood, lift=lift, sub=None if self.msg else "Ctrl+C avbryter")
        with self.lock:
            for t in self.tail:
                for l in wrap(t, w - 8, 2):
                    lines.append("    " + dim(l))
        return lines

    def loop(self):
        i = 0
        while not self.stop.is_set():
            el = time.time() - self.t0
            mood = "blink" if el % 3.1 < 0.18 else ("left" if int(el / 1.3) % 4 == 1 else
                                                     "right" if int(el / 1.3) % 4 == 3 else "wait")
            self.live.draw(self.frame_lines(mood, lift=(i // 4) % 2 if not self.msg else 0))
            i += 1
            self.stop.wait(0.11)

    def __enter__(self):
        if ANIM:
            sys.stdout.write("\033[?25l")
            self.th.start()
        else:
            out(f"  {accent(LOGO)} {self.title} …")
        return self

    def __exit__(self, *a):
        self.stop.set()
        if ANIM:
            self.th.join()
            self.live.clear()
            sys.stdout.write("\033[?25h")
            sys.stdout.flush()


def run_ps(title, code, show_output=True, max_lines=None):
    """Kör i arbetsprocessen med väntanimation. -> (ok, antal, rader)."""
    r = Runner(title)
    try:
        with r:
            ok, n = WORKER.run(code, r.event)
    except KeyboardInterrupt:
        out(yellow("  Avbrutet. PowerShell-sessionen startas om (inloggningen nollställs)."))
        WORKER.stop()
        WORKER.warm()
        return False, 0, r.lines
    if show_output:
        print_output(r.lines, max_lines, quiet_host=not ANIM)
    return ok, n, r.lines


HINTS = [
    (r"AADSTS65001|consent|samtycke", "Appen Microsoft Graph Command Line Tools saknar samtycke för det här scopet. "
     "En admin måste godkänna (logga in som Global Admin och svara ja i webbläsaren)."),
    (r"AADSTS50020|AADSTS50034", "Kontot finns inte i den här tenanten. Logga in med ett konto där."),
    (r"AADSTS50076|AADSTS50079|AADSTS50158", "Inloggningen kräver MFA eller ett villkor. Logga in interaktivt igen."),
    (r"Authorization_RequestDenied|Insufficient privileges|\(403\)|Forbidden|Access is denied",
     "Behörighet saknas: kontot har inte rätt adminroll (eller scopet godkändes inte) för det här."),
    (r"Premium|P1|P2|license.*required|tenant is not licensed|RequestFromNonPremiumTenant",
     "Det här kräver Entra ID P1/P2-licens i tenanten (inloggningsloggar, riskanvändare och PIM gör det)."),
    (r"Request_ResourceNotFound|does not exist|Hittar ingen", "Hittade inget med det namnet. Skriv en del av namnet så "
     "söker Euppa fram rätt objekt."),
    (r"is not recognized as (a|the) name of a cmdlet", "Ett kommando saknas – installera modulen "
     "Microsoft.Graph.Authentication (uppa install Microsoft.Graph)."),
    (r"Authentication needed|Connect-MgGraph", "Inte inloggad i Graph. Välj Logga in i menyn (eller: euppa logga-in)."),
]


def friendly(errs):
    text = " ".join(errs)
    for rx, msg in HINTS:
        if re.search(rx, text, re.I):
            return msg
    return ""


def print_output(lines, max_lines=None, quiet_host=False):
    global DIRTY
    DIRTY = True
    w = width()
    lines = [(k, re.split(r"\s+(?:Date|Headers)\s*:", t, 1)[0] if k == "E" else t) for k, t in lines]
    errs = [t for k, t in lines if k == "E"]
    lines = [(k, t) for k, t in lines if not t.startswith(("EUPPA-WHO|", "EUPPA-PICK|"))]
    lines = [(k, t) for k, t in lines
             if not (k in ("H", "W") and any(e[:50] in t or t.replace("WARNING: ", "")[:50] in e for e in errs if e))]
    body = [t for k, t in lines if k == "O"]
    while body and not body[0].strip():
        body.pop(0)
    while body and not body[-1].strip():
        body.pop()
    if not quiet_host:
        for k, t in lines:
            if k == "H" and t.strip():
                for l in wrap(t, w - 6, 6):
                    out("    " + l)
    for k, t in lines:
        if k == "W":
            for i, l in enumerate(wrap(t, w - 8, 4)):
                out(("  " + yellow("! ") if i == 0 else "    ") + yellow(l))
    limit = max_lines or 400
    for l in body[:limit]:
        out("  " + l)
    if len(body) > limit:
        out(dim(f"  … +{len(body) - limit} rader. Exportera till Excel eller CSV för att få allt."))
    for k, t in lines:
        if k == "E":
            for i, l in enumerate(wrap(t, w - 8, 5)):
                out(("  " + red("✘ ") if i == 0 else "    ") + red(l))
    tip = friendly(errs)
    if tip:
        for i, l in enumerate(wrap(tip, w - 8, 4)):
            out(("  " + yellow("→ ") if i == 0 else "    ") + l)


# ───────────────────────────── kommandovisning ─────────────────────────────
TOK = re.compile(r"('(?:[^']|'')*'|\"[^\"]*\"|#.*$|\$[\w:]+|\b[A-Z][a-z]+-[A-Za-z0-9]+\b|-[A-Za-z]+\b)")


def hl(line):
    if not COLOR:
        return line

    def rep(m):
        t = m.group(0)
        if t.startswith("#"):
            return dim(t)
        if t[0] in "'\"":
            return green(t)
        if t.startswith("$"):
            return cyan(t)
        if t.startswith("-"):
            return dim(t)
        return accent(t)
    return TOK.sub(rep, line)


def show_command(code, title="Kommando som körs"):
    """Hela kommandot syns – långa rader bryts (↪) i stället för att klippas."""
    global DIRTY
    DIRTY = True
    w = width()
    inner = max(16, w - 6)
    out("  " + dim("┌ ") + bold(title) + " " + dim("─" * max(0, min(inner, 70) - len(title) - 1)))
    for raw in code.rstrip().splitlines():
        parts = hard_wrap(raw, inner)
        for k, p in enumerate(parts):
            pre = dim("│ ") if k == 0 else dim("│") + dim("↪")
            out("  " + pre + hl(p))
    out("  " + dim("└" + "─" * max(0, min(inner, 70) + 1)))


# ───────────────────────────── inloggning ─────────────────────────────
AUTHS = {"interactive": "Webbläsare (interaktiv)", "device": "Enhetskod (device login)"}
WHO = "EUPPA-WHO|"
WHO_PS = ("$global:EuMe = $null; $global:EuSp = $null; $c = Get-MgContext; $o = $null\n"
          "try { $o = @(Graph 'organization?$select=id,displayName,verifiedDomains')[0] } catch { }\n"
          "$i = ($o.verifiedDomains | Where-Object isInitial | Select-Object -First 1).name\n"
          "$d = ($o.verifiedDomains | Where-Object isDefault | Select-Object -First 1).name\n"
          "Write-Host ('EUPPA-WHO|{0}|{1}|{2}|{3}|{4}' -f $c.TenantId, $c.Account, $i, $o.displayName, $d)")
# Brett läs-paket för "Logga in" i menyn: det mesta i Euppa kan läsas utan ny inloggning.
READ_SCOPES = ["User.Read.All", "Group.Read.All", "Device.Read.All", "Application.Read.All", "Directory.Read.All",
               "Policy.Read.All", "AuditLog.Read.All", "RoleManagement.Read.Directory", "Organization.Read.All",
               "Domain.Read.All", "UserAuthenticationMethod.Read.All"]


def keep_logins():
    """Spara inloggningar mellan körningar (token i nyckelringen / MSAL-cachen)?
    Av som standard på macOS: en osignerad pwsh (Homebrew) får annars en nyckelringsfråga vid varje inloggning."""
    v = config().get("keep_logins")
    return (not IS_MAC) if v is None else bool(v)


def auth_of(t):
    return (t or {}).get("auth") or config().get("auth") or "interactive"


def connect_code(t, scopes):
    """t = vald tenant eller None (då väljer du konto/tenant i inloggningsrutan)."""
    tid = (t.get("tenant_id") or t.get("tenant")) if t else ""
    sc = ",".join(psq(x) for x in sorted(set(scopes or []) | {"User.Read"}))
    return "Connect-MgGraph" + (f" -TenantId {psq(tid)}" if tid else "") + f" -Scopes {sc} -NoWelcome" + \
        ("" if keep_logins() else " -ContextScope Process") + \
        (" -UseDeviceCode" if auth_of(t) == "device" else "") + " -ErrorAction Stop"


def is_connected(t, scopes):
    cur = WORKER.conn
    return bool(cur and t and cur[0] == t["name"] and set(scopes or []) <= set(cur[1]))


def ensure(t, scopes, show=True):
    """Logga in om det behövs. -> tenanten (kan vara en ny, sparad efter inloggningen) eller None."""
    if is_connected(t, scopes):
        return t
    if t and WORKER.conn and WORKER.conn[0] == t["name"]:
        scopes = sorted(set(scopes or []) | set(WORKER.conn[1]))   # be om det nya utan att tappa det gamla
    code = connect_code(t, scopes)
    if show:
        show_command(code, "Loggar in")
    if DRY:
        return t or {}
    title = "Loggar in i Microsoft Graph" + (f" ({t['name']})" if t and t.get("name") else " – välj konto i webbläsaren")
    ok, _, lines = run_ps(title, code + "\n" + WHO_PS, show_output=False)
    if not ok:
        print_output(lines, quiet_host=not ANIM)
        say("Kunde inte logga in i Microsoft Graph. Felet står ovanför.", "sad")
        return None
    who = next((x[1][len(WHO):].split("|") for x in lines if x[0] == "H" and x[1].startswith(WHO)), None)
    if who and any(who):
        before = t["name"] if t else None
        t, new = remember_login(who)
        if before and before != t["name"]:
            out("  " + yellow("!") + f" Du loggade in i {bold(t['name'])}, inte {before}. Bytte tenant.")
        if new:
            out("  " + green("✔") + f" Ny tenant i listan: {bold(t['name'])} " + dim(t.get("tenant", "")))
    if not t:
        return None
    WORKER.conn = (t["name"], sorted(set(scopes or [])))
    out("  " + green("✔") + " Inloggad i Microsoft Graph" + dim(f" · {t['name']}" +
                                                             (f" · {t['upn']}" if t.get("upn") else "")))
    return t


def conn_status():
    t = active_tenant()
    w = width()
    if not t:
        head = f"  {dim('Tenant')} {dim('ingen vald – jag loggar in när en åtgärd behöver det')}"
    else:
        head = f"  {dim('Tenant')} {bold(t['name'])}"
        if w >= 58:
            head += dim(f"  {t.get('tenant', '')}")
        if w >= 96 and t.get("upn"):
            head += dim(f"  ·  {t['upn']}")
    on = bool(t) and WORKER.conn and WORKER.conn[0] == t["name"]
    graph = (green("● ") + "Graph" + dim(f" · {len(WORKER.conn[1])} scopes")) if on else (dim("○ Graph – inte inloggad"))
    return [head, "  " + graph + dim("  ·  " + AUTHS[auth_of(t)].split(" ")[0].lower())]


def logout():
    print()
    code = "try { Disconnect-MgGraph -ErrorAction Stop *> $null; Write-Host 'utloggad' } catch { }"
    show_command("Disconnect-MgGraph", "Loggar ut")
    if DRY:
        return True
    run_ps("Loggar ut från Microsoft Graph", code, show_output=False)
    WORKER.stop()
    WORKER.warm()
    LOOKUP_CACHE.clear()
    say("Utloggad från Microsoft Graph. Nästa gång frågar jag efter inloggning igen. (Webbläsarens "
        "Microsoft-inloggning ligger kvar – logga ut där också om datorn delas.)", "happy")
    return True


# ───────────────────────────── uppslag (sök i tenanten) ─────────────────────────────
# I en åtgärd: "Grupp | Vilken grupp? | | @group"  (@typ+ = flera värden, kommaseparerade)
# Varje uppslag skriver "EUPPA-PICK|värde|namn|detalj" per träff. $Q är det du skrev.
PICK = "EUPPA-PICK|"


LOOKUPS = {
    "user": (["User.Read.All"], "användare", r"""
$s = $Q -replace '"', ''
$u = 'users?$top=40&$count=true&$select=id,displayName,userPrincipalName,accountEnabled,userType,jobTitle'
if ($s) { $u += '&$search=' + (Enc ('"displayName:{0}" OR "userPrincipalName:{0}" OR "mail:{0}"' -f $s)) }
Graph $u -Eventual | ForEach-Object { Pick $_.userPrincipalName $_.displayName (($_.userPrincipalName, $_.jobTitle,
    $(if ($_.userType -eq 'Guest') { 'gäst' }), $(if (-not $_.accountEnabled) { 'spärrad' }) | Where-Object { $_ }) -join ' · ') }
"""),
    "alluser": (["User.Read.All"], "användare", r"""
Graph 'users?$top=999&$select=id,displayName,userPrincipalName,accountEnabled,userType,jobTitle,department' -All |
  Sort-Object displayName | ForEach-Object {
    $d = @($_.userPrincipalName, $_.jobTitle, $_.department) | Where-Object { $_ }
    if ($_.userType -eq 'Guest') { $d += 'gäst' }
    if (-not $_.accountEnabled) { $d += 'spärrad' }
    Pick $_.userPrincipalName $_.displayName ($d -join ' · ') }
"""),
    "group": (["Group.Read.All"], "grupper", r"""
$s = $Q -replace '"', ''
$u = 'groups?$top=40&$count=true&$select=id,displayName,mail,groupTypes,securityEnabled,mailEnabled'
if ($s) { $u += '&$search=' + (Enc ('"displayName:{0}" OR "mail:{0}"' -f $s)) }
Graph $u -Eventual | ForEach-Object { Pick $_.id $_.displayName ((@($_.mail, (GroupType $_)) | Where-Object { $_ }) -join ' · ') }
"""),
    "device": (["Device.Read.All"], "enheter", r"""
$s = $Q -replace '"', ''
$u = 'devices?$top=40&$count=true&$select=id,deviceId,displayName,operatingSystem,accountEnabled,approximateLastSignInDateTime'
if ($s) { $u += '&$search=' + (Enc ('"displayName:{0}"' -f $s)) }
Graph $u -Eventual | ForEach-Object { Pick $_.deviceId $_.displayName ("$($_.operatingSystem) · senast $(D $_.approximateLastSignInDateTime)" + $(if (-not $_.accountEnabled) { ' · inaktiverad' })) }
"""),
    "app": (["Application.Read.All"], "appregistreringar", r"""
$s = $Q -replace '"', ''
$u = 'applications?$top=40&$count=true&$select=id,appId,displayName,createdDateTime'
if ($s) { $u += '&$search=' + (Enc ('"displayName:{0}"' -f $s)) }
Graph $u -Eventual | ForEach-Object { Pick $_.appId $_.displayName ("$($_.appId) · skapad $(D $_.createdDateTime)") }
"""),
    "sp": (["Application.Read.All"], "företagsappar", r"""
$s = $Q -replace '"', ''
$u = 'servicePrincipals?$top=40&$count=true&$select=id,appId,displayName,servicePrincipalType,accountEnabled'
if ($s) { $u += '&$search=' + (Enc ('"displayName:{0}"' -f $s)) }
else { $u += '&$filter=' + (Enc "tags/any(t:t eq 'WindowsAzureActiveDirectoryIntegratedApp')") }
Graph $u -Eventual | ForEach-Object { Pick $_.appId $_.displayName ("$($_.appId)" + $(if (-not $_.accountEnabled) { ' · avstängd' })) }
"""),
    "role": (["RoleManagement.Read.Directory"], "roller", r"""
Graph 'roleManagement/directory/roleDefinitions?$select=id,displayName,description,isBuiltIn' -All | Sort-Object displayName |
  ForEach-Object { Pick $_.id $_.displayName $_.description }
"""),
    "ca": (["Policy.Read.All"], "policyer", r"""
Graph 'identity/conditionalAccess/policies' | Sort-Object displayName | ForEach-Object {
  $l = switch ($_.state) { 'enabled' { 'på' } 'disabled' { 'av' } default { 'report-only' } }
  Pick $_.id $_.displayName $l }
"""),
    "sku": (["Organization.Read.All"], "licenser", r"""
Graph 'subscribedSkus' | Where-Object { $_.prepaidUnits.enabled -gt 0 } | Sort-Object skuPartNumber | ForEach-Object {
  Pick $_.skuPartNumber $_.skuPartNumber ('{0} av {1} använda' -f $_.consumedUnits, $_.prepaidUnits.enabled) }
"""),
    "deleted": (["User.Read.All"], "raderade användare", r"""
Graph 'directory/deletedItems/microsoft.graph.user?$top=999&$select=id,displayName,userPrincipalName,deletedDateTime' -All |
  ForEach-Object { Pick $_.id $_.displayName ("$($_.userPrincipalName) · raderad $(D $_.deletedDateTime)") }
"""),
}
LIST_ALL = {"alluser", "role", "ca", "sku", "deleted"}   # hämtas en gång per session och filtreras här i Euppa
LOOKUP_CACHE = {}
LABELS = {}            # värde -> namn, så att kommandot kan visa "# Säljare" efter ett grupp-id


def run_lookup(kind, t, query):
    """-> (t, [(värde, namn, detalj)]) eller (t, None) om uppslaget inte gick."""
    scopes, label, code = LOOKUPS[kind]
    t2 = ensure(t, scopes, show=False)
    if not t2:
        return t, None
    t = t2
    key = (kind, t["name"])
    if kind in LIST_ALL and key in LOOKUP_CACHE:
        items = LOOKUP_CACHE[key]
    else:
        q = "" if kind in LIST_ALL else query
        ok, _, lines = run_ps(f"Söker {label}" + (f" ”{query}”" if q else ""), f"$Q = {psq(q)}\n{code}",
                              show_output=False)
        items = []
        for k, txt in lines:
            if k == "H" and txt.startswith(PICK):
                f = (txt[len(PICK):].split("|") + ["", "", ""])[:3]
                if f[0]:
                    items.append(tuple(f))
        if not ok and not items:
            print_output(lines, quiet_host=True)
            return t, None
        if kind in LIST_ALL:
            LOOKUP_CACHE[key] = items
    for v, n, _ in items:
        LABELS[v] = n
    ranked = fuzzy_rank(query, items, lambda x: f"{x[1]} {x[0]} {x[2]}") if query else items
    return t, ranked


def lookup_param(prm, t, default):
    """Fråga efter en del av namnet, sök i tenanten och låt användaren välja (flera: kryssa). -> (t, värde)"""
    kind = prm["uppslag"].rstrip("+")
    many = prm["uppslag"].endswith("+")
    label = LOOKUPS[kind][1]
    chosen = []
    q_default = default
    while True:
        what = prm["fraga"] + (" – fler" if many and chosen else "")
        q = ask_text(what + (" (valfritt)" if prm["valfri"] else "") + " · skriv en del av namnet" +
                     (" (tomt = lista alla)" if kind in LIST_ALL and not prm["valfri"] else ""),
                     q_default, required=not prm["valfri"] and not chosen and kind not in LIST_ALL)
        q_default = ""
        if not q and (prm["valfri"] or chosen):
            break
        if q and ("@" in q and " " not in q and kind == "user" or GUID.match(q)):
            vals = [q]                                       # fullständig adress / id: använd direkt
        else:
            t, hits = run_lookup(kind, t, q)
            if hits is None:
                if not q:
                    raise Cancel("kunde inte söka")
                out("    " + yellow(f"Kunde inte söka – använder ”{q}” som det är."))
                vals = [q]
            elif not hits:
                out("    " + yellow(f"Hittade inga {label} som matchar ”{q}”."))
                v = choose([dict(label="Sök igen", value="__again")] +
                           ([dict(label=f"Använd ”{q}” som det är", value=q)] if q else []), title="Vad vill du göra?")
                if v is None:
                    raise Cancel()
                if v == "__again":
                    continue
                vals = [v]
            elif len(hits) == 1 or len([h for h in hits if q and q.lower() in (h[0].lower(), h[1].lower())]) == 1:
                exact = [h for h in hits if q and q.lower() in (h[0].lower(), h[1].lower())]
                h = (exact or hits)[0]
                vals = [h[0]]
                out("    " + green("→ ") + bold(h[1]) + dim(f"  {h[0]}" if h[0] != h[1] else ""))
            elif many:
                vals = pick_many(hits, f"{len(hits)} {label}" + (f" som matchar ”{q}”" if q else ""))
                if not vals:
                    raise Cancel()
                out("    " + green("→ ") + ", ".join(bold(LABELS.get(v, v)) for v in vals))
            else:
                rows = [dict(label=h[1], hint=h[2] or h[0], value=h[0]) for h in hits[:300]]
                rows += [("head", ""), dict(label="Sök igen…", value="__again", key="s")]
                if q:
                    rows.append(dict(label=f"Använd ”{q}” som det är", value="__raw", key="r"))
                v = choose(rows, title=f"{len(hits)} {label}" + (f" matchar ”{q}”" if q else "") +
                           (" – skriv för att filtrera" if len(hits) > 9 else ""))
                if v is None:
                    raise Cancel()
                if v == "__again":
                    continue
                vals = [q if v == "__raw" else v]
        chosen += [v for v in vals if v not in chosen]
        if not many or not confirm("Lägg till fler?", False):
            break
    return t, ",".join(chosen)


def cmd_lookup(kind, query, as_json, all_):
    """euppa uppslag TYP [sök]: sök i tenanten (samma uppslag som åtgärderna använder). --json för Ruppa."""
    if all_ and kind == "user":
        kind = "alluser"
    if kind not in LOOKUPS:
        raise Cancel(f"okänt uppslag '{kind}' (finns: {', '.join(k for k in LOOKUPS if k != 'alluser')})")
    t, hits = run_lookup(kind, active_tenant(), query or "")
    if hits is None:
        if as_json:
            emit_json({"ok": False, "error": "kunde inte söka (inloggad?)", "items": []})
        return False
    if as_json:
        emit_json({"ok": True, "tenant": (t or {}).get("name", ""),
                   "items": [{"value": v, "name": n, "detail": d} for v, n, d in hits]})
    else:
        w = width()
        for v, n, d in hits:
            out(fit(f"  {bold(n)}  {dim(d or v)}", w))
    return True


REAL_STDOUT = sys.stdout
EMITTED = False


def emit_json(obj):
    """--json: en enda rad JSON på riktiga stdout (allt annat har skickats till stderr)."""
    global EMITTED
    EMITTED = True
    REAL_STDOUT.write(json.dumps(obj, ensure_ascii=False, default=str) + "\n")
    REAL_STDOUT.flush()


# ───────────────────────────── åtgärder ─────────────────────────────
def prm(spec):
    """'Namn? | Fråga | standard | @user+'  eller  '… | a, b, c'  -> dict"""
    f = [x.strip() for x in spec.split("|")]
    name = f[0]
    spec4 = f[3] if len(f) > 3 else ""
    look = spec4[1:] if spec4.startswith("@") else ""
    return dict(namn=name.rstrip("?"), fraga=f[1] if len(f) > 1 and f[1] else name.rstrip("?"),
                standard=f[2] if len(f) > 2 else "", valfri=name.endswith("?"), uppslag=look,
                val=[] if look else [x.strip() for x in spec4.split(",") if x.strip()])


def act(id_, namn, omrade, typ, scopes, beskrivning, params, ps):
    return dict(id=id_, namn=namn, omrade=omrade, typ=typ, scopes=scopes.split(), beskrivning=beskrivning,
                params=[prm(p) for p in params], ps=ps.strip("\n"))


# Områdena, i menyordning. farg = xterm-256 (samma palett som släkten), ikon = två bokstäver i Ruppa.
AREAS = [
    dict(namn="Användare", ikon="AN", farg=49, beskrivning="Skapa, ändra, spärra, lösenord, MFA, licenser och chef"),
    dict(namn="Gäster", ikon="GÄ", farg=80, beskrivning="Bjud in, se status och städa bland externa"),
    dict(namn="Grupper", ikon="GR", farg=177, beskrivning="Skapa, medlemmar, ägare och dynamiska regler"),
    dict(namn="Enheter", ikon="EN", farg=75, beskrivning="Lista, inaktivera, BitLocker-nycklar och LAPS"),
    dict(namn="Appar", ikon="AP", farg=141, beskrivning="Appregistreringar och företagsappar: hemligheter, behörigheter, samtycke"),
    dict(namn="Roller och PIM", ikon="RO", farg=214, beskrivning="Adminroller, berättigade och aktiva tilldelningar"),
    dict(namn="Villkorlig åtkomst", ikon="CA", farg=203, beskrivning="Policyer, report-only, undantag och namngivna platser"),
    dict(namn="Loggar", ikon="LO", farg=110, beskrivning="Inloggningar och granskningslogg"),
    dict(namn="Säkerhet", ikon="SÄ", farg=197, beskrivning="Riskanvändare, riskhändelser och MFA-status"),
    dict(namn="Tenant", ikon="TE", farg=220, beskrivning="Licenser, domäner, företagsinfo och inställningar"),
    dict(namn="Lokalt", ikon="LK", farg=245, beskrivning="Sessionen och modulerna på den här datorn"),
]

U = "Anvandare | Vilka användare? | | @user+"
ACTIONS = [
    # ─────────── Användare ───────────
    act("anvandare-lista", "Lista användare", "Användare", "läs", "User.Read.All",
        "Alla användare med status, typ, titel, avdelning och antal licenser.",
        ["Sok? | Filtrera (namn, e-post, titel, avdelning) |"], r"""
$all = @(Graph 'users?$top=999&$select=id,displayName,userPrincipalName,accountEnabled,userType,jobTitle,department,createdDateTime,assignedLicenses' -All)
if ($Sok) { $all = @($all | Where-Object { "$($_.displayName) $($_.userPrincipalName) $($_.jobTitle) $($_.department)" -like "*$Sok*" }) }
$all | Sort-Object displayName | ForEach-Object { [pscustomobject][ordered]@{
  Namn = $_.displayName; UPN = $_.userPrincipalName; Aktiv = JaNej $_.accountEnabled
  Typ = $(if ($_.userType -eq 'Guest') { 'gäst' } else { 'medlem' }); Titel = $_.jobTitle; Avdelning = $_.department
  Licenser = @($_.assignedLicenses).Count; Skapad = D $_.createdDateTime } }
"""),
    act("anvandare-info", "Användarinfo", "Användare", "läs", "User.Read.All AuditLog.Read.All",
        "Allt om en eller flera användare: kontakt, chef, status, synk, licenser, grupper och senaste inloggning.",
        [U], r"""
$sel = 'id,displayName,userPrincipalName,mail,jobTitle,department,officeLocation,mobilePhone,companyName,employeeId,accountEnabled,userType,createdDateTime,usageLocation,onPremisesSyncEnabled,onPremisesLastSyncDateTime,lastPasswordChangeDateTime'
foreach ($u in Users $Anvandare) {
  $x = $null
  try { $x = Graph ("users/$($u.id)" + '?$select=' + $sel + ',signInActivity') } catch { $x = Graph ("users/$($u.id)" + '?$select=' + $sel) }
  $chef = ''
  try { $chef = (Graph ("users/$($u.id)/manager" + '?$select=displayName,userPrincipalName')).userPrincipalName } catch { }
  $lic = @(Graph ("users/$($u.id)/licenseDetails" + '?$select=skuPartNumber') | ForEach-Object { $_.skuPartNumber }) -join ', '
  $grp = @(Graph ("users/$($u.id)/memberOf" + '?$select=id&$top=999') -All).Count
  [pscustomobject][ordered]@{
    Namn = $x.displayName; UPN = $x.userPrincipalName; Epost = $x.mail; Titel = $x.jobTitle; Avdelning = $x.department
    Kontor = $x.officeLocation; Mobil = $x.mobilePhone; Foretag = $x.companyName; AnstallningsId = $x.employeeId; Chef = $chef
    Aktiv = JaNej $x.accountEnabled; Typ = $(if ($x.userType -eq 'Guest') { 'gäst' } else { 'medlem' })
    Kalla = $(if ($x.onPremisesSyncEnabled) { "lokal AD (synk $(D $x.onPremisesLastSyncDateTime))" } else { 'molnet' })
    Plats = $x.usageLocation; Licenser = $lic; Grupper = $grp; Skapad = D $x.createdDateTime
    LosenordBytt = D $x.lastPasswordChangeDateTime; SenastInloggad = D $x.signInActivity.lastSignInDateTime }
}
"""),
    act("anvandare-inaktiva", "Inaktiva användare", "Användare", "läs", "User.Read.All AuditLog.Read.All",
        "Konton som inte loggat in på X dagar (eller aldrig). Kräver Entra ID P1 för inloggningsdatum.",
        ["Dagar | Ingen inloggning på hur många dagar? | 90", "Typ | Vilka? | Alla | Alla, Medlemmar, Gäster"], r"""
$gr = (Get-Date).AddDays(-[int]$Dagar)
$rows = @(Graph 'users?$top=500&$select=displayName,userPrincipalName,accountEnabled,userType,createdDateTime,signInActivity,assignedLicenses' -All)
if ($Typ -eq 'Medlemmar') { $rows = @($rows | Where-Object userType -ne 'Guest') } elseif ($Typ -eq 'Gäster') { $rows = @($rows | Where-Object userType -eq 'Guest') }
$rows | ForEach-Object {
  $l = $_.signInActivity.lastSignInDateTime
  if ((-not $l -and [datetime]$_.createdDateTime -lt $gr) -or ($l -and [datetime]$l -lt $gr)) {
    [pscustomobject][ordered]@{ Namn = $_.displayName; UPN = $_.userPrincipalName; Aktiv = JaNej $_.accountEnabled
      Typ = $(if ($_.userType -eq 'Guest') { 'gäst' } else { 'medlem' }); SenastInloggad = $(if ($l) { D $l } else { 'aldrig' })
      Dagar = $(if ($l) { Ago $l } else { Ago $_.createdDateTime }); Licenser = @($_.assignedLicenses).Count; Skapad = D $_.createdDateTime }
  } } | Sort-Object Dagar -Descending
"""),
    act("anvandare-skapa", "Skapa användare", "Användare", "ändrar", "User.ReadWrite.All",
        "Ny molnanvändare. Lämnar du lösenordet tomt hittar Euppa på ett starkt och visar det en gång.",
        ["Namn | Visningsnamn (t.ex. Anna Svensson) |", "Anvandarnamn? | Användarnamn före @ (tomt = fornamn.efternamn) |",
         "Doman | Domän | {domain}", "Losenord? | Lösenord (tomt = ett nytt, starkt) |", "Titel? | Titel |",
         "Avdelning? | Avdelning |", "Plats | Användningsplats (landskod, behövs för licenser) | SE",
         "Byt | Byt lösenord vid första inloggningen? | Ja | Ja, Nej"], r"""
$nick = $Anvandarnamn
if (-not $nick) { $nick = ($Namn.Trim().ToLower() -replace '[åäà]', 'a' -replace '[öø]', 'o' -replace '[éè]', 'e' -replace 'ü', 'u' -replace '\s+', '.' -replace '[^a-z0-9._-]', '') }
$pw = $Losenord
if (-not $pw) { $pw = NewPw }
$body = @{ accountEnabled = $true; displayName = $Namn.Trim(); mailNickname = ($nick -replace '[^A-Za-z0-9._-]', '')
  userPrincipalName = "$nick@$Doman"; usageLocation = $Plats.ToUpper()
  passwordProfile = @{ password = $pw; forceChangePasswordNextSignIn = ($Byt -eq 'Ja') } }
$parts = @($Namn.Trim() -split '\s+', 2)
$body.givenName = $parts[0]
if ($parts.Count -gt 1) { $body.surname = $parts[1] }
if ($Titel) { $body.jobTitle = $Titel }
if ($Avdelning) { $body.department = $Avdelning }
$n = Graph 'users' -Method POST -Body $body
[pscustomobject][ordered]@{ Namn = $n.displayName; UPN = $n.userPrincipalName; Losenord = $pw; BytVidInloggning = $Byt; Id = $n.id }
"""),
    act("anvandare-andra", "Ändra fält på användare", "Användare", "ändrar", "User.ReadWrite.All",
        "Sätt titel, avdelning, kontor, mobil, företag m.m. på en eller flera användare. Tomt värde rensar fältet.",
        [U, "Falt | Vilket fält? | Titel | Visningsnamn, Förnamn, Efternamn, Titel, Avdelning, Kontor, Mobil, Företag, Anställnings-id, Användningsplats",
         "Varde? | Nytt värde (tomt = rensa) |"], r"""
$map = @{ 'Visningsnamn' = 'displayName'; 'Förnamn' = 'givenName'; 'Efternamn' = 'surname'; 'Titel' = 'jobTitle'
  'Avdelning' = 'department'; 'Kontor' = 'officeLocation'; 'Mobil' = 'mobilePhone'; 'Företag' = 'companyName'
  'Anställnings-id' = 'employeeId'; 'Användningsplats' = 'usageLocation' }
$prop = $map[$Falt]
if (-not $prop) { throw "Okänt fält: $Falt" }
$v = $null
if ($Varde) { $v = $Varde }
foreach ($u in Users $Anvandare) {
  try { Graph "users/$($u.id)" -Method PATCH -Body @{ $prop = $v } | Out-Null; Done $u.userPrincipalName "$Falt = $(if ($v) { $v } else { '(tomt)' })" }
  catch { Fail $u.userPrincipalName $_ }
}
"""),
    act("anvandare-sparra", "Spärra och logga ut", "Användare", "ändrar", "User.ReadWrite.All",
        "Stänger av inloggningen och återkallar alla sessioner (token) direkt. Ditt eget konto hoppas över.",
        [U], r"""
$me = (Me).id
foreach ($u in Users $Anvandare) {
  if ($u.id -eq $me) { Fail $u.userPrincipalName 'det är ditt eget konto – hoppar över'; continue }
  try {
    Graph "users/$($u.id)" -Method PATCH -Body @{ accountEnabled = $false } | Out-Null
    Graph "users/$($u.id)/revokeSignInSessions" -Method POST | Out-Null
    Done $u.userPrincipalName 'spärrad och utloggad överallt'
  } catch { Fail $u.userPrincipalName $_ }
}
"""),
    act("anvandare-aktivera", "Aktivera konto", "Användare", "ändrar", "User.ReadWrite.All",
        "Slår på inloggningen igen för spärrade konton.", [U], r"""
foreach ($u in Users $Anvandare) {
  try { Graph "users/$($u.id)" -Method PATCH -Body @{ accountEnabled = $true } | Out-Null; Done $u.userPrincipalName 'aktiverad' }
  catch { Fail $u.userPrincipalName $_ }
}
"""),
    act("anvandare-losenord", "Återställ lösenord", "Användare", "ändrar", "User.ReadWrite.All",
        "Nytt lösenord för en eller flera användare. Tomt = ett unikt, starkt lösenord per person som visas en gång.",
        [U, "Losenord? | Nytt lösenord (tomt = unikt per användare) |",
         "Byt | Byt lösenord vid nästa inloggning? | Ja | Ja, Nej"], r"""
foreach ($u in Users $Anvandare) {
  $pw = $Losenord
  if (-not $pw) { $pw = NewPw }
  try {
    Graph "users/$($u.id)" -Method PATCH -Body @{ passwordProfile = @{ password = $pw; forceChangePasswordNextSignIn = ($Byt -eq 'Ja') } } | Out-Null
    [pscustomobject][ordered]@{ Objekt = $u.userPrincipalName; Resultat = 'nytt lösenord'; Losenord = $pw; BytVidInloggning = $Byt }
  } catch { Fail $u.userPrincipalName $_ }
}
"""),
    act("anvandare-logga-ut", "Logga ut överallt", "Användare", "ändrar", "User.ReadWrite.All",
        "Återkallar alla sessioner och uppdateringstoken. Användaren måste logga in igen på alla enheter.", [U], r"""
foreach ($u in Users $Anvandare) {
  try { Graph "users/$($u.id)/revokeSignInSessions" -Method POST | Out-Null; Done $u.userPrincipalName 'utloggad överallt' }
  catch { Fail $u.userPrincipalName $_ }
}
"""),
    act("anvandare-mfa", "MFA-metoder", "Användare", "läs", "UserAuthenticationMethod.Read.All User.Read.All",
        "Vilka inloggningsmetoder användarna har registrerat: Authenticator, telefon, FIDO2-nyckel, Windows Hello …",
        [U], r"""
$names = @{ 'password' = 'Lösenord'; 'microsoftAuthenticator' = 'Microsoft Authenticator'; 'phone' = 'Telefon'
  'fido2' = 'FIDO2-nyckel'; 'softwareOath' = 'OATH-app (TOTP)'; 'email' = 'E-post (SSPR)'; 'windowsHelloForBusiness' = 'Windows Hello'
  'temporaryAccessPass' = 'Tillfälligt lösenord (TAP)'; 'platformCredential' = 'Plattform (macOS)'; 'passwordlessMicrosoftAuthenticator' = 'Authenticator (lösenordsfri)' }
foreach ($u in Users $Anvandare) {
  foreach ($m in @(Graph "users/$($u.id)/authentication/methods")) {
    $k = ($m.'@odata.type' -replace '#microsoft.graph.', '') -replace 'AuthenticationMethod$', ''
    $det = @($m.phoneNumber, $m.displayName, $m.emailAddress, $m.model, $m.deviceTag) | Where-Object { $_ } | Select-Object -First 1
    [pscustomobject][ordered]@{ Anvandare = $u.userPrincipalName; Metod = $(if ($names[$k]) { $names[$k] } else { $k })
      Detalj = "$det"; Typ = "$($m.phoneType)"; Skapad = D $m.createdDateTime }
  }
}
"""),
    act("anvandare-mfa-aterstall", "Återställ MFA", "Användare", "ändrar", "UserAuthenticationMethod.ReadWrite.All User.ReadWrite.All",
        "Tar bort alla registrerade MFA-metoder (lösenordet behålls) och loggar ut. Användaren registrerar sig igen vid nästa inloggning.",
        [U], r"""
$seg = @{ '#microsoft.graph.phoneAuthenticationMethod' = 'phoneMethods'
  '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' = 'microsoftAuthenticatorMethods'
  '#microsoft.graph.fido2AuthenticationMethod' = 'fido2Methods'; '#microsoft.graph.softwareOathAuthenticationMethod' = 'softwareOathMethods'
  '#microsoft.graph.emailAuthenticationMethod' = 'emailMethods'; '#microsoft.graph.temporaryAccessPassAuthenticationMethod' = 'temporaryAccessPassMethods'
  '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod' = 'windowsHelloForBusinessMethods'
  '#microsoft.graph.platformCredentialAuthenticationMethod' = 'platformCredentialMethods' }
foreach ($u in Users $Anvandare) {
  $n = 0; $left = @()
  for ($round = 0; $round -lt 3; $round++) {      # standardmetoden går bara att ta bort sist – därför flera varv
    $left = @()
    foreach ($m in @(Graph "users/$($u.id)/authentication/methods")) {
      $s = $seg[$m.'@odata.type']
      if (-not $s) { continue }
      try { Graph "users/$($u.id)/authentication/$s/$($m.id)" -Method DELETE | Out-Null; $n++ }
      catch { $left += (($m.'@odata.type' -replace '#microsoft.graph.', '') + ': ' + (Msg $_)) }
    }
    if (-not $left.Count) { break }
  }
  try { Graph "users/$($u.id)/revokeSignInSessions" -Method POST | Out-Null } catch { }
  if ($left.Count) { Fail $u.userPrincipalName ("$n borttagna, kvar: " + ($left -join '; ')) }
  else { Done $u.userPrincipalName "$n metoder borttagna och utloggad – registrerar MFA igen vid nästa inloggning" }
}
"""),
    act("anvandare-tap", "Tillfälligt lösenord (TAP)", "Användare", "ändrar", "UserAuthenticationMethod.ReadWrite.All User.Read.All",
        "Skapar ett Temporary Access Pass, t.ex. för att registrera MFA eller en ny dator. Kräver att TAP är påslaget i tenanten.",
        [U, "Minuter | Giltigt i minuter | 60", "Engang | Bara en gång? | Ja | Ja, Nej"], r"""
foreach ($u in Users $Anvandare) {
  try {
    $r = Graph "users/$($u.id)/authentication/temporaryAccessPassMethods" -Method POST -Body @{ lifetimeInMinutes = [int]$Minuter; isUsableOnce = ($Engang -eq 'Ja') }
    [pscustomobject][ordered]@{ Objekt = $u.userPrincipalName; Resultat = 'TAP skapat'; Kod = $r.temporaryAccessPass
      GiltigTill = D ([datetime]$r.startDateTime).AddMinutes([int]$r.lifetimeInMinutes); Engang = $Engang }
  } catch { Fail $u.userPrincipalName $_ }
}
"""),
    act("anvandare-licenser", "Användarens licenser", "Användare", "läs", "User.Read.All",
        "Licenser per användare och om de är direkt tilldelade eller kommer via en grupp.", [U], r"""
foreach ($u in Users $Anvandare) {
  $x = Graph ("users/$($u.id)" + '?$select=licenseAssignmentStates')
  $sku = @{}; foreach ($d in @(Graph "users/$($u.id)/licenseDetails")) { $sku[$d.skuId] = $d.skuPartNumber }
  foreach ($s in @($x.licenseAssignmentStates)) {
    [pscustomobject][ordered]@{ Anvandare = $u.userPrincipalName; Licens = $(if ($sku[$s.skuId]) { $sku[$s.skuId] } else { $s.skuId })
      Via = $(if ($s.assignedByGroup) { 'grupp' } else { 'direkt' }); Status = $s.state; Fel = $s.error; Avstangt = @($s.disabledPlans).Count }
  }
}
"""),
    act("anvandare-licens-tilldela", "Tilldela licens", "Användare", "ändrar", "User.ReadWrite.All Organization.Read.All",
        "Ger användarna en licens. Saknar någon användningsplats sätts den först (krävs för licenser).",
        [U, "Licens | Vilken licens? | | @sku", "Plats | Användningsplats om den saknas | SE"], r"""
$sku = Sku $Licens
foreach ($u in Users $Anvandare) {
  try {
    if (-not $u.usageLocation) { Graph "users/$($u.id)" -Method PATCH -Body @{ usageLocation = $Plats.ToUpper() } | Out-Null }
    Graph "users/$($u.id)/assignLicense" -Method POST -Body @{ addLicenses = @(@{ skuId = $sku.skuId; disabledPlans = @() }); removeLicenses = @() } | Out-Null
    Done $u.userPrincipalName "fick $($sku.skuPartNumber)"
  } catch { Fail $u.userPrincipalName $_ }
}
"""),
    act("anvandare-licens-ta-bort", "Ta bort licens", "Användare", "ändrar", "User.ReadWrite.All Organization.Read.All",
        "Tar bort en direkt tilldelad licens. Licenser som kommer via grupp tas bort genom att ta bort användaren ur gruppen.",
        [U, "Licens | Vilken licens? | | @sku"], r"""
$sku = Sku $Licens
foreach ($u in Users $Anvandare) {
  try {
    Graph "users/$($u.id)/assignLicense" -Method POST -Body @{ addLicenses = @(); removeLicenses = @($sku.skuId) } | Out-Null
    Done $u.userPrincipalName "$($sku.skuPartNumber) borttagen"
  } catch { Fail $u.userPrincipalName $_ }
}
"""),
    act("anvandare-chef", "Sätt chef", "Användare", "ändrar", "User.ReadWrite.All",
        "Sätter chefen (manager) för en eller flera användare.", [U, "Chef | Vem är chef? | | @user"], r"""
$c = @(Users $Chef)[0]
if (-not $c) { throw "Hittar ingen chef '$Chef'." }
foreach ($u in Users $Anvandare) {
  try { Graph ('users/' + $u.id + '/manager/$ref') -Method PUT -Body @{ '@odata.id' = "https://graph.microsoft.com/v1.0/users/$($c.id)" } | Out-Null
        Done $u.userPrincipalName "chef: $($c.displayName)" }
  catch { Fail $u.userPrincipalName $_ }
}
"""),
    act("anvandare-grupper", "Användarens grupper och roller", "Användare", "läs", "User.Read.All GroupMember.Read.All",
        "Grupper, adminroller och administrativa enheter som användarna är direkt medlemmar i.", [U], r"""
foreach ($u in Users $Anvandare) {
  foreach ($g in @(Graph ("users/$($u.id)/memberOf" + '?$top=999&$select=id,displayName,mail,groupTypes,securityEnabled,mailEnabled,membershipRule') -All)) {
    $typ = UserType $g
    if ($typ -eq 'grupp') { $typ = GroupType $g }
    [pscustomobject][ordered]@{ Anvandare = $u.userPrincipalName; Namn = $g.displayName; Typ = $typ; Epost = $g.mail
      Medlemskap = $(if ($g.membershipRule) { 'dynamiskt' } else { 'tilldelat' }) }
  }
}
"""),
    act("anvandare-enheter", "Användarens enheter", "Användare", "läs", "User.Read.All Device.Read.All",
        "Enheter som användarna äger eller har registrerat.", [U], r"""
foreach ($u in Users $Anvandare) {
  $seen = @{}
  foreach ($rel in 'ownedDevices', 'registeredDevices') {
    foreach ($d in @(Graph ("users/$($u.id)/$rel" + '?$select=id,deviceId,displayName,operatingSystem,operatingSystemVersion,trustType,accountEnabled,approximateLastSignInDateTime,isCompliant') -All)) {
      if ($seen[$d.id]) { continue }; $seen[$d.id] = 1
      [pscustomobject][ordered]@{ Anvandare = $u.userPrincipalName; Enhet = $d.displayName; OS = "$($d.operatingSystem) $($d.operatingSystemVersion)"
        Anslutning = $(switch ($d.trustType) { 'AzureAd' { 'Entra-ansluten' } 'ServerAd' { 'hybrid' } 'Workplace' { 'registrerad' } default { $d.trustType } })
        Aktiv = JaNej $d.accountEnabled; Kompatibel = JaNej $d.isCompliant; Senast = D $d.approximateLastSignInDateTime; DeviceId = $d.deviceId }
    }
  }
}
"""),
    act("anvandare-ta-bort", "Ta bort användare", "Användare", "ändrar", "User.ReadWrite.All",
        "Raderar användarna. De ligger kvar under Raderade användare i 30 dagar och kan återställas. Ditt eget konto hoppas över.",
        [U], r"""
$me = (Me).id
foreach ($u in Users $Anvandare) {
  if ($u.id -eq $me) { Fail $u.userPrincipalName 'det är ditt eget konto – hoppar över'; continue }
  try { Graph "users/$($u.id)" -Method DELETE | Out-Null; Done $u.userPrincipalName 'raderad (kan återställas i 30 dagar)' }
  catch { Fail $u.userPrincipalName $_ }
}
"""),
    act("anvandare-raderade", "Raderade användare", "Användare", "läs", "User.Read.All",
        "Användare i papperskorgen och hur många dagar som är kvar innan de försvinner för gott.", [], r"""
Graph 'directory/deletedItems/microsoft.graph.user?$top=999&$select=id,displayName,userPrincipalName,mail,deletedDateTime' -All |
  Sort-Object deletedDateTime -Descending | ForEach-Object { [pscustomobject][ordered]@{ Namn = $_.displayName; UPN = $_.userPrincipalName
    Raderad = D $_.deletedDateTime; DagarKvar = [math]::Max(0, 30 - [int](Ago $_.deletedDateTime)); Id = $_.id } }
"""),
    act("anvandare-aterstall", "Återställ raderad användare", "Användare", "ändrar", "User.ReadWrite.All",
        "Tar tillbaka raderade användare med licenser och gruppmedlemskap.",
        ["Raderad | Vilka raderade användare? | | @deleted+"], r"""
foreach ($id in Items $Raderad) {
  try { $r = Graph "directory/deletedItems/$id/restore" -Method POST; Done "$($r.userPrincipalName)" 'återställd' }
  catch { Fail $id $_ }
}
"""),
    # ─────────── Gäster ───────────
    act("gaster-lista", "Lista gäster", "Gäster", "läs", "User.Read.All AuditLog.Read.All",
        "Alla gästkonton: inbjudan accepterad eller inte, när de bjöds in och senaste inloggning.", [], r"""
$sel = 'displayName,mail,userPrincipalName,createdDateTime,externalUserState,externalUserStateChangeDateTime,accountEnabled'
$f = '&$filter=' + (Enc "userType eq 'Guest'")
$rows = $null
try { $rows = @(Graph ('users?$top=500&$select=' + $sel + ',signInActivity' + $f) -All) } catch { $rows = @(Graph ('users?$top=999&$select=' + $sel + $f) -All) }
$rows | Sort-Object createdDateTime -Descending | ForEach-Object { [pscustomobject][ordered]@{
  Namn = $_.displayName; Epost = $_.mail; Status = $(if ($_.externalUserState -eq 'PendingAcceptance') { 'väntar' } else { 'accepterad' })
  Inbjuden = D $_.createdDateTime; SenastInloggad = D $_.signInActivity.lastSignInDateTime; Aktiv = JaNej $_.accountEnabled; UPN = $_.userPrincipalName } }
"""),
    act("gaster-ej-accepterade", "Gäster som inte accepterat", "Gäster", "läs", "User.Read.All",
        "Inbjudningar som legat oaccepterade längre än X dagar – bra att städa bort.",
        ["Dagar | Äldre än hur många dagar? | 30"], r"""
Graph ('users?$top=999&$select=displayName,mail,userPrincipalName,createdDateTime,externalUserState&$filter=' + (Enc "userType eq 'Guest'")) -All |
  Where-Object { $_.externalUserState -eq 'PendingAcceptance' -and (Ago $_.createdDateTime) -ge [int]$Dagar } |
  ForEach-Object { [pscustomobject][ordered]@{ Namn = $_.displayName; Epost = $_.mail; Inbjuden = D $_.createdDateTime; Dagar = Ago $_.createdDateTime; UPN = $_.userPrincipalName } }
"""),
    act("gaster-bjud-in", "Bjud in gäst", "Gäster", "ändrar", "User.Invite.All",
        "Bjuder in en eller flera externa (kommaseparerade e-postadresser) och visar inlösningslänken.",
        ["Epost | E-postadress(er) |", "Namn? | Visningsnamn (bara vid en adress) |",
         "Skicka | Skicka inbjudan via e-post? | Ja | Ja, Nej", "Lank | Skicka gästen hit efteråt | https://myapps.microsoft.com",
         "Meddelande? | Eget meddelande i mejlet |"], r"""
foreach ($e in Items $Epost) {
  $b = @{ invitedUserEmailAddress = $e; inviteRedirectUrl = $Lank; sendInvitationMessage = ($Skicka -eq 'Ja') }
  if ($Namn -and @(Items $Epost).Count -eq 1) { $b.invitedUserDisplayName = $Namn }
  if ($Meddelande) { $b.invitedUserMessageInfo = @{ customizedMessageBody = $Meddelande } }
  try { $r = Graph 'invitations' -Method POST -Body $b
        [pscustomobject][ordered]@{ Objekt = $e; Resultat = "inbjuden ($($r.status))"; Lank = $r.inviteRedeemUrl; Id = $r.invitedUser.id } }
  catch { Fail $e $_ }
}
"""),
    # ─────────── Grupper ───────────
    act("grupper-lista", "Lista grupper", "Grupper", "läs", "Group.Read.All",
        "Alla grupper med typ (Microsoft 365, säkerhet, distribution), dynamiskt eller tilldelat medlemskap och e-post.",
        ["Sok? | Filtrera på namn eller e-post |"], r"""
$all = @(Graph 'groups?$top=999&$select=id,displayName,mail,groupTypes,securityEnabled,mailEnabled,membershipRule,createdDateTime,description,onPremisesSyncEnabled' -All)
if ($Sok) { $all = @($all | Where-Object { "$($_.displayName) $($_.mail)" -like "*$Sok*" }) }
$all | Sort-Object displayName | ForEach-Object { [pscustomobject][ordered]@{ Namn = $_.displayName; Typ = GroupType $_
  Medlemskap = $(if ($_.groupTypes -contains 'DynamicMembership') { 'dynamiskt' } else { 'tilldelat' }); Epost = $_.mail
  Kalla = $(if ($_.onPremisesSyncEnabled) { 'lokal AD' } else { 'molnet' }); Skapad = D $_.createdDateTime; Beskrivning = $_.description; Id = $_.id } }
"""),
    act("grupper-medlemmar", "Gruppmedlemmar", "Grupper", "läs", "GroupMember.Read.All User.Read.All",
        "Vilka som är med i gruppen – direkt, eller även via undergrupper.",
        ["Grupp | Vilken grupp? | | @group", "Alla | Även via undergrupper? | Nej | Nej, Ja"], r"""
$rel = $(if ($Alla -eq 'Ja') { 'transitiveMembers' } else { 'members' })
foreach ($g in Groups $Grupp) {
  Graph ("groups/$($g.id)/$rel" + '?$top=999&$select=id,displayName,userPrincipalName,mail,accountEnabled') -All | ForEach-Object {
    [pscustomobject][ordered]@{ Grupp = $g.displayName; Namn = $_.displayName; Typ = UserType $_; Konto = $(if ($_.userPrincipalName) { $_.userPrincipalName } else { $_.mail }); Aktiv = JaNej ($_.accountEnabled -ne $false) } } |
    Sort-Object Typ, Namn
}
"""),
    act("grupper-agare", "Gruppens ägare", "Grupper", "läs", "Group.Read.All", "Ägarna till en eller flera grupper.",
        ["Grupp | Vilka grupper? | | @group+"], r"""
foreach ($g in Groups $Grupp) {
  $o = @(Graph ("groups/$($g.id)/owners" + '?$select=id,displayName,userPrincipalName'))
  if (-not $o.Count) { [pscustomobject][ordered]@{ Grupp = $g.displayName; Agare = '(ingen ägare)'; Konto = '' } }
  $o | ForEach-Object { [pscustomobject][ordered]@{ Grupp = $g.displayName; Agare = $_.displayName; Konto = $_.userPrincipalName } }
}
"""),
    act("grupper-utan-agare", "Grupper utan ägare", "Grupper", "läs", "Group.Read.All",
        "Grupper som ingen äger – vanligt städobjekt.", [], r"""
Graph 'groups?$top=100&$select=id,displayName,mail,groupTypes,securityEnabled,mailEnabled,createdDateTime,onPremisesSyncEnabled&$expand=owners($select=id)' -All |
  Where-Object { -not @($_.owners).Count -and -not $_.onPremisesSyncEnabled } | Sort-Object displayName |
  ForEach-Object { [pscustomobject][ordered]@{ Namn = $_.displayName; Typ = GroupType $_; Epost = $_.mail; Skapad = D $_.createdDateTime; Id = $_.id } }
"""),
    act("grupper-dynamiska", "Dynamiska grupper", "Grupper", "läs", "Group.Read.All",
        "Grupper med dynamiskt medlemskap, deras regel och om regeln är på eller pausad.", [], r"""
Graph ('groups?$top=999&$select=id,displayName,mail,groupTypes,securityEnabled,mailEnabled,membershipRule,membershipRuleProcessingState&$count=true&$filter=' + (Enc "groupTypes/any(c:c eq 'DynamicMembership')")) -All -Eventual |
  Sort-Object displayName | ForEach-Object { [pscustomobject][ordered]@{ Namn = $_.displayName; Typ = GroupType $_
    Lage = $(if ($_.membershipRuleProcessingState -eq 'Paused') { 'pausad' } else { 'på' }); Regel = $_.membershipRule; Id = $_.id } }
"""),
    act("grupper-skapa", "Skapa grupp", "Grupper", "ändrar", "Group.ReadWrite.All User.Read.All",
        "Ny säkerhetsgrupp eller Microsoft 365-grupp, med ägare och medlemmar direkt.",
        ["Namn | Gruppens namn |", "Typ | Typ | Säkerhet | Säkerhet, Microsoft 365", "Beskrivning? | Beskrivning |",
         "Agare? | Ägare | | @user+", "Medlemmar? | Medlemmar | | @user+"], r"""
$nick = ($Namn.ToLower() -replace '[åä]', 'a' -replace 'ö', 'o' -replace '[^a-z0-9-]', '')
if (-not $nick) { $nick = 'grupp' + (Get-Random -Maximum 99999) }
$b = @{ displayName = $Namn; mailNickname = $nick }
if ($Typ -eq 'Microsoft 365') { $b.groupTypes = @('Unified'); $b.mailEnabled = $true; $b.securityEnabled = $false; $b.visibility = 'Private' }
else { $b.groupTypes = @(); $b.mailEnabled = $false; $b.securityEnabled = $true }
if ($Beskrivning) { $b.description = $Beskrivning }
$own = @(if ($Agare) { Users $Agare }); $mem = @(if ($Medlemmar) { Users $Medlemmar })
if ($own.Count) { $b['owners@odata.bind'] = @($own | ForEach-Object { "https://graph.microsoft.com/v1.0/users/$($_.id)" }) }
if ($mem.Count) { $b['members@odata.bind'] = @($mem | ForEach-Object { "https://graph.microsoft.com/v1.0/users/$($_.id)" }) }
$g = Graph 'groups' -Method POST -Body $b
[pscustomobject][ordered]@{ Namn = $g.displayName; Typ = GroupType $g; Epost = $g.mail; Agare = $own.Count; Medlemmar = $mem.Count; Id = $g.id }
"""),
    act("grupper-lagg-till", "Lägg till i grupp", "Grupper", "ändrar", "GroupMember.ReadWrite.All User.Read.All",
        "Lägger till användarna som medlemmar i en eller flera grupper. Redan medlem räknas som klart.",
        [U, "Grupp | Vilka grupper? | | @group+"], r"""
$us = @(Users $Anvandare)
foreach ($g in Groups $Grupp) {
  foreach ($u in $us) {
    try { Graph ("groups/$($g.id)/members/" + '$ref') -Method POST -Body (Ref $u.id) | Out-Null; Done $u.userPrincipalName "tillagd i $($g.displayName)" }
    catch { if ((Msg $_) -match 'already exist') { Done $u.userPrincipalName "redan med i $($g.displayName)" } else { Fail $u.userPrincipalName $_ } }
  }
}
"""),
    act("grupper-ta-bort", "Ta bort ur grupp", "Grupper", "ändrar", "GroupMember.ReadWrite.All User.Read.All",
        "Tar bort användarna ur en eller flera grupper.", [U, "Grupp | Vilka grupper? | | @group+"], r"""
$us = @(Users $Anvandare)
foreach ($g in Groups $Grupp) {
  foreach ($u in $us) {
    try { Graph ("groups/$($g.id)/members/$($u.id)/" + '$ref') -Method DELETE | Out-Null; Done $u.userPrincipalName "borttagen ur $($g.displayName)" }
    catch { if ((Msg $_) -match 'does not exist|ResourceNotFound') { Done $u.userPrincipalName "var inte med i $($g.displayName)" } else { Fail $u.userPrincipalName $_ } }
  }
}
"""),
    act("grupper-agare-lagg-till", "Lägg till ägare", "Grupper", "ändrar", "Group.ReadWrite.All User.Read.All",
        "Gör användarna till ägare av en eller flera grupper.", [U, "Grupp | Vilka grupper? | | @group+"], r"""
$us = @(Users $Anvandare)
foreach ($g in Groups $Grupp) {
  foreach ($u in $us) {
    try { Graph ("groups/$($g.id)/owners/" + '$ref') -Method POST -Body (Ref $u.id) | Out-Null; Done $u.userPrincipalName "ägare av $($g.displayName)" }
    catch { if ((Msg $_) -match 'already exist') { Done $u.userPrincipalName "redan ägare av $($g.displayName)" } else { Fail $u.userPrincipalName $_ } }
  }
}
"""),
    act("grupper-agare-ta-bort", "Ta bort ägare", "Grupper", "ändrar", "Group.ReadWrite.All User.Read.All",
        "Tar bort användarna som ägare. En grupp måste ha minst en ägare kvar om den har ägare.",
        [U, "Grupp | Vilka grupper? | | @group+"], r"""
$us = @(Users $Anvandare)
foreach ($g in Groups $Grupp) {
  foreach ($u in $us) {
    try { Graph ("groups/$($g.id)/owners/$($u.id)/" + '$ref') -Method DELETE | Out-Null; Done $u.userPrincipalName "inte längre ägare av $($g.displayName)" }
    catch { Fail $u.userPrincipalName $_ }
  }
}
"""),
    act("grupper-dynamisk-regel", "Sätt dynamisk regel", "Grupper", "ändrar", "Group.ReadWrite.All",
        "Gör gruppen dynamisk med en regel, t.ex. (user.department -eq \"Sälj\"). Tilldelade medlemmar ersätts av regeln.",
        ["Grupp | Vilken grupp? | | @group", "Regel | Regel | (user.department -eq \"Sälj\")",
         "Lage | Läge | På | På, Pausad"], r"""
foreach ($g in Groups $Grupp) {
  $types = @(@($g.groupTypes) | Where-Object { $_ -and $_ -ne 'DynamicMembership' }) + 'DynamicMembership'
  try {
    Graph "groups/$($g.id)" -Method PATCH -Body @{ groupTypes = $types; membershipRule = $Regel
      membershipRuleProcessingState = $(if ($Lage -eq 'Pausad') { 'Paused' } else { 'On' }) } | Out-Null
    Done $g.displayName "dynamisk ($Lage): $Regel"
  } catch { Fail $g.displayName $_ }
}
"""),
    act("grupper-radera", "Radera grupp", "Grupper", "ändrar", "Group.ReadWrite.All",
        "Raderar grupperna. Microsoft 365-grupper kan återställas i 30 dagar, säkerhetsgrupper inte.",
        ["Grupp | Vilka grupper? | | @group+"], r"""
foreach ($g in Groups $Grupp) {
  try { Graph "groups/$($g.id)" -Method DELETE | Out-Null; Done $g.displayName 'raderad' } catch { Fail $g.displayName $_ }
}
"""),
    # ─────────── Enheter ───────────
    act("enheter-lista", "Lista enheter", "Enheter", "läs", "Device.Read.All",
        "Alla enheter i Entra ID: OS, anslutning (Entra-ansluten, hybrid, registrerad), status och senaste inloggning.",
        ["Sok? | Filtrera på namn eller OS |"], r"""
$all = @(Graph 'devices?$top=999&$select=id,deviceId,displayName,operatingSystem,operatingSystemVersion,trustType,accountEnabled,approximateLastSignInDateTime,isCompliant,isManaged,registrationDateTime' -All)
if ($Sok) { $all = @($all | Where-Object { "$($_.displayName) $($_.operatingSystem)" -like "*$Sok*" }) }
$all | Sort-Object displayName | ForEach-Object { [pscustomobject][ordered]@{ Namn = $_.displayName; OS = "$($_.operatingSystem) $($_.operatingSystemVersion)"
  Anslutning = $(switch ($_.trustType) { 'AzureAd' { 'Entra-ansluten' } 'ServerAd' { 'hybrid' } 'Workplace' { 'registrerad' } default { $_.trustType } })
  Aktiv = JaNej $_.accountEnabled; Hanterad = JaNej $_.isManaged; Kompatibel = JaNej $_.isCompliant
  Senast = D $_.approximateLastSignInDateTime; Registrerad = D $_.registrationDateTime; DeviceId = $_.deviceId } }
"""),
    act("enheter-inaktiva", "Inaktiva enheter", "Enheter", "läs", "Device.Read.All",
        "Enheter som inte loggat in på X dagar – kandidater att inaktivera och sedan ta bort.",
        ["Dagar | Ingen aktivitet på hur många dagar? | 90"], r"""
Graph 'devices?$top=999&$select=deviceId,displayName,operatingSystem,trustType,accountEnabled,approximateLastSignInDateTime' -All |
  Where-Object { (Ago $_.approximateLastSignInDateTime) -ge [int]$Dagar } | Sort-Object approximateLastSignInDateTime |
  ForEach-Object { [pscustomobject][ordered]@{ Namn = $_.displayName; OS = $_.operatingSystem; Anslutning = $_.trustType; Aktiv = JaNej $_.accountEnabled
    Senast = D $_.approximateLastSignInDateTime; Dagar = Ago $_.approximateLastSignInDateTime; DeviceId = $_.deviceId } }
"""),
    act("enheter-inaktivera", "Inaktivera enhet", "Enheter", "ändrar", "Directory.AccessAsUser.All Device.Read.All",
        "Stänger av enheterna i Entra ID (de kan inte logga in), utan att ta bort dem.",
        ["Enhet | Vilka enheter? | | @device+"], r"""
foreach ($d in Devices $Enhet) {
  try { Graph "devices/$($d.id)" -Method PATCH -Body @{ accountEnabled = $false } | Out-Null; Done $d.displayName 'inaktiverad' } catch { Fail $d.displayName $_ }
}
"""),
    act("enheter-aktivera", "Aktivera enhet", "Enheter", "ändrar", "Directory.AccessAsUser.All Device.Read.All",
        "Slår på inaktiverade enheter igen.", ["Enhet | Vilka enheter? | | @device+"], r"""
foreach ($d in Devices $Enhet) {
  try { Graph "devices/$($d.id)" -Method PATCH -Body @{ accountEnabled = $true } | Out-Null; Done $d.displayName 'aktiverad' } catch { Fail $d.displayName $_ }
}
"""),
    act("enheter-ta-bort", "Ta bort enhet", "Enheter", "ändrar", "Directory.AccessAsUser.All Device.Read.All",
        "Raderar enheterna ur Entra ID. BitLocker-nycklar och LAPS-lösenord försvinner med dem – läs ut dem först.",
        ["Enhet | Vilka enheter? | | @device+"], r"""
foreach ($d in Devices $Enhet) {
  try { Graph "devices/$($d.id)" -Method DELETE | Out-Null; Done $d.displayName 'borttagen' } catch { Fail $d.displayName $_ }
}
"""),
    act("enheter-bitlocker", "BitLocker-nycklar", "Enheter", "läs", "BitlockerKey.Read.All Device.Read.All",
        "Återställningsnycklar för enheterna. Varje uppslag loggas i granskningsloggen.",
        ["Enhet | Vilka enheter? | | @device+"], r"""
foreach ($d in Devices $Enhet) {
  $keys = @(Graph ('informationProtection/bitlocker/recoveryKeys?$filter=' + (Enc "deviceId eq '$($d.deviceId)'")) -Headers $global:EuH)
  if (-not $keys.Count) { [pscustomobject][ordered]@{ Enhet = $d.displayName; Enhetstyp = ''; Nyckel = '(ingen nyckel sparad)'; NyckelId = ''; Skapad = '' } }
  foreach ($k in $keys) {
    $full = Graph ("informationProtection/bitlocker/recoveryKeys/$($k.id)" + '?$select=key') -Headers $global:EuH
    [pscustomobject][ordered]@{ Enhet = $d.displayName; Enhetstyp = $(switch ("$($k.volumeType)") { '1' { 'OS-disk' } '2' { 'fast disk' } '3' { 'flyttbar' } default { "$($k.volumeType)" } })
      Nyckel = $full.key; NyckelId = $k.id; Skapad = D $k.createdDateTime }
  }
}
"""),
    act("enheter-laps", "LAPS-lösenord", "Enheter", "läs", "DeviceLocalCredential.Read.All Device.Read.All",
        "Det lokala adminlösenordet (Windows LAPS i Entra ID) för enheterna. Varje uppslag loggas.",
        ["Enhet | Vilka enheter? | | @device+"], r"""
foreach ($d in Devices $Enhet) {
  try {
    $c = Graph ("directory/deviceLocalCredentials/$($d.deviceId)" + '?$select=credentials,deviceName,lastBackupDateTime,refreshDateTime') -Headers $global:EuH
    $cr = @($c.credentials) | Sort-Object backupDateTime -Descending | Select-Object -First 1
    [pscustomobject][ordered]@{ Enhet = $d.displayName; Konto = $cr.accountName
      Losenord = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($cr.passwordBase64)); Sparat = D $cr.backupDateTime; NastaByte = D $c.refreshDateTime }
  } catch { Fail $d.displayName $_ }
}
"""),
    # ─────────── Appar ───────────
    act("appar-lista", "Appregistreringar", "Appar", "läs", "Application.Read.All",
        "Alla appregistreringar med antal hemligheter och certifikat och när första går ut.", [], r"""
Graph 'applications?$top=999&$select=id,appId,displayName,createdDateTime,passwordCredentials,keyCredentials,signInAudience' -All | Sort-Object displayName |
  ForEach-Object {
    $ends = @(@($_.passwordCredentials) + @($_.keyCredentials) | Where-Object { $_.endDateTime } | ForEach-Object { [datetime]$_.endDateTime } | Sort-Object)
    [pscustomobject][ordered]@{ Namn = $_.displayName; AppId = $_.appId; Hemligheter = @($_.passwordCredentials).Count; Certifikat = @($_.keyCredentials).Count
      ForstaUtgang = $(if ($ends.Count) { D $ends[0] } else { '' }); Konton = $_.signInAudience; Skapad = D $_.createdDateTime }
  }
"""),
    act("appar-hemligheter-gar-ut", "Hemligheter och certifikat som går ut", "Appar", "läs", "Application.Read.All",
        "Klienthemligheter och certifikat (även SAML-signering på företagsappar) som går ut inom X dagar eller redan gått ut.",
        ["Dagar | Inom hur många dagar? | 30"], r"""
$lim = (Get-Date).AddDays([int]$Dagar)
$src = @(@{ t = 'appregistrering'; l = @(Graph 'applications?$top=999&$select=displayName,appId,passwordCredentials,keyCredentials' -All) },
         @{ t = 'företagsapp'; l = @(Graph 'servicePrincipals?$top=999&$select=displayName,appId,passwordCredentials,keyCredentials' -All) })
& { foreach ($s in $src) {
  foreach ($a in $s.l) {
    foreach ($k in @(@(@($a.passwordCredentials) | Where-Object { $_ } | ForEach-Object { $_ | Add-Member -NotePropertyName Sort -NotePropertyValue 'hemlighet' -PassThru -Force }) +
                     @(@($a.keyCredentials) | Where-Object { $_ } | ForEach-Object { $_ | Add-Member -NotePropertyName Sort -NotePropertyValue 'certifikat' -PassThru -Force }))) {
      if (-not $k.endDateTime -or [datetime]$k.endDateTime -gt $lim) { continue }
      $left = [int]([datetime]$k.endDateTime - (Get-Date)).TotalDays
      [pscustomobject][ordered]@{ Var = $s.t; App = $a.displayName; Sort = $k.Sort; Namn = $k.displayName; GarUt = D $k.endDateTime
        DagarKvar = $left; Status = $(if ($left -lt 0) { 'utgången' } else { 'går ut snart' }); AppId = $a.appId }
    }
  }
} } | Sort-Object DagarKvar
"""),
    act("appar-behorigheter", "Appens behörigheter", "Appar", "läs", "Application.Read.All Directory.Read.All",
        "API-behörigheter appen begär, delegerade och programbehörigheter, och om de har samtycke.",
        ["App | Vilken app? | | @app"], r"""
$a = @(Apps $App)[0]
if (-not $a) { return }
$sp = SpByAppId $a.appId
$dele = @{}; $roles = @{}
if ($sp) {
  foreach ($g in @(Graph ('oauth2PermissionGrants?$filter=' + (Enc "clientId eq '$($sp.id)'")) -All)) {
    foreach ($s in "$($g.scope)".Split(' ', [StringSplitOptions]::RemoveEmptyEntries)) { $dele["$($g.resourceId)|$s"] = $(if ($g.consentType -eq 'AllPrincipals') { 'ja (admin)' } else { 'ja (användare)' }) }
  }
  foreach ($r in @(Graph "servicePrincipals/$($sp.id)/appRoleAssignments" -All)) { $roles["$($r.resourceId)|$($r.appRoleId)"] = 'ja (admin)' }
}
foreach ($rra in @($a.requiredResourceAccess)) {
  $res = SpByAppId $rra.resourceAppId
  foreach ($ra in @($rra.resourceAccess)) {
    if ($ra.type -eq 'Scope') {
      $p = @($res.oauth2PermissionScopes | Where-Object id -eq $ra.id)[0]
      $ok = $dele["$($res.id)|$($p.value)"]
      [pscustomobject][ordered]@{ Resurs = $res.displayName; Behorighet = $p.value; Typ = 'delegerad'; Samtycke = $(if ($ok) { $ok } else { 'nej' }); Beskrivning = $p.adminConsentDisplayName }
    } else {
      $p = @($res.appRoles | Where-Object id -eq $ra.id)[0]
      $ok = $roles["$($res.id)|$($ra.id)"]
      [pscustomobject][ordered]@{ Resurs = $res.displayName; Behorighet = $p.value; Typ = 'program'; Samtycke = $(if ($ok) { $ok } else { 'nej' }); Beskrivning = $p.displayName }
    }
  }
}
"""),
    act("appar-hemlighet-ny", "Ny klienthemlighet", "Appar", "ändrar", "Application.ReadWrite.All",
        "Skapar en ny hemlighet på appen. Värdet visas bara den här gången – kopiera det direkt.",
        ["App | Vilken app? | | @app", "Namn | Namn på hemligheten | Euppa {datum}",
         "Manader | Giltig i månader | 6 | 3, 6, 12, 24"], r"""
foreach ($a in Apps $App) {
  try {
    $r = Graph "applications/$($a.id)/addPassword" -Method POST -Body @{ passwordCredential = @{ displayName = $Namn
      endDateTime = (Get-Date).ToUniversalTime().AddMonths([int]$Manader).ToString('yyyy-MM-ddTHH:mm:ssZ') } }
    Write-Warning 'Kopiera hemligheten nu – den går inte att visa igen.'
    [pscustomobject][ordered]@{ App = $a.displayName; AppId = $a.appId; Namn = $r.displayName; Hemlighet = $r.secretText; GarUt = D $r.endDateTime; KeyId = $r.keyId }
  } catch { Fail $a.displayName $_ }
}
"""),
    act("appar-hemligheter-rensa", "Ta bort utgångna hemligheter", "Appar", "ändrar", "Application.ReadWrite.All",
        "Rensar bort klienthemligheter som redan gått ut från apparna.", ["App | Vilka appar? | | @app+"], r"""
foreach ($a in Apps $App) {
  $old = @($a.passwordCredentials | Where-Object { $_.endDateTime -and [datetime]$_.endDateTime -lt (Get-Date) })
  if (-not $old.Count) { Done $a.displayName 'inga utgångna hemligheter'; continue }
  foreach ($k in $old) {
    try { Graph "applications/$($a.id)/removePassword" -Method POST -Body @{ keyId = $k.keyId } | Out-Null; Done $a.displayName "tog bort '$($k.displayName)' (gick ut $(D $k.endDateTime))" }
    catch { Fail $a.displayName $_ }
  }
}
"""),
    act("foretagsappar-lista", "Företagsappar", "Appar", "läs", "Application.Read.All",
        "Företagsappar (enterprise applications): på/av, om tilldelning krävs, och om appen är egen eller extern.", [], r"""
$tid = (Get-MgContext).TenantId
Graph ('servicePrincipals?$top=999&$count=true&$select=id,appId,displayName,accountEnabled,appRoleAssignmentRequired,publisherName,appOwnerOrganizationId,servicePrincipalType&$filter=' + (Enc "tags/any(t:t eq 'WindowsAzureActiveDirectoryIntegratedApp')")) -All -Eventual |
  Sort-Object displayName | ForEach-Object { [pscustomobject][ordered]@{ Namn = $_.displayName; Aktiv = JaNej $_.accountEnabled
    KraverTilldelning = JaNej $_.appRoleAssignmentRequired; Utgivare = $_.publisherName; Agare = $(if ($_.appOwnerOrganizationId -eq $tid) { 'egen' } else { 'extern' }); AppId = $_.appId } }
"""),
    act("foretagsappar-samtycken", "Samtycken (delegerade)", "Appar", "läs", "Directory.Read.All",
        "Vilka appar som fått delegerade behörigheter, av admin för alla eller av enskilda användare.",
        ["Fapp? | Bara för en app (tomt = alla) | | @sp"], r"""
$grants = @(Graph 'oauth2PermissionGrants?$top=999' -All)
$name = @{}; $who = @{}
function nm($id) { if (-not $name[$id]) { try { $name[$id] = (Graph ("servicePrincipals/$id" + '?$select=displayName')).displayName } catch { $name[$id] = $id } }; $name[$id] }
if ($Fapp) { $sp = @(Sps $Fapp)[0]; $grants = @($grants | Where-Object clientId -eq $sp.id) }
foreach ($g in $grants) {
  $u = ''
  if ($g.principalId) { if (-not $who[$g.principalId]) { try { $who[$g.principalId] = (Graph ("users/$($g.principalId)" + '?$select=userPrincipalName')).userPrincipalName } catch { $who[$g.principalId] = $g.principalId } }; $u = $who[$g.principalId] }
  [pscustomobject][ordered]@{ App = nm $g.clientId; Resurs = nm $g.resourceId; Samtycke = $(if ($g.consentType -eq 'AllPrincipals') { 'admin – alla' } else { 'användare' }); Anvandare = $u; Behorigheter = "$($g.scope)".Trim() }
}
"""),
    act("foretagsappar-anvandare", "Vem har åtkomst till appen", "Appar", "läs", "Application.Read.All Directory.Read.All",
        "Användare, grupper och appar som är tilldelade företagsappen, med roll.", ["Fapp | Vilken företagsapp? | | @sp"], r"""
foreach ($sp in Sps $Fapp) {
  $rn = @{ '00000000-0000-0000-0000-000000000000' = 'standardåtkomst' }
  foreach ($r in @($sp.appRoles)) { $rn[$r.id] = $r.displayName }
  Graph "servicePrincipals/$($sp.id)/appRoleAssignedTo" -All | ForEach-Object { [pscustomobject][ordered]@{ App = $sp.displayName
    Namn = $_.principalDisplayName; Typ = $_.principalType; Roll = $rn[$_.appRoleId]; Sedan = D $_.createdDateTime } }
}
"""),
    act("foretagsappar-samtycke", "Bevilja adminsamtycke", "Appar", "ändrar",
        "Application.ReadWrite.All DelegatedPermissionGrant.ReadWrite.All AppRoleAssignment.ReadWrite.All",
        "Ger samtycke för hela organisationen till alla behörigheter appen begär (som knappen i portalen).",
        ["App | Vilken app? | | @app"], r"""
foreach ($a in Apps $App) {
  $sp = SpByAppId $a.appId
  if (-not $sp) { $sp = Graph 'servicePrincipals' -Method POST -Body @{ appId = $a.appId }; Done $a.displayName 'skapade företagsappen' }
  foreach ($rra in @($a.requiredResourceAccess)) {
    $res = SpByAppId $rra.resourceAppId
    if (-not $res) { Fail $a.displayName "resursen $($rra.resourceAppId) finns inte i tenanten"; continue }
    $want = @($rra.resourceAccess | Where-Object type -eq 'Scope' | ForEach-Object { $id = $_.id; @($res.oauth2PermissionScopes | Where-Object id -eq $id)[0].value } | Where-Object { $_ })
    if ($want.Count) {
      $g = @(Graph ('oauth2PermissionGrants?$filter=' + (Enc "clientId eq '$($sp.id)' and consentType eq 'AllPrincipals' and resourceId eq '$($res.id)'")))[0]
      try {
        if ($g) { $all = @(("$($g.scope)".Split(' ', [StringSplitOptions]::RemoveEmptyEntries) + $want) | Select-Object -Unique)
                  Graph "oauth2PermissionGrants/$($g.id)" -Method PATCH -Body @{ scope = ($all -join ' ') } | Out-Null }
        else { Graph 'oauth2PermissionGrants' -Method POST -Body @{ clientId = $sp.id; consentType = 'AllPrincipals'; resourceId = $res.id; scope = ($want -join ' ') } | Out-Null }
        Done $a.displayName "$($res.displayName) (delegerad): $($want -join ', ')"
      } catch { Fail $a.displayName $_ }
    }
    $have = @(Graph "servicePrincipals/$($sp.id)/appRoleAssignments" -All | ForEach-Object { "$($_.resourceId)|$($_.appRoleId)" })
    foreach ($ra in @($rra.resourceAccess | Where-Object type -eq 'Role')) {
      $rv = @($res.appRoles | Where-Object id -eq $ra.id)[0].value
      if ($have -contains "$($res.id)|$($ra.id)") { Done $a.displayName "$($res.displayName) (program): $rv – fanns redan"; continue }
      try { Graph "servicePrincipals/$($res.id)/appRoleAssignedTo" -Method POST -Body @{ principalId = $sp.id; resourceId = $res.id; appRoleId = $ra.id } | Out-Null
            Done $a.displayName "$($res.displayName) (program): $rv" }
      catch { Fail $a.displayName $_ }
    }
  }
}
"""),
    act("foretagsappar-aterkalla", "Återkalla samtycke", "Appar", "ändrar",
        "Application.Read.All DelegatedPermissionGrant.ReadWrite.All AppRoleAssignment.ReadWrite.All",
        "Tar bort alla samtycken (delegerade och/eller programbehörigheter) som företagsappen har fått.",
        ["Fapp | Vilken företagsapp? | | @sp", "Vad | Vilka? | Båda | Båda, Delegerade, Program"], r"""
foreach ($sp in Sps $Fapp) {
  if ($Vad -ne 'Program') {
    foreach ($g in @(Graph ('oauth2PermissionGrants?$filter=' + (Enc "clientId eq '$($sp.id)'")) -All)) {
      try { Graph "oauth2PermissionGrants/$($g.id)" -Method DELETE | Out-Null; Done $sp.displayName "återkallade delegerat: $("$($g.scope)".Trim())" } catch { Fail $sp.displayName $_ }
    }
  }
  if ($Vad -ne 'Delegerade') {
    foreach ($r in @(Graph "servicePrincipals/$($sp.id)/appRoleAssignments" -All)) {
      try { Graph "servicePrincipals/$($sp.id)/appRoleAssignments/$($r.id)" -Method DELETE | Out-Null; Done $sp.displayName "återkallade program: $($r.resourceDisplayName)" } catch { Fail $sp.displayName $_ }
    }
  }
}
"""),
    act("foretagsappar-av-pa", "Stäng av eller slå på företagsapp", "Appar", "ändrar", "Application.ReadWrite.All",
        "Av = ingen kan logga in i appen (bra vid misstänkt app). På = tillbaka igen.",
        ["Fapp | Vilka företagsappar? | | @sp+", "Lage | Läge | Av | Av, På"], r"""
foreach ($sp in Sps $Fapp) {
  try { Graph "servicePrincipals/$($sp.id)" -Method PATCH -Body @{ accountEnabled = ($Lage -eq 'På') } | Out-Null; Done $sp.displayName $(if ($Lage -eq 'På') { 'påslagen' } else { 'avstängd' }) }
  catch { Fail $sp.displayName $_ }
}
"""),
    # ─────────── Roller och PIM ───────────
    act("roller-lista", "Vem har adminroller", "Roller och PIM", "läs", "RoleManagement.Read.Directory Directory.Read.All",
        "Alla aktiva rolltilldelningar: roll, vem och omfattning.", [], r"""
$defs = @{}; Graph 'roleManagement/directory/roleDefinitions?$select=id,displayName' -All | ForEach-Object { $defs[$_.id] = $_.displayName }
Graph 'roleManagement/directory/roleAssignments?$expand=principal' -All | ForEach-Object { [pscustomobject][ordered]@{
  Roll = $defs[$_.roleDefinitionId]; Namn = $_.principal.displayName; Konto = $(if ($_.principal.userPrincipalName) { $_.principal.userPrincipalName } else { $_.principal.appId })
  Typ = UserType $_.principal; Omfattning = $(if ($_.directoryScopeId -eq '/') { 'hela tenanten' } else { $_.directoryScopeId }) } } | Sort-Object Roll, Namn
"""),
    act("roller-anvandarens", "Användarens roller", "Roller och PIM", "läs", "RoleManagement.Read.Directory User.Read.All",
        "Aktiva och berättigade (PIM) roller för användarna.", [U], r"""
$defs = @{}; Graph 'roleManagement/directory/roleDefinitions?$select=id,displayName' -All | ForEach-Object { $defs[$_.id] = $_.displayName }
foreach ($u in Users $Anvandare) {
  Graph ('roleManagement/directory/roleAssignments?$filter=' + (Enc "principalId eq '$($u.id)'")) -All | ForEach-Object {
    [pscustomobject][ordered]@{ Anvandare = $u.userPrincipalName; Roll = $defs[$_.roleDefinitionId]; Status = 'aktiv'; Till = '' } }
  try { Graph ('roleManagement/directory/roleEligibilitySchedules?$filter=' + (Enc "principalId eq '$($u.id)'")) -All | ForEach-Object {
    [pscustomobject][ordered]@{ Anvandare = $u.userPrincipalName; Roll = $defs[$_.roleDefinitionId]; Status = 'berättigad (PIM)'; Till = $(if ($_.scheduleInfo.expiration.endDateTime) { D $_.scheduleInfo.expiration.endDateTime } else { 'tills vidare' }) } } }
  catch { Write-Warning "PIM gick inte att läsa (kräver Entra ID P2): $(Msg $_)" }
}
"""),
    act("pim-berattigade", "Berättigade roller (PIM)", "Roller och PIM", "läs", "RoleManagement.Read.Directory",
        "Alla som kan aktivera en roll i PIM, och hur länge de får. Kräver Entra ID P2.", [], r"""
Graph 'roleManagement/directory/roleEligibilitySchedules?$expand=principal,roleDefinition' -All | ForEach-Object { [pscustomobject][ordered]@{
  Roll = $_.roleDefinition.displayName; Namn = $_.principal.displayName; Konto = $_.principal.userPrincipalName; Typ = UserType $_.principal
  Fran = D $_.scheduleInfo.startDateTime; Till = $(if ($_.scheduleInfo.expiration.endDateTime) { D $_.scheduleInfo.expiration.endDateTime } else { 'tills vidare' }) } } | Sort-Object Roll, Namn
"""),
    act("pim-aktiva", "Aktiva PIM-tilldelningar", "Roller och PIM", "läs", "RoleManagement.Read.Directory",
        "Roller som är aktiva just nu – permanenta och aktiverade via PIM – med sluttid. Kräver Entra ID P2.", [], r"""
Graph 'roleManagement/directory/roleAssignmentScheduleInstances?$expand=principal,roleDefinition' -All | ForEach-Object { [pscustomobject][ordered]@{
  Roll = $_.roleDefinition.displayName; Namn = $_.principal.displayName; Konto = $_.principal.userPrincipalName
  Hur = $(if ($_.assignmentType -eq 'Activated') { 'aktiverad i PIM' } else { 'tilldelad' }); Fran = D $_.startDateTime; Till = $(if ($_.endDateTime) { D $_.endDateTime } else { 'permanent' }) } } | Sort-Object Roll, Namn
"""),
    act("roller-tilldela", "Tilldela roll", "Roller och PIM", "ändrar", "RoleManagement.ReadWrite.Directory User.Read.All",
        "Ger användarna en adminroll – helst som berättigad i PIM (aktiveras vid behov) i stället för permanent.",
        [U, "Roll | Vilken roll? | | @role", "Typ | Hur? | Berättigad (PIM) | Berättigad (PIM), Aktiv",
         "Dagar | Hur många dagar? (0 = tills vidare) | 365", "Motivering | Motivering | Tilldelad med Euppa"], r"""
$r = @(Roles $Roll)[0]
if (-not $r) { return }
$exp = $(if ([int]$Dagar -gt 0) { @{ type = 'afterDuration'; duration = "P$([int]$Dagar)D" } } else { @{ type = 'noExpiration' } })
$start = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
foreach ($u in Users $Anvandare) {
  try {
    if ($Typ -like 'Berättigad*') {
      Graph 'roleManagement/directory/roleEligibilityScheduleRequests' -Method POST -Body @{ action = 'adminAssign'; justification = $Motivering
        roleDefinitionId = $r.id; directoryScopeId = '/'; principalId = $u.id; scheduleInfo = @{ startDateTime = $start; expiration = $exp } } | Out-Null
      Done $u.userPrincipalName "berättigad till $($r.displayName)"
    } elseif ([int]$Dagar -gt 0) {
      Graph 'roleManagement/directory/roleAssignmentScheduleRequests' -Method POST -Body @{ action = 'adminAssign'; justification = $Motivering
        roleDefinitionId = $r.id; directoryScopeId = '/'; principalId = $u.id; scheduleInfo = @{ startDateTime = $start; expiration = $exp } } | Out-Null
      Done $u.userPrincipalName "aktiv $($r.displayName) i $Dagar dagar"
    } else {
      Graph 'roleManagement/directory/roleAssignments' -Method POST -Body @{ principalId = $u.id; roleDefinitionId = $r.id; directoryScopeId = '/' } | Out-Null
      Done $u.userPrincipalName "permanent $($r.displayName)"
    }
  } catch { Fail $u.userPrincipalName $_ }
}
"""),
    act("roller-ta-bort", "Ta bort roll", "Roller och PIM", "ändrar", "RoleManagement.ReadWrite.Directory User.Read.All",
        "Tar bort rollen från användarna – både aktiva och berättigade (PIM) tilldelningar.",
        [U, "Roll | Vilken roll? | | @role", "Motivering | Motivering | Borttagen med Euppa"], r"""
$r = @(Roles $Roll)[0]
if (-not $r) { return }
foreach ($u in Users $Anvandare) {
  $n = 0
  $f = Enc "principalId eq '$($u.id)' and roleDefinitionId eq '$($r.id)'"
  foreach ($x in @(Graph "roleManagement/directory/roleAssignments?`$filter=$f")) {
    try { Graph "roleManagement/directory/roleAssignments/$($x.id)" -Method DELETE | Out-Null; $n++ }
    catch {
      try { Graph 'roleManagement/directory/roleAssignmentScheduleRequests' -Method POST -Body @{ action = 'adminRemove'; principalId = $u.id; roleDefinitionId = $r.id; directoryScopeId = '/'; justification = $Motivering } | Out-Null; $n++ }
      catch { Fail $u.userPrincipalName $_ }
    }
  }
  try {
    foreach ($x in @(Graph "roleManagement/directory/roleEligibilitySchedules?`$filter=$f")) {
      Graph 'roleManagement/directory/roleEligibilityScheduleRequests' -Method POST -Body @{ action = 'adminRemove'; principalId = $u.id; roleDefinitionId = $r.id; directoryScopeId = $x.directoryScopeId; justification = $Motivering } | Out-Null; $n++
    }
  } catch { if ((Msg $_) -notmatch 'P2|license|Premium') { Fail $u.userPrincipalName $_ } }
  Done $u.userPrincipalName $(if ($n) { "$($r.displayName) borttagen ($n tilldelningar)" } else { "hade inte $($r.displayName)" })
}
"""),
    act("pim-aktivera", "Aktivera min roll (PIM)", "Roller och PIM", "ändrar", "RoleAssignmentSchedule.ReadWrite.Directory RoleManagement.Read.Directory",
        "Aktiverar en roll du är berättigad till, i några timmar. Kräver ofta MFA och motivering.",
        ["Roll | Vilken roll? | | @role", "Timmar | Hur många timmar? | 4 | 1, 2, 4, 8",
         "Motivering | Varför? |"], r"""
$r = @(Roles $Roll)[0]
if (-not $r) { return }
$me = Me
$x = Graph 'roleManagement/directory/roleAssignmentScheduleRequests' -Method POST -Body @{ action = 'selfActivate'; principalId = $me.id
  roleDefinitionId = $r.id; directoryScopeId = '/'; justification = $Motivering
  scheduleInfo = @{ startDateTime = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'); expiration = @{ type = 'afterDuration'; duration = "PT$([int]$Timmar)H" } } }
[pscustomobject][ordered]@{ Objekt = $me.userPrincipalName; Resultat = "$($r.displayName): $($x.status)"; Till = D (Get-Date).AddHours([int]$Timmar) }
"""),
    # ─────────── Villkorlig åtkomst ───────────
    act("ca-lista", "Policyer för villkorlig åtkomst", "Villkorlig åtkomst", "läs", "Policy.Read.All",
        "Alla policyer: på, av eller report-only, vem och vilka appar de gäller och vad de kräver.", [], r"""
function sum($inc, $exc) {
  $i = @($inc) | Where-Object { $_ }
  $t = $(if ($i -contains 'All') { 'alla' } elseif ($i -contains 'GuestsOrExternalUsers') { 'gäster' } elseif ($i -contains 'None' -or -not $i.Count) { '–' } else { "$($i.Count) st" })
  $t
}
Graph 'identity/conditionalAccess/policies' | Sort-Object displayName | ForEach-Object {
  $c = $_.conditions; $g = $_.grantControls
  $krav = @($g.builtInControls) + @($(if ($g.authenticationStrength) { "styrka: $($g.authenticationStrength.displayName)" }))
  [pscustomobject][ordered]@{ Namn = $_.displayName; Lage = $(switch ($_.state) { 'enabled' { 'på' } 'disabled' { 'av' } default { 'report-only' } })
    Anvandare = sum $c.users.includeUsers; Grupper = @($c.users.includeGroups).Count; Roller = @($c.users.includeRoles).Count
    Undantag = @($c.users.excludeUsers).Count + @($c.users.excludeGroups).Count + @($c.users.excludeRoles).Count
    Appar = sum $c.applications.includeApplications; Krav = (($krav | Where-Object { $_ }) -join " $($g.operator) "); Andrad = D $_.modifiedDateTime; Id = $_.id }
}
"""),
    act("ca-detalj", "Policy i klartext", "Villkorlig åtkomst", "läs", "Policy.Read.All Directory.Read.All Application.Read.All",
        "En policy uppdelad i vem, appar, villkor och krav – med namn i stället för id:n.",
        ["Policy | Vilken policy? | | @ca"], r"""
function names($ids) {
  $ids = @($ids) | Where-Object { $_ }
  $plain = @($ids | Where-Object { -not (IsGuid $_) })
  $g = @($ids | Where-Object { IsGuid $_ })
  $n = @()
  if ($g.Count) { try { $n = @(Graph 'directoryObjects/getByIds' -Method POST -Body @{ ids = $g } | ForEach-Object { $_.value } | ForEach-Object { if ($_.userPrincipalName) { $_.userPrincipalName } else { $_.displayName } }) } catch { $n = $g } }
  (($plain + $n) | Where-Object { $_ }) -join ', '
}
function rolenames($ids) { $ids = @($ids) | Where-Object { $_ }; if (-not $ids.Count) { return '' }
  $d = @{}; Graph 'roleManagement/directory/roleDefinitions?$select=templateId,displayName' -All | ForEach-Object { $d[$_.templateId] = $_.displayName }
  ($ids | ForEach-Object { if ($d[$_]) { $d[$_] } else { $_ } }) -join ', ' }
function appnames($ids) { (@($ids) | Where-Object { $_ } | ForEach-Object { if (IsGuid $_) { $s = SpByAppId $_; if ($s) { $s.displayName } else { $_ } } else { $_ } }) -join ', ' }
function locnames($ids) { $ids = @($ids) | Where-Object { $_ }; if (-not $ids.Count) { return '' }
  $l = @{}; Graph 'identity/conditionalAccess/namedLocations' -All | ForEach-Object { $l[$_.id] = $_.displayName }
  ($ids | ForEach-Object { if ($l[$_]) { $l[$_] } else { $_ } }) -join ', ' }
foreach ($p in CaPolicies $Policy) {
  $c = $p.conditions; $g = $p.grantControls; $s = $p.sessionControls
  $rows = [ordered]@{
    'Läge' = $(switch ($p.state) { 'enabled' { 'på' } 'disabled' { 'av' } default { 'report-only' } })
    'Användare – gäller' = names $c.users.includeUsers; 'Användare – undantag' = names $c.users.excludeUsers
    'Grupper – gäller' = names $c.users.includeGroups; 'Grupper – undantag' = names $c.users.excludeGroups
    'Roller – gäller' = rolenames $c.users.includeRoles; 'Roller – undantag' = rolenames $c.users.excludeRoles
    'Appar – gäller' = appnames $c.applications.includeApplications; 'Appar – undantag' = appnames $c.applications.excludeApplications
    'Användaråtgärder' = (@($c.applications.includeUserActions) -join ', ')
    'Plattformar' = ((@($c.platforms.includePlatforms) -join ', ') + $(if (@($c.platforms.excludePlatforms).Count) { ' utom ' + (@($c.platforms.excludePlatforms) -join ', ') }))
    'Platser – gäller' = locnames $c.locations.includeLocations; 'Platser – undantag' = locnames $c.locations.excludeLocations
    'Klientappar' = (@($c.clientAppTypes) -join ', '); 'Inloggningsrisk' = (@($c.signInRiskLevels) -join ', '); 'Användarrisk' = (@($c.userRiskLevels) -join ', ')
    'Krav' = ((@($g.builtInControls) + @($(if ($g.authenticationStrength) { "autentiseringsstyrka: $($g.authenticationStrength.displayName)" })) | Where-Object { $_ }) -join " $($g.operator) ")
    'Session' = ((@($(if ($s.signInFrequency.isEnabled) { "logga in var $($s.signInFrequency.value) $($s.signInFrequency.type)" }),
                   $(if ($s.persistentBrowser.isEnabled) { "beständig webbläsare: $($s.persistentBrowser.mode)" }),
                   $(if ($s.applicationEnforcedRestrictions.isEnabled) { 'appstyrda begränsningar' })) | Where-Object { $_ }) -join ', ')
    'Ändrad' = D $p.modifiedDateTime }
  foreach ($k in $rows.Keys) { if ("$($rows[$k])".Trim()) { [pscustomobject][ordered]@{ Policy = $p.displayName; Del = $k; Varde = $rows[$k] } } }
}
"""),
    act("ca-lage", "Byt läge på policy", "Villkorlig åtkomst", "ändrar",
        "Policy.ReadWrite.ConditionalAccess Policy.Read.All Application.Read.All",
        "Sätt policyer i report-only (loggar vad som skulle hänt), på eller av.",
        ["Policy | Vilka policyer? | | @ca+", "Lage | Nytt läge | Report-only | Report-only, På, Av"], r"""
$st = @{ 'Report-only' = 'enabledForReportingButNotEnforced'; 'På' = 'enabled'; 'Av' = 'disabled' }[$Lage]
foreach ($p in CaPolicies $Policy) {
  try { Graph "identity/conditionalAccess/policies/$($p.id)" -Method PATCH -Body @{ state = $st } | Out-Null; Done $p.displayName "läge: $Lage" } catch { Fail $p.displayName $_ }
}
"""),
    act("ca-undanta", "Undanta användare från policy", "Villkorlig åtkomst", "ändrar",
        "Policy.ReadWrite.ConditionalAccess Policy.Read.All Application.Read.All User.Read.All",
        "Lägger till användare som undantag i policyer, t.ex. break-glass-konton.",
        [U, "Policy | Vilka policyer? | | @ca+"], r"""
$ids = @(Users $Anvandare | ForEach-Object { $_.id })
foreach ($p in CaPolicies $Policy) {
  $c = $p.conditions
  $c.users.excludeUsers = @(@($c.users.excludeUsers) + $ids | Where-Object { $_ } | Select-Object -Unique)
  try { Graph "identity/conditionalAccess/policies/$($p.id)" -Method PATCH -Body @{ conditions = $c } | Out-Null; Done $p.displayName "$($ids.Count) undantagna" }
  catch { Fail $p.displayName $_ }
}
"""),
    act("ca-platser", "Namngivna platser", "Villkorlig åtkomst", "läs", "Policy.Read.All",
        "IP-intervall och länder som används i policyerna, och vilka som är betrodda.", [], r"""
Graph 'identity/conditionalAccess/namedLocations' -All | Sort-Object displayName | ForEach-Object {
  $ip = $_.'@odata.type' -like '*ipNamedLocation'
  [pscustomobject][ordered]@{ Namn = $_.displayName; Typ = $(if ($ip) { 'IP-intervall' } else { 'länder' }); Betrodd = JaNej $_.isTrusted
    Innehall = $(if ($ip) { (@($_.ipRanges.cidrAddress) -join ', ') } else { (@($_.countriesAndRegions) -join ', ') }); Andrad = D $_.modifiedDateTime; Id = $_.id } }
"""),
    act("ca-plats-ny", "Ny namngiven plats", "Villkorlig åtkomst", "ändrar", "Policy.ReadWrite.ConditionalAccess Policy.Read.All",
        "Skapar en plats av IP-intervall (CIDR, t.ex. 203.0.113.0/24) eller länder (SE, NO …).",
        ["Namn | Platsens namn |", "Typ | Typ | IP-intervall | IP-intervall, Länder",
         "Varden | IP-intervall eller landskoder, kommaseparerade |", "Betrodd | Betrodd plats? (bara IP) | Nej | Nej, Ja"], r"""
if ($Typ -eq 'Länder') {
  $b = @{ '@odata.type' = '#microsoft.graph.countryNamedLocation'; displayName = $Namn; countriesAndRegions = @(Items $Varden | ForEach-Object { $_.ToUpper() }); includeUnknownCountriesAndRegions = $false }
} else {
  $b = @{ '@odata.type' = '#microsoft.graph.ipNamedLocation'; displayName = $Namn; isTrusted = ($Betrodd -eq 'Ja')
    ipRanges = @(Items $Varden | ForEach-Object { if ($_ -match ':') { @{ '@odata.type' = '#microsoft.graph.iPv6CidrRange'; cidrAddress = $_ } } else { @{ '@odata.type' = '#microsoft.graph.iPv4CidrRange'; cidrAddress = $(if ($_ -match '/') { $_ } else { "$_/32" }) } } }) }
}
$r = Graph 'identity/conditionalAccess/namedLocations' -Method POST -Body $b
[pscustomobject][ordered]@{ Namn = $r.displayName; Typ = $Typ; Innehall = $Varden; Id = $r.id }
"""),
    # ─────────── Loggar ───────────
    act("loggar-inloggningar", "Inloggningar", "Loggar", "läs", "AuditLog.Read.All Directory.Read.All",
        "Inloggningsloggen: vem, app, IP, plats, resultat och villkorlig åtkomst. Kräver Entra ID P1 (max 999 rader).",
        ["Anvandare? | Bara för användare (tomt = alla) | | @user", "Timmar | Hur långt bakåt (timmar)? | 24",
         "BaraFel | Bara misslyckade? | Nej | Nej, Ja"], r"""
$f = "createdDateTime ge $(Since $Timmar)"
if ($Anvandare) { $f += " and userPrincipalName eq '$(Q $Anvandare)'" }
if ($BaraFel -eq 'Ja') { $f += ' and status/errorCode ne 0' }
Graph ('auditLogs/signIns?$top=999&$filter=' + (Enc $f)) | ForEach-Object { [pscustomobject][ordered]@{
  Tid = D $_.createdDateTime; Anvandare = $_.userPrincipalName; App = $_.appDisplayName; IP = $_.ipAddress
  Plats = ((@($_.location.city, $_.location.countryOrRegion) | Where-Object { $_ }) -join ', ')
  Resultat = $(if ($_.status.errorCode -eq 0) { 'OK' } else { "fel $($_.status.errorCode): $($_.status.failureReason)" })
  VillkorligAtkomst = $(switch ($_.conditionalAccessStatus) { 'success' { 'godkänd' } 'failure' { 'nekad' } 'notApplied' { 'ej tillämpad' } default { $_ } })
  Klient = $_.clientAppUsed; Enhet = "$($_.deviceDetail.operatingSystem) $($_.deviceDetail.browser)".Trim() } }
"""),
    act("loggar-misslyckade", "Misslyckade inloggningar per användare", "Loggar", "läs", "AuditLog.Read.All Directory.Read.All",
        "Sammanställning: vem som misslyckats, hur många gånger, med vilket fel och från hur många IP.",
        ["Timmar | Hur långt bakåt (timmar)? | 24"], r"""
Graph ('auditLogs/signIns?$top=999&$filter=' + (Enc "createdDateTime ge $(Since $Timmar) and status/errorCode ne 0")) |
  Group-Object userPrincipalName | ForEach-Object { $g = $_.Group
    [pscustomobject][ordered]@{ Anvandare = $_.Name; Antal = $_.Count; VanligasteFel = ($g | Group-Object { "$($_.status.errorCode) $($_.status.failureReason)" } | Sort-Object Count -Descending | Select-Object -First 1).Name
      IP = @($g.ipAddress | Select-Object -Unique).Count; Lander = (@($g.location.countryOrRegion | Where-Object { $_ } | Select-Object -Unique) -join ', '); Senast = D ($g | Sort-Object createdDateTime -Descending | Select-Object -First 1).createdDateTime } } |
  Sort-Object Antal -Descending
"""),
    act("loggar-granskning", "Granskningslogg", "Loggar", "läs", "AuditLog.Read.All Directory.Read.All",
        "Vem ändrade vad i katalogen: användare, grupper, roller, appar, policyer (max 999 rader).",
        ["Timmar | Hur långt bakåt (timmar)? | 24", "Sok? | Filtrera på aktivitet eller mål |",
         "Anvandare? | Bara ändringar gjorda av (UPN) | | @user"], r"""
$f = "activityDateTime ge $(Since $Timmar)"
if ($Anvandare) { $f += " and initiatedBy/user/userPrincipalName eq '$(Q $Anvandare)'" }
$rows = Graph ('auditLogs/directoryAudits?$top=999&$filter=' + (Enc $f)) | ForEach-Object { [pscustomobject][ordered]@{
  Tid = D $_.activityDateTime; Aktivitet = $_.activityDisplayName; Kategori = $_.category
  UtfordAv = $(if ($_.initiatedBy.user.userPrincipalName) { $_.initiatedBy.user.userPrincipalName } else { $_.initiatedBy.app.displayName })
  Mal = ((@($_.targetResources | ForEach-Object { if ($_.userPrincipalName) { $_.userPrincipalName } else { $_.displayName } }) | Where-Object { $_ }) -join ', ')
  Resultat = $(if ($_.result -eq 'success') { 'OK' } else { "$($_.result) $($_.resultReason)".Trim() }) } }
if ($Sok) { $rows = $rows | Where-Object { "$($_.Aktivitet) $($_.Mal) $($_.Kategori)" -like "*$Sok*" } }
$rows
"""),
    # ─────────── Säkerhet ───────────
    act("risk-anvandare", "Riskanvändare", "Säkerhet", "läs", "IdentityRiskyUser.Read.All",
        "Användare som Entra ID Protection flaggat som riskabla. Kräver Entra ID P2.",
        ["Alla | Visa även åtgärdade? | Nej | Nej, Ja"], r"""
$u = 'identityProtection/riskyUsers?$top=500'
if ($Alla -ne 'Ja') { $u += '&$filter=' + (Enc "riskState eq 'atRisk' or riskState eq 'confirmedCompromised'") }
Graph $u -All | Sort-Object riskLastUpdatedDateTime -Descending | ForEach-Object { [pscustomobject][ordered]@{
  Namn = $_.userDisplayName; UPN = $_.userPrincipalName; Risk = $_.riskLevel; Status = $_.riskState; Detalj = $_.riskDetail; Uppdaterad = D $_.riskLastUpdatedDateTime } }
"""),
    act("risk-handelser", "Riskhändelser", "Säkerhet", "läs", "IdentityRiskEvent.Read.All",
        "Upptäckta risker: okänd plats, läckta lösenord, omöjlig resa m.m. Kräver Entra ID P2.",
        ["Dagar | Hur långt bakåt (dagar)? | 7"], r"""
Graph ('identityProtection/riskDetections?$top=500&$filter=' + (Enc "detectedDateTime ge $(Since (24 * [int]$Dagar))")) -All | Sort-Object detectedDateTime -Descending |
  ForEach-Object { [pscustomobject][ordered]@{ Tid = D $_.detectedDateTime; Anvandare = $_.userPrincipalName; Typ = $_.riskEventType; Niva = $_.riskLevel
    Status = $_.riskState; IP = $_.ipAddress; Plats = ((@($_.location.city, $_.location.countryOrRegion) | Where-Object { $_ }) -join ', ') } }
"""),
    act("risk-avfarda", "Avfärda användarrisk", "Säkerhet", "ändrar", "IdentityRiskyUser.ReadWrite.All User.Read.All",
        "Markerar risken som hanterad (t.ex. efter lösenordsbyte).", [U], r"""
$us = @(Users $Anvandare)
Graph 'identityProtection/riskyUsers/dismiss' -Method POST -Body @{ userIds = @($us.id) } | Out-Null
$us | ForEach-Object { Done $_.userPrincipalName 'risk avfärdad' }
"""),
    act("risk-komprometterad", "Bekräfta komprometterad", "Säkerhet", "ändrar", "IdentityRiskyUser.ReadWrite.All User.Read.All",
        "Sätter användarrisken till hög (bekräftat komprometterad), så att riskpolicyer slår till. Spärra kontot också.",
        [U], r"""
$us = @(Users $Anvandare)
Graph 'identityProtection/riskyUsers/confirmCompromised' -Method POST -Body @{ userIds = @($us.id) } | Out-Null
$us | ForEach-Object { Done $_.userPrincipalName 'bekräftad komprometterad' }
"""),
    act("mfa-status", "MFA-status i tenanten", "Säkerhet", "läs", "AuditLog.Read.All UserAuthenticationMethod.Read.All",
        "Vilka som har MFA registrerat, med vilka metoder, och om de är admins. Kräver Entra ID P1.",
        ["BaraUtan | Bara de utan MFA? | Nej | Nej, Ja"], r"""
$rows = Graph 'reports/authenticationMethods/userRegistrationDetails?$top=999' -All
if ($BaraUtan -eq 'Ja') { $rows = $rows | Where-Object { -not $_.isMfaRegistered } }
$rows | Sort-Object userDisplayName | ForEach-Object { [pscustomobject][ordered]@{ Namn = $_.userDisplayName; UPN = $_.userPrincipalName
  MFA = JaNej $_.isMfaRegistered; Metoder = (@($_.methodsRegistered) -join ', '); Standard = $_.defaultMfaMethod; Admin = JaNej $_.isAdmin; SSPR = JaNej $_.isSsprRegistered
  Typ = $(if ($_.userType -eq 'guest') { 'gäst' } else { 'medlem' }) } }
"""),
    # ─────────── Tenant ───────────
    act("tenant-licenser", "Licensöversikt", "Tenant", "läs", "Organization.Read.All",
        "Alla licenser: köpta, använda, lediga och status.", [], r"""
Graph 'subscribedSkus' | Sort-Object skuPartNumber | ForEach-Object { [pscustomobject][ordered]@{ Licens = $_.skuPartNumber
  Anvanda = $_.consumedUnits; Totalt = $_.prepaidUnits.enabled; Lediga = $_.prepaidUnits.enabled - $_.consumedUnits
  Status = $_.capabilityStatus; Varning = $_.prepaidUnits.warning; Avstangda = $_.prepaidUnits.suspended; SkuId = $_.skuId } }
"""),
    act("tenant-licens-anvandare", "Vilka har licensen", "Tenant", "läs", "User.Read.All Organization.Read.All",
        "Alla användare med en viss licens, och om den är direkt eller via grupp.", ["Licens | Vilken licens? | | @sku"], r"""
$sku = Sku $Licens
Graph ('users?$top=999&$count=true&$select=displayName,userPrincipalName,accountEnabled,licenseAssignmentStates&$filter=' + (Enc "assignedLicenses/any(x:x/skuId eq $($sku.skuId))")) -All -Eventual |
  Sort-Object displayName | ForEach-Object { $st = @($_.licenseAssignmentStates | Where-Object skuId -eq $sku.skuId)
    [pscustomobject][ordered]@{ Namn = $_.displayName; UPN = $_.userPrincipalName; Aktiv = JaNej $_.accountEnabled
      Via = $(if ($st | Where-Object assignedByGroup) { 'grupp' } else { 'direkt' }); Licens = $sku.skuPartNumber } }
"""),
    act("tenant-domaner", "Domäner", "Tenant", "läs", "Domain.Read.All",
        "Domänerna i tenanten: verifierade, standard, federerade och vad de används till.", [], r"""
Graph 'domains' | Sort-Object { -not $_.isDefault }, id | ForEach-Object { [pscustomobject][ordered]@{ Doman = $_.id; Verifierad = JaNej $_.isVerified
  Standard = JaNej $_.isDefault; Initial = JaNej $_.isInitial; Inloggning = $(if ($_.authenticationType -eq 'Federated') { 'federerad' } else { 'molnet' })
  Tjanster = (@($_.supportedServices) -join ', ') } }
"""),
    act("tenant-info", "Företagsinfo", "Tenant", "läs", "Organization.Read.All Policy.Read.All",
        "Namn, tenant-id, land, skapad, katalogsynk och om säkerhetsstandarder är på.", [], r"""
$o = @(Graph 'organization')[0]
$sd = $null; try { $sd = (Graph 'policies/identitySecurityDefaultsEnforcementPolicy').isEnabled } catch { }
$rows = [ordered]@{ 'Namn' = $o.displayName; 'Tenant-ID' = $o.id; 'Land' = $o.countryLetterCode; 'Skapad' = D $o.createdDateTime
  'Standarddomän' = ($o.verifiedDomains | Where-Object isDefault).name; 'Initial domän' = ($o.verifiedDomains | Where-Object isInitial).name
  'Katalogsynk (AD Connect)' = $(if ($o.onPremisesSyncEnabled) { "på, senast $(D $o.onPremisesLastSyncDateTime)" } else { 'av' })
  'Säkerhetsstandarder' = $(if ($null -eq $sd) { 'okänt' } else { JaNej $sd })
  'Tekniska kontakter' = (@($o.technicalNotificationMails) -join ', '); 'Integritetskontakt' = $o.privacyProfile.contactEmail }
foreach ($k in $rows.Keys) { [pscustomobject][ordered]@{ Installning = $k; Varde = $rows[$k] } }
"""),
    act("tenant-installningar", "Företagsinställningar", "Tenant", "läs", "Policy.Read.All",
        "Vad användare får göra: registrera appar, skapa grupper och tenants, bjuda in gäster, samtycka till appar m.m.", [], r"""
$a = Graph 'policies/authorizationPolicy'
if ($a.value) { $a = @($a.value)[0] }
$d = $a.defaultUserRolePermissions
$sd = $null; try { $sd = (Graph 'policies/identitySecurityDefaultsEnforcementPolicy').isEnabled } catch { }
$guest = @{ 'a0b1b346-4d3e-4e8b-98f8-753987be4970' = 'samma som medlemmar'; '10dae51f-b6af-4016-8d66-8c2a99b929b3' = 'begränsad (standard)'; '2af84b1e-32c8-42b7-82bc-daa82404023b' = 'mest begränsad' }
$inv = @{ 'none' = 'ingen'; 'adminsAndGuestInviters' = 'bara admins och gästinbjudare'; 'adminsGuestInvitersAndAllMembers' = 'admins och alla medlemmar'; 'everyone' = 'alla, även gäster' }
$consent = @($d.permissionGrantPoliciesAssigned | Where-Object { $_ -like 'ManagePermissionGrantsForSelf.*' })
$rows = [ordered]@{
  'Användare får registrera appar' = JaNej $d.allowedToCreateApps; 'Användare får skapa säkerhetsgrupper' = JaNej $d.allowedToCreateSecurityGroups
  'Användare får skapa tenants' = JaNej $d.allowedToCreateTenants; 'Användare får läsa andra användare' = JaNej $d.allowedToReadOtherUsers
  'Användare får läsa BitLocker-nyckel för egna enheter' = JaNej $d.allowedToReadBitlockerKeysForOwnedDevice
  'Vem får bjuda in gäster' = $inv[$a.allowInvitesFrom]; 'Gästers åtkomst' = $guest[$a.guestUserRoleId]
  'Användarsamtycke till appar' = $(if ($consent.Count) { ($consent -replace 'ManagePermissionGrantsForSelf.', '') -join ', ' } else { 'av' })
  'Självbetjäning av lösenord för admins' = JaNej $a.allowedToUseSSPR; 'Blockera MSOnline PowerShell' = JaNej $a.blockMsolPowerShell
  'Säkerhetsstandarder' = $(if ($null -eq $sd) { 'okänt' } else { JaNej $sd }) }
foreach ($k in $rows.Keys) { [pscustomobject][ordered]@{ Installning = $k; Varde = $rows[$k] } }
"""),
    act("tenant-installning-andra", "Ändra företagsinställning", "Tenant", "ändrar",
        "Policy.ReadWrite.Authorization Policy.ReadWrite.SecurityDefaults Policy.Read.All",
        "Slå på eller av: appregistrering, grupp- och tenantskapande, gästinbjudningar, användarsamtycke, säkerhetsstandarder.",
        ["Installning | Vilken inställning? | Användare får registrera appar | Användare får registrera appar, "
         "Användare får skapa säkerhetsgrupper, Användare får skapa tenants, Medlemmar får bjuda in gäster, "
         "Användarsamtycke till appar (låg risk), Säkerhetsstandarder, Blockera MSOnline PowerShell",
         "Varde | Ja eller nej? | Nej | Nej, Ja"], r"""
$on = ($Varde -eq 'Ja')
switch ($Installning) {
  'Användare får registrera appar' { Graph 'policies/authorizationPolicy' -Method PATCH -Body @{ defaultUserRolePermissions = @{ allowedToCreateApps = $on } } | Out-Null }
  'Användare får skapa säkerhetsgrupper' { Graph 'policies/authorizationPolicy' -Method PATCH -Body @{ defaultUserRolePermissions = @{ allowedToCreateSecurityGroups = $on } } | Out-Null }
  'Användare får skapa tenants' { Graph 'policies/authorizationPolicy' -Method PATCH -Body @{ defaultUserRolePermissions = @{ allowedToCreateTenants = $on } } | Out-Null }
  'Medlemmar får bjuda in gäster' { Graph 'policies/authorizationPolicy' -Method PATCH -Body @{ allowInvitesFrom = $(if ($on) { 'adminsGuestInvitersAndAllMembers' } else { 'adminsAndGuestInviters' }) } | Out-Null }
  'Användarsamtycke till appar (låg risk)' {
    $a = Graph 'policies/authorizationPolicy'
    if ($a.value) { $a = @($a.value)[0] }
    $keep = @($a.defaultUserRolePermissions.permissionGrantPoliciesAssigned | Where-Object { $_ -notlike 'ManagePermissionGrantsForSelf.*' })
    if ($on) { $keep += 'ManagePermissionGrantsForSelf.microsoft-user-default-low' }
    Graph 'policies/authorizationPolicy' -Method PATCH -Body @{ defaultUserRolePermissions = @{ permissionGrantPoliciesAssigned = @($keep) } } | Out-Null }
  'Säkerhetsstandarder' { Graph 'policies/identitySecurityDefaultsEnforcementPolicy' -Method PATCH -Body @{ isEnabled = $on } | Out-Null }
  'Blockera MSOnline PowerShell' { Graph 'policies/authorizationPolicy' -Method PATCH -Body @{ blockMsolPowerShell = $on } | Out-Null }
  default { throw "Okänd inställning: $Installning" }
}
Done $Installning $(if ($on) { 'ja' } else { 'nej' })
"""),
    # ─────────── Lokalt ───────────
    act("lokalt-session", "Min inloggning", "Lokalt", "lokal", "",
        "Vem du är inloggad som, i vilken tenant och med vilka scopes.", [], r"""
$c = Get-MgContext
if (-not $c) { Write-Warning 'Inte inloggad i Microsoft Graph än.'; return }
[pscustomobject][ordered]@{ Konto = $c.Account; Tenant = $c.TenantId; Inloggning = "$($c.AuthType) $($c.TokenCredentialType)".Trim()
  Sparas = "$($c.ContextScope)"; Scopes = ((@($c.Scopes) | Sort-Object) -join ', ') }
"""),
    act("lokalt-moduler", "Moduler", "Lokalt", "lokal", "",
        "PowerShell och modulerna Euppa använder, med version.", [], r"""
[pscustomobject][ordered]@{ Namn = 'PowerShell'; Version = "$($PSVersionTable.PSVersion)"; Behovs = 'ja' }
foreach ($m in 'Microsoft.Graph.Authentication', 'ImportExcel') {
  $x = Get-Module -ListAvailable $m | Sort-Object Version -Descending | Select-Object -First 1
  [pscustomobject][ordered]@{ Namn = $m; Version = $(if ($x) { "$($x.Version)" } else { 'saknas' }); Behovs = $(if ($m -eq 'ImportExcel') { 'för Excel-export' } else { 'ja' }) }
}
"""),
]


def writes(a):
    return a["typ"].startswith("änd")


def badge(a):
    if writes(a):
        return yellow("ändrar")
    if a["typ"].startswith("lok"):
        return cyan("lokal")
    return green("läs")


def find_action(name):
    n = name.lower()
    for a in ACTIONS:
        if a["id"] == n or a["namn"].lower() == n:
            return a
    hits = [a for a in ACTIONS if n in a["id"] or n in a["namn"].lower()]
    return hits[0] if len(hits) == 1 else None


def user_param(a):
    """Parametern som tar användare (för Användare…-flödet), eller None."""
    return next((p for p in a["params"] if p["uppslag"].rstrip("+") == "user" and not p["valfri"]), None)


def placeholders(t):
    t = t or {}
    dom = t.get("domain") or t.get("tenant", "")
    return {"tenant": t.get("tenant", ""), "domain": dom if not GUID.match(dom or "") else "",
            "upn": t.get("upn", ""), "kund": t.get("name") or "euppa", "datum": datetime.now().strftime("%Y-%m-%d")}


def fill(s, ph):
    return re.sub(r"\{(\w+)\}", lambda m: ph.get(m.group(1), m.group(0)), s)


def canon_choice(p, v):
    for ch in p["val"]:
        if ch.lower() == str(v).strip().lower():
            return ch
    raise Cancel(f"{p['namn']}={v} går inte – välj bland: {', '.join(p['val'])}")


def collect_params(a, t, given):
    """-> (värden, tenant). Uppslag (@user, @group …) kan logga in och därmed sätta tenanten."""
    vals = {}
    for p in a["params"]:
        default = fill(p["standard"], placeholders(t))
        if p["namn"] in given:
            v = given[p["namn"]]
            vals[p["namn"]] = canon_choice(p, v) if p["val"] and v else v
            continue
        if not ITTY:
            if not default and not p["valfri"]:
                raise Cancel(f"saknar {p['namn']}=…  ({p['fraga']})")
            vals[p["namn"]] = default
            continue
        if p["uppslag"] and not DRY:
            t, v = lookup_param(p, t, default)
        elif p["val"]:
            v = choose([dict(label=ch, value=ch) for ch in p["val"]], title=p["fraga"], start=default or None)
            if v is None:
                raise Cancel()
        else:
            v = ask_text(p["fraga"] + (" (valfritt)" if p["valfri"] else ""), default, required=not p["valfri"])
        vals[p["namn"]] = v
    return vals, t


def build_script(a, vals, t=None):
    head = []
    for k, v in vals.items():
        names = [LABELS[x] for x in str(v).split(",") if x in LABELS and LABELS[x] != x]
        head.append(f"${k} = {psq(v)}" + (f"   # {', '.join(names)[:80]}" if names else ""))
    return ("\n".join(head) + "\n" if head else "") + fill(a["ps"], placeholders(t))


FLAT_PS = r"""
$rows = @($global:EuppaLast | Where-Object { $_ -isnot [string] -and $_ -isnot [bool] } | ForEach-Object {
  $o = [ordered]@{}
  foreach ($pp in $_.PSObject.Properties) {
    $v = $pp.Value
    if ($v -is [datetime]) { $v = $v.ToString('yyyy-MM-dd HH:mm') }
    elseif ($v -is [System.Collections.IEnumerable] -and $v -isnot [string]) { $v = (@($v) | ForEach-Object { "$_" }) -join ', ' }
    elseif ($null -ne $v -and $v -isnot [string] -and $v -isnot [ValueType]) { $v = "$v" }
    $o[$pp.Name] = $v
  }
  [pscustomobject]$o })
if (-not $rows.Count) { throw 'Inget resultat att exportera.' }
"""


def export_result(a, t, how, open_=True):
    """Spara resultatet nyss som Excel (ImportExcel), CSV (semikolon, UTF-8) eller JSON. -> sökväg eller None."""
    name = slug(f"{(t or {}).get('name', 'euppa')}-{a['id']}-{datetime.now():%Y-%m-%d-%H%M}")
    if how == "json":
        fd, path = tempfile.mkstemp(prefix="euppa-", suffix=".json")
        os.close(fd)
        code = FLAT_PS + f"ConvertTo-Json -InputObject $rows -Depth 4 -Compress | Set-Content -LiteralPath {psq(path)} -Encoding utf8"
        ok, _ = WORKER.run(code, lambda k, s: None)
        try:
            rows = load_json(path, None) if ok else None
        finally:
            os.unlink(path)
        if isinstance(rows, dict):
            rows = [rows]
        return rows
    REPORTS.mkdir(parents=True, exist_ok=True)
    base = REPORTS / name
    if how == "excel":
        code = FLAT_PS + (f"$p = {psq(str(base) + '.xlsx')}\n"
                          "if (Get-Module -ListAvailable ImportExcel) {\n"
                          "  Import-Module ImportExcel\n"
                          f"  $rows | Export-Excel -Path $p -WorksheetName {psq(a['namn'][:30])} -TableName Euppa -TableStyle Medium2 "
                          "-AutoSize -FreezeTopRow -BoldTopRow\n"
                          f"}} else {{ $p = {psq(str(base) + '.csv')}; Write-Warning 'ImportExcel saknas – sparar CSV i stället (uppa install ImportExcel).'\n"
                          "  $rows | Export-Csv -Path $p -Delimiter ';' -Encoding utf8BOM -NoTypeInformation }\n"
                          "Write-Host \"Sparat $($rows.Count) rader i $p\"")
    else:
        code = FLAT_PS + (f"$p = {psq(str(base) + '.csv')}\n"
                          "$rows | Export-Csv -Path $p -Delimiter ';' -Encoding utf8BOM -NoTypeInformation\n"
                          "Write-Host \"Sparat $($rows.Count) rader i $p\"")
    ok, _, lines = run_ps("Sparar resultatet", code, show_output=False)
    saved = next((t_[len("Sparat "):] for k, t_ in lines if k == "H" and t_.startswith("Sparat ")), "")
    if not ok or not saved:
        print_output(lines, quiet_host=True)
        return None
    path = saved.split(" rader i ", 1)[-1]
    out("  " + green("✔") + f" Sparat {saved}")
    if open_ and ITTY:
        open_file(path)
    return path


def offer_export(a, t):
    """Efter en läsåtgärd med flera rader: exportera direkt, eller tillbaka."""
    global DIRTY
    try:
        v = choose([dict(label="Tillbaka", value="back"),
                    dict(label="Exportera till Excel", hint="formaterad .xlsx i ~/Documents/Euppa", value="excel", key="x"),
                    dict(label="Exportera till CSV", hint="semikolon, UTF-8", value="csv", key="c")],
                   title="Vad vill du göra med resultatet?", footer="↑/↓ · enter · x Excel · c CSV · q tillbaka")
    except Cancel:
        return
    if v in (None, "back"):
        DIRTY = False            # användaren har redan sett resultatet – ingen extra paus
        return
    export_result(a, t, v)


LAST_ERROR = [""]


def run_action(a, given=None, yes=False):
    """Kör en åtgärd: frågor → inloggning → kommandot syns → (bekräfta) → kör → resultat."""
    t = active_tenant()
    w = width()
    out(f"  {accent('▍')}{bold(a['namn'])}  {badge(a)}" + (dim(f"  {t['name']}") if t else ""))
    for l in wrap(a["beskrivning"], w - 6, 4) if a["beskrivning"] else []:
        out("    " + dim(l))
    print()
    local = a["typ"].startswith("lok")
    looks = {p["uppslag"].rstrip("+") for p in a["params"] if p["uppslag"] and p["namn"] not in (given or {})}
    if looks and not local and not DRY and ITTY:
        # en inloggning med allt som behövs – uppslagen och själva åtgärden
        t2 = ensure(t, sorted(set(a["scopes"]).union(*[LOOKUPS[k][0] for k in looks])))
        if not t2:
            return False
        t = t2
    try:
        vals, t = collect_params(a, t, given or {})
    except Cancel as e:
        if e.args:
            out(red("  " + e.args[0]))
            LAST_ERROR[0] = e.args[0]
        out(dim("  Avbrutet."))
        return False
    if DRY:
        if not local and not is_connected(t, a["scopes"]):
            show_command(connect_code(t, a["scopes"]), "Först: logga in")
        show_command(build_script(a, vals, t))
        say("Torrkörning: inget kördes. Så här skulle det se ut.", "happy")
        return True
    # Logga in först (det ändrar inget) – då vet vi vilken tenant ändringen hamnar i innan du bekräftar.
    if not local:
        t2 = ensure(t, a["scopes"])
        if not t2:
            return False
        t = t2
    script = build_script(a, vals, t)
    show_command(script)
    if writes(a) and not yes:
        print()
        if not ITTY:
            say("Det här ändrar i tenanten. Lägg till --ja för att köra utan fråga.", "worry")
            LAST_ERROR[0] = "ändrar i tenanten – kräver --ja"
            return False
        say(f"Det här ändrar i {t['name']} ({t.get('tenant', '')}). Kolla kommandot ovan först." if t else
            "Det här ändrar något. Kolla kommandot ovan först.", "worry")
        if not confirm("Kör på riktigt?", False):
            say("Okej, inget ändrat.", "happy")
            return False
    remember_action(a["id"])
    ok, n, _ = run_ps(f"Kör {a['namn']}", script)
    print()
    if ok:
        if writes(a):
            celebrate("Klart! Ändringen är gjord.")
        else:
            say(f"Klart! {n} rader." if n != 1 else "Klart! 1 rad.", "proud")
    else:
        say("Det gick inte hela vägen. Felet står ovanför i rött.", "sad")
    if n and EXPORT_NOW == "json":
        rows = export_result(a, t, "json") or []
        emit_json({"ok": ok, "tenant": (t or {}).get("name", ""), "action": a["id"], "rows": rows})
    elif EXPORT_NOW == "json":
        emit_json({"ok": ok, "tenant": (t or {}).get("name", ""), "action": a["id"], "rows": []})
    elif n and EXPORT_NOW:
        export_result(a, t, EXPORT_NOW, open_=False)
    elif ok and n > 1 and ITTY and not writes(a):
        offer_export(a, t)
    return ok


def remember_action(aid):
    rec = [x for x in config().get("recent", []) if x != aid]
    set_config(recent=([aid] + rec)[:6])


# ───────────────────────────── menyer ─────────────────────────────
def area_of(name):
    return next((x for x in AREAS if x["namn"] == name), {"namn": name, "ikon": name[:2].upper(), "farg": 205,
                                                          "beskrivning": ""})


def action_rows(acts):
    rows = [dict(label=a["namn"], hint=a["beskrivning"], value=a["id"], badge=badge(a)) for a in acts if not writes(a)]
    w = [dict(label=a["namn"], hint=a["beskrivning"], value=a["id"], badge=badge(a)) for a in acts if writes(a)]
    if rows and w:
        rows.append(("head", "ändrar"))
    return rows + w


def action_search_rows():
    return [dict(label=a["namn"], hint=f"{a['omrade']} · {a['beskrivning']}", value="a:" + a["id"], badge=badge(a),
                 search=a["id"]) for a in ACTIONS]


def run_by_value(v):
    a = find_action(v[2:])
    if a:
        run_action(a)


def areas_menu(start=None):
    """Två steg: område → åtgärd. Efter körning stannar man i samma område."""
    last_area = start
    while True:
        rows = []
        for ar in AREAS:
            its = [a for a in ACTIONS if a["omrade"] == ar["namn"]]
            if its:
                rows.append(dict(label=ar["namn"], hint=ar["beskrivning"], value=ar["namn"], badge=dim(str(len(its)))))
        rows += [("head", ""), dict(label="Tillbaka", value=None, key="q")]
        cat = choose(rows, title="Områden · allt i Entra ID", start=last_area, status=conn_status, screen=True,
                     search_rows=action_search_rows())
        if cat is None:
            return
        if cat.startswith("a:"):
            run_by_value(cat)
            continue
        last_area = cat
        its = [a for a in ACTIONS if a["omrade"] == cat]
        last = None
        while True:
            v = choose(action_rows(its) + [("head", ""), dict(label="Tillbaka", value=None, key="q")],
                       title=cat, start=last, status=conn_status, screen=True,
                       footer="↑/↓ · enter kör · siffra väljer · q tillbaka" if width() >= 50 else None)
            if not v:
                break
            last = v
            run_action(find_action(v))


def users_menu(preselected=None):
    """Användare…: lista tenantens användare, kryssa en eller flera, gör något med dem."""
    t = active_tenant()
    if preselected:
        chosen = list(preselected)
    else:
        t, users = run_lookup("alluser", t, "")
        if users is None:
            say("Kunde inte hämta användarna. Logga in först (Logga in…).", "sad")
            return
        if not users:
            say("Hittade inga användare i tenanten.", "worry")
            return
        new_screen()
        chosen = pick_many(users, f"Användare i {t['name']}" if t else "Användare")
        if not chosen:
            return
    while True:
        names = ", ".join(LABELS.get(u, u) for u in chosen[:3]) + (f" och {len(chosen) - 3} till" if len(chosen) > 3 else "")
        acts = sorted([a for a in ACTIONS if user_param(a)], key=lambda a: (writes(a), [x["namn"] for x in AREAS].index(a["omrade"])
                                                                          if a["omrade"] in [x["namn"] for x in AREAS] else 99))
        rows, head = [], None
        for a in acts:
            h = "Ändra" if writes(a) else "Visa"
            if h != head:
                rows.append(("head", h))
                head = h
            rows.append(dict(label=a["namn"], hint=a["omrade"] + " · " + a["beskrivning"], value=a["id"], badge=badge(a)))
        rows += [("head", ""), dict(label="Välj andra användare…", value="__again", key="v"),
                 dict(label="Tillbaka", value=None, key="q")]
        v = choose(rows, title=f"Vad vill du göra med {names}?", screen=True,
                   msg=f"{len(chosen)} {'användare valda' if len(chosen) != 1 else 'användare vald'}. Allt körs på en gång – "
                       "ändringar bekräftar du en gång för alla.", footer="↑/↓ · enter väljer · v välj om · q tillbaka")
        if not v:
            return
        if v == "__again":
            return users_menu()
        a = find_action(v)
        run_action(a, {user_param(a)["namn"]: ",".join(chosen)})


def tenant_rows():
    act_ = config().get("active")
    w = width()
    ts = sorted(tenants(), key=lambda t: t.get("last") or t.get("created") or "", reverse=True)
    rows = []
    for t in ts:
        hint = t.get("tenant", "")
        if w >= 96 and t.get("upn"):
            hint += "  · " + t["upn"]
        if w >= 120 and t.get("last"):
            hint += "  · " + ago(t["last"])
        rows.append(dict(label=t["name"], hint=hint, value=t["name"], badge=green("aktiv") if t["name"] == act_ else ""))
    return rows


def ago(iso):
    try:
        s = (datetime.now() - datetime.fromisoformat(iso)).total_seconds()
    except Exception:
        return ""
    if s < 3600:
        return "nyss"
    if s < 86400:
        return f"{int(s // 3600)} h sedan"
    return f"{int(s // 86400)} d sedan"


def add_tenant():
    traw = ask_text("Tenantens domän eller id? (contoso, contoso.onmicrosoft.com, contoso.se eller en GUID)", required=True,
                    check=lambda v: "Inga mellanslag." if " " in v.strip() else None)
    dom = tenant_domain(traw)
    ts = tenants()
    hit = next((x for x in ts if dom in (x.get("tenant", "").lower(), x.get("tenant_id", "").lower())), None)
    if hit:
        say(f"{dom} finns redan i listan som {hit['name']}.", "happy")
        set_config(active=hit["name"])
        return hit
    base = dom.split(".")[0].capitalize() if not GUID.match(dom) else "Tenant"
    name = ask_text("Vad ska den heta i listan?", unique_name(base, ts))
    t = dict(name=unique_name(name, ts), tenant=dom, tenant_id=dom if GUID.match(dom) else "", upn="",
             auth=config().get("auth", "interactive"), created=datetime.now().isoformat(timespec="seconds"))
    ts.append(t)
    save_tenants(ts)
    use_tenant(t, quiet=True)
    say(f"{t['name']} är tillagd och vald. Jag loggar in när en åtgärd behöver det.", "proud")
    return t


def use_tenant(t, quiet=False):
    update_tenant(t["name"], last=datetime.now().isoformat(timespec="seconds"))
    if config().get("active") != t["name"]:
        set_config(active=t["name"])
        if WORKER.conn:
            WORKER.stop()
            WORKER.warm()
    if not quiet:
        say(f"Nu jobbar vi i {t['name']} ({t.get('tenant', '')}).", "proud")


def tenants_menu():
    while True:
        ts = tenants()
        rows = tenant_rows()
        if rows:
            rows.append(("head", ""))
        rows += [dict(label="Lägg till tenant…", hint="oftast onödigt – inloggning sparar tenanten", value="__new", key="n"),
                 dict(label="Tillbaka", value=None, key="q")]
        name = choose(rows, title="Tenants" + (f" · {len(ts)} st" if ts else ""), status=conn_status, screen=True,
                      msg="" if ts else "Listan fylls på av sig själv när du loggar in. Kör en åtgärd, eller lägg "
                                        "till en tenant här.", footer="enter väljer · n ny · q tillbaka")
        if name is None:
            return
        if name == "__new":
            try:
                add_tenant()
            except Cancel:
                out(dim("  Avbrutet."))
            continue
        t = find_tenant(name)
        if not t:
            continue
        out(f"  {accent('▍')}{bold(t['name'])}")
        for k, v in (("Domän", t.get("domain") or t.get("tenant") or "–"), ("Tenant-ID", t.get("tenant_id") or "–"),
                     ("Konto", t.get("upn") or "–"), ("Inloggning", AUTHS[auth_of(t)])):
            out("    " + dim(pad(k, 14)) + v)
        other = "device" if auth_of(t) == "interactive" else "interactive"
        a = choose([dict(label="Jobba här", hint="gör den aktiv", value="use"),
                    dict(label=f"Logga in med {AUTHS[other].lower()}", hint="byter inloggningssätt för tenanten", value="auth"),
                    dict(label="Ta bort ur listan…", hint="inget ändras i själva tenanten", value="del"),
                    dict(label="Tillbaka", value=None, key="q")], title=t["name"])
        if a == "use":
            use_tenant(t, quiet=True)
            return
        if a == "auth":
            update_tenant(t["name"], auth=other)
            if WORKER.conn and WORKER.conn[0] == t["name"]:
                WORKER.stop()
                WORKER.warm()
            say(f"{t['name']} loggar nu in med {AUTHS[other].lower()}.", "happy")
        elif a == "del" and confirm(f"Ta bort {t['name']} ur listan?", False):
            save_tenants([x for x in tenants() if x["name"] != t["name"]])
            if config().get("active") == t["name"]:
                set_config(active=None)
                WORKER.stop()
                WORKER.warm()
            say(f"{t['name']} är borttagen ur listan.", "happy")


def connect_menu():
    t = active_tenant()
    keep = keep_logins()
    auth = auth_of(t)
    rows = [dict(label="Logga in", hint="läsbehörighet för nästan allt i Euppa på en gång", value="read"),
            dict(label="Logga in med egna scopes…", hint="t.ex. User.ReadWrite.All,Group.ReadWrite.All", value="scopes"),
            dict(label="Logga in i en annan tenant", hint="välj konto i webbläsaren", value="other"),
            dict(label="Inloggningssätt: " + AUTHS[auth].lower(), hint="byt till " +
                 AUTHS["device" if auth == "interactive" else "interactive"].lower(), value="auth"),
            dict(label="Kom ihåg inloggningar: " + ("på" if keep else "av"),
                 hint=("sparas i nyckelringen mellan körningar" if keep else "bara medan Euppa är igång – ingen nyckelringsfråga"),
                 value="keep"),
            dict(label="Logga ut", hint="Disconnect-MgGraph", value="off", key="u"),
            dict(label="Tillbaka", value=None, key="q")]
    v = choose(rows, title=f"Logga in · {t['name'] if t else 'ingen tenant vald'}", status=conn_status, screen=True)
    try:
        if v == "read":
            ensure(t, READ_SCOPES)
        elif v == "scopes":
            sc = ask_text("Scopes (kommaseparerade)", "User.Read.All,Group.Read.All")
            ensure(t, [s.strip() for s in sc.split(",") if s.strip()])
        elif v == "other":
            WORKER.stop()
            WORKER.warm()
            ensure(None, READ_SCOPES)
        elif v == "auth":
            new = "device" if auth == "interactive" else "interactive"
            if t:
                update_tenant(t["name"], auth=new)
            set_config(auth=new)
            WORKER.stop()
            WORKER.warm()
            say(f"Nu loggar jag in med {AUTHS[new].lower()}.", "happy")
        elif v == "keep":
            set_keep(not keep)
        elif v == "off":
            logout()
    except Cancel:
        out(dim("  Avbrutet."))


def set_keep(on):
    set_config(keep_logins=bool(on))
    WORKER.stop()
    if on:
        say("Inloggningar sparas nu mellan körningar (Graph lägger token i nyckelringen). Får du en "
            "nyckelringsfråga varje gång: välj 'Tillåt alltid', eller installera en signerad pwsh.", "happy")
    else:
        say("Inloggningar sparas inte längre – du loggar in en gång per körning och slipper nyckelringen.", "happy")


def greeting():
    ts = tenants()
    t = active_tenant()
    if not ts:
        return ("Hej! Jag är Euppa. Jag gör allt i Entra ID: användare, grupper, enheter, appar, roller och villkorlig "
                "åtkomst. Välj något – jag loggar in när det behövs."), "happy"
    if not t:
        return "Hej! Jag är Euppa. Välj en tenant, eller kör en åtgärd så loggar jag in.", "happy"
    if WORKER.conn and WORKER.conn[0] == t["name"]:
        return "Inloggad och redo. Vad blir nästa steg?", "happy"
    return f"Hej! Jag är Euppa. Välkommen tillbaka till {t['name']}. Börja skriva för att söka bland allt jag kan.", "happy"


def main_menu():
    last = None
    while True:
        msg, mood = greeting()
        t = active_tenant()
        rows = []
        by_id = {a["id"]: a for a in ACTIONS}
        recent = [by_id[x] for x in config().get("recent", []) if x in by_id][:3]
        if t and recent:
            rows.append(("head", f"Fortsätt där du lämnade · {t['name']}"))
            rows += [dict(label=a["namn"], hint=a["omrade"], value="a:" + a["id"], badge=badge(a)) for a in recent]
            rows.append(("head", ""))
        rows += [dict(label="Användare…", hint="lista, kryssa en eller flera och gör något med dem", value="users", key="a"),
                 dict(label="Områden…", hint=f"{len(ACTIONS)} åtgärder: grupper, enheter, appar, roller, villkorlig åtkomst …",
                      value="areas", key="o"),
                 dict(label="Tenants…", hint=(f"{t['name']} · " if t else "") + f"{len(tenants())} sparade", value="tenants",
                      key="t"),
                 dict(label="Logga in…", hint="Microsoft Graph · inloggningssätt · kom ihåg", value="connect", key="l"),
                 dict(label="Avsluta", value="quit", key="q")]
        v = choose(rows, msg=msg, mood=mood, status=conn_status, start=last or ("a:" + recent[0]["id"] if t and recent else "users"),
                   footer="↑/↓ · enter · siffra/bokstav väljer · q avslutar",
                   sub=f"euppa {VERSION}" + ("  ·  torrkörning" if DRY else ""), screen=True,
                   search_rows=action_search_rows())
        last = v
        if v in (None, "quit"):
            print()
            if WORKER.conn and keep_logins() and ITTY and confirm("Logga ut från Microsoft Graph innan du går?", False):
                logout()
            say(random.choice(["Hej då!", "Vi ses!", "Hej då! Glöm inte break-glass-kontot."]), "happy")
            print()
            return
        try:
            if v.startswith("a:"):
                run_by_value(v)
            elif v == "areas":
                areas_menu()
            elif v == "users":
                users_menu()
            elif v == "tenants":
                tenants_menu()
            elif v == "connect":
                connect_menu()
        except Cancel as e:
            if e.args:
                out(red("  " + e.args[0]))
            out(dim("  Avbrutet."))


# ───────────────────────────── CLI ─────────────────────────────
def list_tenants():
    ts = tenants()
    if not ts:
        say("Inga tenants än. De sparas automatiskt när du loggar in (euppa logga-in), eller: euppa ny.", "happy")
        return
    act_ = config().get("active")
    nw = min(max(len(t["name"]) for t in ts) + 2, 24)
    for t in sorted(ts, key=lambda x: x.get("last") or x.get("created") or "", reverse=True):
        mark = green("●") if t["name"] == act_ else dim("○")
        out(f"  {mark} {pad(bold(cut(t['name'], nw)), nw)}{pad(cut(t.get('tenant', ''), 32), 34)}" + dim(t.get("upn", "")))


def list_actions(as_json):
    if as_json:
        emit_json({"version": VERSION, "areas": AREAS, "actions": [{k: v for k, v in a.items() if k != "ps"} for a in ACTIONS],
                   "lookups": {k: v[1] for k, v in LOOKUPS.items()}})
        return
    w = width()
    nw = min(max(len(a["id"]) for a in ACTIONS) + 2, 34)
    for ar in AREAS:
        out("  " + accent("▍") + bold(ar["namn"]) + dim("  " + ar["beskrivning"]))
        for a in [x for x in ACTIONS if x["omrade"] == ar["namn"]]:
            out(f"    {pad(cut(a['id'], nw), nw)}{pad(badge(a), 8)}" + (dim(a["namn"]) if w >= 80 else ""))


def main(argv=None):
    import argparse
    global ANIM, DRY, EXPORT_NOW, TTY, COLOR
    argv = list(sys.argv[1:] if argv is None else argv)
    pre = argparse.ArgumentParser(add_help=False)
    pre.add_argument("--no-anim", action="store_true")
    pre.add_argument("-n", "--torr", action="store_true")
    known, rest = pre.parse_known_args(argv)
    if known.no_anim:
        ANIM = False
    DRY = known.torr

    ap = argparse.ArgumentParser(prog="euppa", description="Euppa – allt i Entra ID med några knapptryck.")
    ap.add_argument("--version", action="version", version="euppa " + VERSION)
    sub = ap.add_subparsers(dest="cmd")
    sub.add_parser("tenants", aliases=["kunder"], help="tenants du loggat in i")
    u = sub.add_parser("tenant", aliases=["kund", "use"], help="växla tenant")
    u.add_argument("namn")
    sub.add_parser("ny", help="lägg till en tenant för hand")
    li = sub.add_parser("logga-in", aliases=["login"], help="logga in i Microsoft Graph")
    li.add_argument("--scopes", default="", help="kommaseparerade (standard: läs nästan allt)")
    li.add_argument("--enhetskod", action="store_true", help="logga in med enhetskod (device code)")
    li.add_argument("--tenant", default="", help="tenantens domän eller id")
    la = sub.add_parser("åtgärder", aliases=["atgarder", "lista"], help="lista åtgärderna")
    la.add_argument("--json", action="store_true", help="maskinläsbart (för Ruppa)")
    an = sub.add_parser("användare", aliases=["anvandare", "users"], help="välj användare och gör något med dem")
    an.add_argument("upn", nargs="*", help="hoppa över listan: de här användarna")
    up = sub.add_parser("uppslag", aliases=["lookup"], help="sök i tenanten: user, group, device, app, sp, role, ca, sku")
    up.add_argument("typ")
    up.add_argument("sok", nargs="?", default="")
    up.add_argument("--json", action="store_true", help="maskinläsbart (för Ruppa)")
    up.add_argument("--alla", action="store_true", help="user: alla användare i tenanten")
    k = sub.add_parser("kör", aliases=["kor", "run"], help="kör en åtgärd")
    k.add_argument("atgard")
    k.add_argument("par", nargs="*", help="Namn=värde")
    k.add_argument("--ja", action="store_true", help="bekräfta ändrande åtgärder utan att fråga")
    k.add_argument("--excel", action="store_true", help="spara resultatet som Excel i ~/Documents/Euppa")
    k.add_argument("--csv", action="store_true", help="spara resultatet som CSV i ~/Documents/Euppa")
    k.add_argument("--json", action="store_true", help="skriv resultatet som JSON (för Ruppa)")
    kh = sub.add_parser("kom-ihåg", aliases=["kom-ihag"], help="spara inloggningar mellan körningar: på eller av")
    kh.add_argument("val", choices=["på", "pa", "av", "on", "off"])
    sub.add_parser("logga-ut", aliases=["logout"], help="logga ut från Microsoft Graph")
    sub.add_parser("hej")
    a = ap.parse_args(rest)

    as_json = getattr(a, "json", False)
    if as_json:                           # allt utom JSON-raden till stderr, utan färger och animationer
        sys.stdout = sys.stderr
        TTY = COLOR = ANIM = False

    if a.cmd in ("åtgärder", "atgarder", "lista"):
        list_actions(as_json)
        return 0
    if a.cmd == "hej":
        print()
        say(random.choice(JOKES), random.choice(["happy", "proud", "wow", "left"]))
        print()
        return 0
    if a.cmd in ("tenants", "kunder"):
        list_tenants()
        return 0
    if a.cmd in ("kom-ihåg", "kom-ihag"):
        set_keep(a.val in ("på", "pa", "on"))
        return 0
    if a.cmd in ("tenant", "kund", "use"):
        t = find_tenant(a.namn)
        if not t:
            say(f"Hittar ingen tenant {a.namn}. Se euppa tenants.", "sad")
            return 1
        use_tenant(t)
        return 0

    if not (shutil.which("pwsh") or shutil.which("pwsh-preview")):
        say("Jag hittar inte pwsh (PowerShell 7). Installera det: Windows: winget install Microsoft.PowerShell · "
            "macOS: brew install --cask powershell · eller: uppa install powershell", "sad")
        if as_json:
            emit_json({"ok": False, "error": "pwsh saknas"})
        return 1

    try:
        if a.cmd is None:
            if not ITTY or not TTY:
                print(__doc__)
                return 0
            signal.signal(signal.SIGINT, signal.default_int_handler)
            WORKER.warm()
            main_menu()
            return 0
        if a.cmd == "ny":
            if not ITTY:
                sys.exit("  behöver en riktig terminal")
            return 0 if add_tenant() else 1
        if a.cmd in ("logga-in", "login"):
            t = active_tenant()
            if a.tenant:
                t = find_tenant(a.tenant) or dict(name="", tenant=tenant_domain(a.tenant))
            if a.enhetskod:
                t = dict(t or {"name": ""}, auth="device")
            scopes = [s.strip() for s in a.scopes.split(",") if s.strip()] or READ_SCOPES
            t = ensure(t, scopes)
            if t and a.enhetskod and t.get("name"):
                update_tenant(t["name"], auth="device")
            return 0 if t else 1
        if a.cmd in ("logga-ut", "logout"):
            return 0 if logout() else 1
        if a.cmd in ("användare", "anvandare", "users"):
            if not ITTY:
                sys.exit("  behöver en riktig terminal (eller: euppa kör ÅTGÄRD Anvandare=a@x.se,b@x.se --ja)")
            WORKER.warm()
            users_menu(a.upn or None)
            return 0
        if a.cmd in ("uppslag", "lookup"):
            WORKER.warm()
            return 0 if cmd_lookup(a.typ, a.sok, as_json, a.alla) else 1
        if a.cmd in ("kör", "kor", "run"):
            act_ = find_action(a.atgard)
            if not act_:
                hits = [x["id"] for x in ACTIONS if a.atgard.lower() in x["id"] or a.atgard.lower() in x["namn"].lower()]
                say(f"Flera åtgärder matchar {a.atgard}: {', '.join(hits[:6])}." if hits else
                    f"Hittar ingen åtgärd som heter {a.atgard}. Se euppa åtgärder.", "worry" if hits else "sad")
                if as_json:
                    emit_json({"ok": False, "error": "okänd åtgärd", "rows": []})
                return 1
            given = dict(x.split("=", 1) for x in a.par if "=" in x)
            EXPORT_NOW = "json" if as_json else "excel" if a.excel else "csv" if a.csv else None
            WORKER.warm()
            ok = run_action(act_, given, yes=a.ja)
            if as_json and not EMITTED:
                emit_json({"ok": False, "error": LAST_ERROR[0] or "inget kördes (inloggningen misslyckades – se loggen)", "rows": []})
            return 0 if ok else 1
    except Cancel as e:
        if e.args:
            out(red("  " + e.args[0]))
        out(dim("  Avbrutet."))
        if as_json:
            emit_json({"ok": False, "error": e.args[0] if e.args else "avbrutet", "rows": [], "items": []})
        return 1
    except KeyboardInterrupt:
        sys.stdout.write("\033[?25h\n")
        return 130
    finally:
        WORKER.stop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
