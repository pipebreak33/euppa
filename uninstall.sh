#!/usr/bin/env sh
# Avinstallera Euppa på macOS och Linux:
#   curl -fsSL https://raw.githubusercontent.com/pipebreak33/euppa/main/uninstall.sh | sh
# Från en lokal kopia:  ./uninstall.sh
#
# Tar bort programmet (~/.euppa/app) och kommandot ~/.local/bin/euppa.
# Behåller dina tenants och inställningar (~/.euppa) och rapporterna i ~/Documents/Euppa.
# EUPPA_PURGE=1 tar bort även ~/.euppa med dina data. EUPPA_YES=1 hoppar över frågan.
set -eu
# Language: TRUPPA_LANG, else ~/.truppa/config.json (shared by the whole troupe), else English
TL="${TRUPPA_LANG:-}"
if [ -z "$TL" ] && [ -f "$HOME/.truppa/config.json" ]; then
  TL="$(sed -n 's/.*"lang"[^"]*"\([a-z]*\)".*/\1/p' "$HOME/.truppa/config.json" 2>/dev/null | head -n 1)"
fi
t() { case "$TL" in sv*) printf '%s' "$2" ;; *) printf '%s' "$1" ;; esac; }
START="$(date +%s)"

DIR="$HOME/.euppa"
SHIM="$HOME/.local/bin/euppa"

say()  { printf '  \033[38;5;205mE›\033[0m %s\n' "$1"; }
die()  { printf '  \033[31m✘\033[0m %s\n' "$1" >&2; exit 1; }

ask() {   # ask "Fråga" -> 0 om ja. Läser från terminalen även när skriptet kommer via curl | sh
  if [ -r /dev/tty ]; then
    printf "$(t '  %s [y/N] ' '  %s [j/N] ')" "$1"
    { read -r SVAR < /dev/tty; } 2>/dev/null || SVAR=""
    case "$SVAR" in j|J|ja|Ja|y|Y|yes|Yes) return 0 ;; esac
    return 1
  fi
  die "$(t "No terminal to ask in. Run with EUPPA_YES=1 to uninstall without asking." "Ingen terminal att fråga i. Kör med EUPPA_YES=1 för att avinstallera utan fråga.")"
}

FOUND=""
TARGETS=""
if [ "${EUPPA_PURGE:-}" = "1" ]; then
  [ -e "$DIR" ] && { FOUND="$FOUND\n    $DIR (med dina data)"; TARGETS="$TARGETS $DIR"; }
else
  for p in app; do
    [ -e "$DIR/$p" ] && { FOUND="$FOUND\n    $DIR/$p"; TARGETS="$TARGETS $DIR/$p"; }
  done
fi
# startkommandot: bara om det är vårt (pekar på ~/.euppa)
if [ -e "$SHIM" ] && grep -qsF "/.euppa/" "$SHIM"; then FOUND="$FOUND\n    $SHIM"; TARGETS="$TARGETS $SHIM"; fi
if [ -z "$FOUND" ]; then
  say "$(t "Euppa is not installed here. Nothing to remove." "Euppa är inte installerad här. Inget att ta bort.")"
  exit 0
fi

printf "$(t '\n  This is removed:' '\n  Det här tas bort:')"
printf "$FOUND\n"
[ "${EUPPA_PURGE:-}" = "1" ] || printf "$(t '  Kept: your tenants and settings (~/.euppa) and the reports in ~/Documents/Euppa.\n' '  Behålls: dina tenants och inställningar (~/.euppa) och rapporterna i ~/Documents/Euppa.\n')"
printf '\n'
if [ "${EUPPA_YES:-}" != "1" ]; then
  ask "$(t "Uninstall Euppa?" "Avinstallera Euppa?")" || { say "$(t "Cancelled. Nothing has been removed." "Avbrutet. Ingenting har tagits bort.")"; exit 0; }
fi

for t in $TARGETS; do rm -rf "$t"; done
# tom ~/.euppa: städa bort mappen också
rmdir "$DIR" 2>/dev/null || true

SECS=$(( $(date +%s) - START ))
printf "$(t '\n  \033[38;5;42m✔\033[0m Euppa is uninstalled%s (%s s).\n\n' '\n  \033[38;5;42m✔\033[0m Euppa är avinstallerad%s (%s s).\n\n')" "" "$SECS"
