#!/usr/bin/env sh
# Avinstallera Euppa på macOS och Linux:
#   curl -fsSL https://raw.githubusercontent.com/pipebreak33/euppa/main/uninstall.sh | sh
# Från en lokal kopia:  ./uninstall.sh
#
# Tar bort programmet (~/.euppa/app) och kommandot ~/.local/bin/euppa.
# Behåller dina tenants och inställningar (~/.euppa) och rapporterna i ~/Documents/Euppa.
# EUPPA_PURGE=1 tar bort även ~/.euppa med dina data. EUPPA_YES=1 hoppar över frågan.
set -eu
START="$(date +%s)"

DIR="$HOME/.euppa"
SHIM="$HOME/.local/bin/euppa"

say()  { printf '  \033[38;5;205mE›\033[0m %s\n' "$1"; }
die()  { printf '  \033[31m✘\033[0m %s\n' "$1" >&2; exit 1; }

ask() {   # ask "Fråga" -> 0 om ja. Läser från terminalen även när skriptet kommer via curl | sh
  if [ -r /dev/tty ]; then
    printf '  %s [j/N] ' "$1"
    { read -r SVAR < /dev/tty; } 2>/dev/null || SVAR=""
    case "$SVAR" in j|J|ja|Ja|y|Y|yes) return 0 ;; esac
    return 1
  fi
  die "Ingen terminal att fråga i. Kör med EUPPA_YES=1 för att avinstallera utan fråga."
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
  say "Euppa är inte installerad här. Inget att ta bort."
  exit 0
fi

printf '\n  Det här tas bort:'
printf "$FOUND\n"
[ "${EUPPA_PURGE:-}" = "1" ] || printf '  Behålls: dina tenants och inställningar (~/.euppa) och rapporterna i ~/Documents/Euppa.\n'
printf '\n'
if [ "${EUPPA_YES:-}" != "1" ]; then
  ask "Avinstallera Euppa?" || { say "Avbrutet. Ingenting har tagits bort."; exit 0; }
fi

for t in $TARGETS; do rm -rf "$t"; done
# tom ~/.euppa: städa bort mappen också
rmdir "$DIR" 2>/dev/null || true

SECS=$(( $(date +%s) - START ))
printf '\n  \033[38;5;42m✔\033[0m Euppa är avinstallerad%s (%s s).\n\n' "" "$SECS"
