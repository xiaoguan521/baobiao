#!/usr/bin/env sh
set -eu

PROJECT_ROOT="$(CDPATH= cd -- "$(dirname -- "$0")/.." && pwd)"
SOURCE_SKILL_DIR="$PROJECT_ROOT/.codex/skills/oracle-report-service-api"
TARGET_SKILL_DIR="${HOME}/.codex/skills/oracle-report-service-api"

mkdir -p "${HOME}/.codex/skills"
rm -rf "$TARGET_SKILL_DIR"
ln -s "$SOURCE_SKILL_DIR" "$TARGET_SKILL_DIR"

printf 'Installed project skill:\n%s -> %s\n' "$TARGET_SKILL_DIR" "$SOURCE_SKILL_DIR"
