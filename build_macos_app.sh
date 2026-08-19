#!/usr/bin/env bash
set -euo pipefail

version="${1:-local}"
project_root="$(cd "$(dirname "$0")" && pwd)"
build_venv="$project_root/.packaging-venv-macos"
python_exe="$build_venv/bin/python"

if [[ "$(uname -m)" != "arm64" || "$(python3 -c 'import platform; print(platform.machine())')" != "arm64" ]]; then
  echo "This package must be built on an Apple Silicon (arm64) macOS runner." >&2
  exit 1
fi

python3 -m venv "$build_venv"
"$python_exe" -m pip install --upgrade pip
"$python_exe" -m pip install -r "$project_root/requirements.txt" pyinstaller

"$python_exe" -m PyInstaller --noconfirm --clean --onedir --windowed \
  --name "Clinic Lead Collector" \
  --add-data "$project_root/app.py:." \
  --collect-all streamlit \
  --collect-all pandas \
  --collect-all pyarrow \
  --collect-all selenium \
  --collect-all webdriver_manager \
  --collect-all openpyxl \
  "$project_root/desktop_launcher.py"

"$python_exe" -m PyInstaller --noconfirm --clean --onefile \
  --name "Clinic Lead Updater" \
  "$project_root/updater.py"

app_bundle="$project_root/dist/Clinic Lead Collector.app"
resources="$app_bundle/Contents/Resources"
updater_binary="$resources/Clinic Lead Updater"
mkdir -p "$resources"
cp "$project_root/dist/Clinic Lead Updater" "$updater_binary"
chmod +x "$updater_binary"
printf '%s' "$version" > "$resources/version.txt"

# Ad-hoc signing keeps the arm64 bundle internally consistent. A paid Apple
# Developer certificate can replace this later for notarized distribution.
codesign --force --deep --sign - "$app_bundle"

zip_path="$project_root/dist/Clinic-Lead-Collector-macos-arm64.zip"
rm -f "$zip_path"
/usr/bin/ditto -c -k --sequesterRsrc --keepParent "$app_bundle" "$zip_path"

echo "Build complete: $zip_path"
