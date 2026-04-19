#!/bin/bash
# 编译 thesis.qmd 并应用后处理样式

set -e

OUTPUT_DIR="output"
OUTPUT_BASENAME="thesis.docx"
OUTPUT_DOCX="$OUTPUT_DIR/$OUTPUT_BASENAME"

mkdir -p "$OUTPUT_DIR"

echo "正在编译 thesis.qmd -> $OUTPUT_DOCX ..."
quarto render thesis.qmd --to docx --output "$OUTPUT_BASENAME"
mv -f "$OUTPUT_BASENAME" "$OUTPUT_DOCX"

echo "正在应用论文后处理样式..."
python3 scripts/thesis_post_process.py "$OUTPUT_DOCX" thesis.qmd

if command -v cmd.exe >/dev/null 2>&1 && command -v wslpath >/dev/null 2>&1 && command -v powershell.exe >/dev/null 2>&1; then
  echo "正在用 Word 打开 $OUTPUT_DOCX ..."
  DOCX_WIN_PATH="$(wslpath -w "$(realpath "$OUTPUT_DOCX")")"

  PS_SCRIPT="$(mktemp --suffix=.ps1)"
  cat > "$PS_SCRIPT" <<'EOF'
param(
  [Parameter(Mandatory = $true)]
  [string]$target
)

$ErrorActionPreference = 'Stop'

$wordExe = 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE'

$word = $null
try {
  $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
} catch {
  $word = $null
}

if ($word -ne $null) {
  foreach ($doc in @($word.Documents)) {
    if ($doc.FullName -ieq $target) {
      # 只关闭当前目标文档，不影响其他已打开文档；有改动则先保存
      $doc.Close([ref]-1)
      break
    }
  }
}

Start-Process -FilePath $wordExe -ArgumentList @($target)
EOF

  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$(wslpath -w "$PS_SCRIPT")" -target "$DOCX_WIN_PATH"
  rm -f "$PS_SCRIPT"
else
  echo "未检测到 cmd.exe/wslpath/powershell.exe，跳过自动打开 Word。"
fi

