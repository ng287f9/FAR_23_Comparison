#!/bin/zsh
# 一体化：按 C 分部详版口径逐卷生成 → 自检 → 补译 → 提交推送。
# 用法：./build_and_push.sh B D E F G AppA AppB AppC AppD AppF AppG AppH AppI AppJ
# 用法：./build_and_push.sh D E F G AppA …
#       FORCE=1 ./build_and_push.sh AppA …   # 强制重建已存在的卷
PY=/Users/glennchou/.workbuddy/binaries/python/envs/default/bin/python
cd "/Users/glennchou/FAR 23"
set -o pipefail

for s in "$@"; do
  f="FAR23_Subpart${s}_条款修订历史与背景分析_详版.docx"
  echo "===== $s 开始 $(date '+%H:%M:%S')"

  # 0) 已有的卷先自检，无漏译就只补提交（FORCE=1 时跳过，一律重建）
  if [[ -f "$f" && "$FORCE" != "1" ]]; then
    out=$($PY check_volume.py "$f" 2>/dev/null)
    n=$(echo "$out" | grep -o '未译 [0-9]*' | grep -o '[0-9]*')
    if [[ -z "$n" ]]; then n=1; fi
    if [[ "$n" != "0" ]]; then
      echo "  [补译] 现有卷有 $n 条漏译，先修"
      $PY patch_untranslated.py "$f" --workers 8 --rounds 3 2>&1 | tail -2
      out=$($PY check_volume.py "$f" 2>/dev/null)
      n=$(echo "$out" | grep -o '未译 [0-9]*' | grep -o '[0-9]*')
    fi
    if [[ "$n" == "0" ]]; then
      echo "  [跳过] $s 已完整"
      git add "$f" FAR23_译文缓存.json 2>/dev/null
      git diff --cached --quiet || {
        git commit -q -m "Subpart $s 详版（五节+附录、NPRM/FR 原始论述直译）

$(echo "$out" | head -3 | tr '\n' ' ')" && git push origin main 2>&1 | tail -1
        echo "  [已推送] $s"
      }
      continue
    fi
  fi

  # 1) 生成
  $PY -u build_clause_doc.py "$s" --full-bg --discuss 0 --workers 16 \
      > "_build_${s}.log" 2>&1
  if [[ $? -ne 0 || ! -f "$f" ]]; then
    echo "  [失败] $s，见 _build_${s}.log"; tail -6 "_build_${s}.log"; continue
  fi
  echo "  生成完成 $(date '+%H:%M:%S')：$(tail -2 "_build_${s}.log" | tr '\n' ' ')"

  # 2) 自检 + 补译（最多 3 轮）
  for r in 1 2 3; do
    out=$($PY check_volume.py "$f")
    echo "  $out"
    n=$(echo "$out" | grep -o '未译 [0-9]*' | grep -o '[0-9]*')
    [[ "$n" == "0" ]] && break
    echo "  [补译第 $r 轮] 剩余 $n 条"
    $PY patch_untranslated.py "$f" --workers 8 --rounds 3 2>&1 | tail -2
  done

  # 3) 提交推送
  git add "$f" FAR23_译文缓存.json
  if git diff --cached --quiet; then
    echo "  [无改动]"
  else
    git commit -q -m "Subpart $s 详版（五节+附录、NPRM/FR 原始论述直译）

$(echo "$out" | head -3 | tr '\n' ' ')" && git push origin main 2>&1 | tail -1
    echo "  [已推送] $s $(date '+%H:%M:%S')"
  fi
done
echo "===== 全部结束 $(date '+%H:%M:%S')"
