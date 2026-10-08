#!/bin/zsh
# 每卷生成完（含漏译补丁跑完）就自动校验并提交推送。
# 用法：./watch_and_commit.sh B D E F G
PY=/Users/glennchou/.workbuddy/binaries/python/envs/default/bin/python
cd "/Users/glennchou/FAR 23"

for s in "$@"; do
  f="FAR23_Subpart${s}_条款修订历史与背景分析_详版.docx"
  echo "[watch] 等待 $f … $(date '+%H:%M:%S')"

  # 1) 等文件出现
  while [[ ! -f "$f" ]]; do sleep 30; done

  # 2) 等补丁进程结束 + 文件稳定 3 分钟
  while true; do
    if pgrep -f "patch_untranslated.py $f" >/dev/null 2>&1; then sleep 20; continue; fi
    age=$(( $(date +%s) - $(stat -f %m "$f") ))
    [[ $age -ge 180 ]] && break
    sleep 30
  done
  echo "[watch] $f 落盘稳定 $(date '+%H:%M:%S')"

  # 3) 校验；仍有漏译就再补一轮
  for round in 1 2; do
    out=$($PY check_volume.py "$f")
    echo "$out"
    n=$(echo "$out" | grep -o '未译 [0-9]*' | grep -o '[0-9]*')
    if [[ "$n" == "0" ]]; then break; fi
    echo "[watch] 仍有 $n 条漏译，再补一轮"
    $PY patch_untranslated.py "$f" --workers 8 --rounds 3 2>&1 | tail -3
  done

  # 4) 提交推送
  git add "$f" FAR23_译文缓存.json
  if git diff --cached --quiet; then
    echo "[watch] 无改动，跳过"
  else
    meta=$(echo "$out" | head -3 | tr '\n' ' ')
    git commit -q -m "Subpart $s 详版（按 C 分部口径：五节+附录、原始论述直译）

$f
$meta" && git push origin main 2>&1 | tail -2
    echo "[watch] $s 已提交推送 $(date '+%H:%M:%S')"
  fi
done
echo "[watch] 全部提交完毕 $(date '+%H:%M:%S')"
