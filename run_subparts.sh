#!/bin/zsh
# 按 C 分部详版口径，依次生成剩余各卷（D E F G + 附录 A~J）。
# 与 watch_and_commit.sh 配合使用：本脚本只生成，后者负责校验/提交。
# 用法：./run_subparts.sh D E F G AppA AppB ...
PY=/Users/glennchou/.workbuddy/binaries/python/envs/default/bin/python
cd "/Users/glennchou/FAR 23"

for s in "$@"; do
  # 已存在且通过校验的卷直接跳过
  f="FAR23_Subpart${s}_条款修订历史与背景分析_详版.docx"
  if [[ -f "$f" ]]; then
    n=$($PY check_volume.py "$f" | grep -o '未译 [0-9]*' | grep -o '[0-9]*')
    if [[ "$n" == "0" ]]; then echo "[run] $s 已完成且无漏译，跳过"; continue; fi
  fi
  echo "[run] 开始 $s … $(date '+%H:%M:%S')"
  $PY -u build_clause_doc.py "$s" --full-bg --discuss 0 --workers 16 \
      > "_build_${s}.log" 2>&1
  if [[ $? -ne 0 ]]; then
    echo "[run] $s 失败，日志 _build_${s}.log"; tail -5 "_build_${s}.log"; continue
  fi
  echo "[run] $s 完成 $(date '+%H:%M:%S')：$(tail -2 "_build_${s}.log" | tr '\n' ' ')"
done
echo "[run] 全部生成结束 $(date '+%H:%M:%S')"
