#!/bin/bash
# Send N requests to the convert service and report results.
# Assumes fake_ocr + service are already running.
#
# Usage: bash batch_process_test.sh <file.pptx> [total=100] [concurrency=5]

PPTX="${1:-}"
TOTAL="${2:-100}"
CONCURRENCY="${3:-5}"
URL="http://localhost:8000/ppt"
TMPDIR_R=$(mktemp -d)

if [[ -z "$PPTX" || ! -f "$PPTX" ]]; then
    echo "Usage: bash batch_process_test.sh <file.pptx> [total] [concurrency]"
    exit 1
fi

echo "file=$PPTX  total=$TOTAL  concurrency=$CONCURRENCY  target=$URL"
echo ""

send_one() {
    local idx=$1
    curl -s -o /dev/null -w "%{http_code} %{time_total}" \
        -X POST "$URL" -F "file=@$PPTX" --max-time 180 \
        > "$TMPDIR_R/$idx"
}

# fire requests in batches of CONCURRENCY
batch=()
for i in $(seq 1 $TOTAL); do
    send_one "$i" &
    batch+=($!)
    printf "\r  dispatched %d / %d" "$i" "$TOTAL"
    if (( ${#batch[@]} >= CONCURRENCY )); then
        wait "${batch[@]}"
        batch=()
        printf "\r  dispatched %d / %d (batch done)" "$i" "$TOTAL"
    fi
done
wait "${batch[@]}" 2>/dev/null
echo ""
echo ""

# tally results
ok=0; fail=0
elapsed_vals=()
for i in $(seq 1 $TOTAL); do
    read code elapsed < "$TMPDIR_R/$i"
    if [[ "$code" == "200" ]]; then
        (( ok++ ))
    else
        (( fail++ ))
        echo "  [FAIL] #$i -> HTTP $code"
    fi
    elapsed_vals+=("$elapsed")
done
rm -rf "$TMPDIR_R"

avg=$(echo "scale=1; ($(IFS=+; echo "${elapsed_vals[*]}"))/${#elapsed_vals[@]}" | bc)
echo "total=$TOTAL  success=$ok  fail=$fail  avg_latency=${avg}s"
