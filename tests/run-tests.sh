#!/bin/bash
# Full test phase for -Anonymize.
PWSH=${PWSH:-pwsh}
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO="$(dirname "$HERE")"
SCRIPT="$REPO/Export-EntraTenantDocs.ps1"
T="${TMPDIR:-/tmp}/entra-tenant-docs-tests"
export REPO T   # the quoted heredocs below read these from the environment
PASS=0; FAIL=0

ok()   { echo "PASS $1"; PASS=$((PASS+1)); }
bad()  { echo "FAIL $1"; echo "      $2"; FAIL=$((FAIL+1)); }
check(){ if [ "$2" = "$3" ]; then ok "$1"; else bad "$1" "expected [$3] got [$2]"; fi; }

rm -rf $T/out-* $T/hist-* ; mkdir -p $T

echo "--- 1. leak scan -----------------------------------------------------"
rm -rf $T/out-anon
$PWSH -NoProfile -File $SCRIPT -SampleData -Anonymize -OutputPath $T/out-anon >/dev/null 2>&1
if python3 $HERE/leakscan.py $T/out-anon > $T/leak.txt 2>&1; then
  ok "leak scan: no identifier from the snapshot or its history reaches any output"
  grep -q "SURVIVORS: all" $T/leak.txt && ok "survivors: SKUs, role names, policy names and counts kept" \
    || bad "survivors" "$(grep MISSING -A5 $T/leak.txt)"
else
  bad "leak scan" "$(cat $T/leak.txt)"
fi

echo
echo "--- 2. determinism ---------------------------------------------------"
rm -rf $T/out-s1 $T/out-s2 $T/out-s3
$PWSH -NoProfile -File $SCRIPT -SampleData -Anonymize -AnonymizeSalt fixed-salt -OutputPath $T/out-s1 >/dev/null 2>&1
$PWSH -NoProfile -File $SCRIPT -SampleData -Anonymize -AnonymizeSalt fixed-salt -OutputPath $T/out-s2 >/dev/null 2>&1
$PWSH -NoProfile -File $SCRIPT -SampleData -Anonymize -AnonymizeSalt other-salt -OutputPath $T/out-s3 >/dev/null 2>&1
if diff -q $T/out-s1/docs/03-roles.md $T/out-s2/docs/03-roles.md >/dev/null; then
  ok "same salt renders identical pseudonyms"
else bad "same salt" "docs differ between two runs"; fi
if diff -q $T/out-s1/docs/03-roles.md $T/out-s3/docs/03-roles.md >/dev/null; then
  bad "different salt" "a different salt produced identical output"
else ok "different salt renders different pseudonyms"; fi
D1=$(grep -c "Person " $T/out-s1/docs/03-roles.md)
D3=$(grep -c "Person " $T/out-s3/docs/03-roles.md)
check "salt change does not change how many people are listed" "$D1" "$D3"

echo
echo "--- 3. structural parity vs the clear render -------------------------"
rm -rf $T/out-clear
$PWSH -NoProfile -File $SCRIPT -SampleData -OutputPath $T/out-clear >/dev/null 2>&1
for f in index 01-tenant 02-conditional-access 03-roles 04-groups 05-authentication 06-user-settings 07-applications 08-intune 09-changelog; do
  A=$(wc -l < $T/out-clear/docs/$f.md); B=$(wc -l < $T/out-anon/docs/$f.md)
  # index.md gains the anonymized banner (2 lines); everything else must match.
  EXP=$A; [ "$f" = "index" ] && EXP=$((A+2))
  if [ "$B" = "$EXP" ]; then ok "docs/$f.md line count unchanged ($B)"
  else bad "docs/$f.md" "clear=$A anon=$B expected=$EXP"; fi
done

echo
echo "--- 4. derived data unaffected ---------------------------------------"
CC=$(python3 -c "import json;d=json.load(open('$T/out-clear/run-summary.json'));print(len(d['NewChanges']),d['SnapshotCount'],d['Kpis']['Members'],d['Kpis']['CaEnabled'],d['Kpis']['CaGapsCritical'],d['Kpis']['CredsInWindow'])")
AC=$(python3 -c "import json;d=json.load(open('$T/out-anon/run-summary.json'));print(len(d['NewChanges']),d['SnapshotCount'],d['Kpis']['Members'],d['Kpis']['CaEnabled'],d['Kpis']['CaGapsCritical'],d['Kpis']['CredsInWindow'])")
check "run-summary KPIs and change count identical" "$AC" "$CC"
CG=$(grep -c "^| " $T/out-clear/docs/02-conditional-access.md)
AG=$(grep -c "^| " $T/out-anon/docs/02-conditional-access.md)
check "conditional access table rows identical" "$AG" "$CG"
CL=$(grep -c "^- \*\*" $T/out-clear/docs/09-changelog.md)
AL=$(grep -c "^- \*\*" $T/out-anon/docs/09-changelog.md)
check "change log event count identical (mapping is consistent across history)" "$AL" "$CL"

echo
echo "--- 5. history on disk is never rewritten ----------------------------"
BEFORE=$(find $REPO/sample-history -name '*.json' -exec md5sum {} \; | sort | md5sum)
$PWSH -NoProfile -File $SCRIPT -SampleData -Anonymize -OutputPath $T/out-hist >/dev/null 2>&1
AFTER=$(find $REPO/sample-history -name '*.json' -exec md5sum {} \; | sort | md5sum)
check "sample-history files byte-identical after an anonymized run" "$AFTER" "$BEFORE"

echo
echo "--- 6. mock live run: the archive keeps REAL values ------------------"
# The archive-then-anonymize ordering only exists on the live collection path,
# so stub Get-TenantData on a copy and exercise it for real.
sed 's/^function Get-TenantData {/function Get-TenantDataReal {/' $SCRIPT > $T/mock.ps1
cat > $T/stub.ps1 <<'STUB'
function Get-TenantData {
    param([int]$StaleCredDays, [switch]$SkipIntune)
    $d = Get-Content "$env:REPO/sample-data.json" -Raw | ConvertFrom-Json -AsHashtable
    $d['GeneratedUtc'] = (Get-Date).ToUniversalTime().ToString('o')
    return $d
}
STUB
python3 - <<'PY'
import os
T = os.environ['T']
stub = open(T + '/stub.ps1').read()
src  = open(T + '/mock.ps1').read()
marker = "# --------------------------------------------------------------------------- #\n# Main\n"
assert marker in src, "main marker not found"
open(T + '/mock.ps1','w').write(src.replace(marker, stub + "\n" + marker, 1))
PY
rm -rf $T/out-live $T/hist-live
$PWSH -NoProfile -File $T/mock.ps1 -Anonymize -OutputPath $T/out-live -HistoryPath $T/hist-live >/dev/null 2>&1
ARCHIVED=$(ls $T/hist-live/*.json 2>/dev/null | head -1)
if [ -z "$ARCHIVED" ]; then bad "mock live run archived a snapshot" "no snapshot written to $T/hist-live"
else
  ok "mock live run archived a snapshot"
  if grep -q "northwindtraders.com" "$ARCHIVED"; then ok "archived snapshot keeps REAL values (your history stays usable)"
  else bad "archived snapshot" "real values are missing - the archive was anonymized"; fi
  if grep -q "northwindtraders.com" $T/out-live/tenant.json; then bad "rendered tenant.json" "real domain leaked into the rendered output"
  else ok "rendered tenant.json in the same run is anonymized"; fi
fi

echo
echo "--- 7. degradation on thin snapshots ---------------------------------"
python3 - <<'PY'
import json, os
d = json.load(open(os.environ['REPO'] + "/sample-data.json"))
d['Intune'] = {'Available': False}
d['AuthMethods'] = []
d['Roles'] = []
d['Groups']['Dynamic'] = []
d['Groups']['RoleAssignable'] = []
d['Applications'] = []
d['ConditionalAccess']['NamedLocations'] = []
json.dump(d, open(os.environ['T'] + '/thin.json','w'))
PY
rm -rf $T/out-thin
if $PWSH -NoProfile -File $SCRIPT -FromJson $T/thin.json -Anonymize -NoHistory -OutputPath $T/out-thin > $T/thin.log 2>&1; then
  ok "thin snapshot (no Intune, no auth methods, no roles, no apps) renders without error"
  [ -f $T/out-thin/report.html ] && ok "thin snapshot still produced report.html" || bad "thin report" "missing"
else bad "thin snapshot" "$(tail -5 $T/thin.log)"; fi

echo
echo "--- 8. the output says it is anonymized ------------------------------"
grep -q "Anonymized" $T/out-anon/docs/index.md && ok "docs/index.md carries the anonymized banner" \
  || bad "docs banner" "not found in index.md"
grep -q "t-anon" $T/out-anon/report.html && ok "report.html carries the anonymized banner" \
  || bad "html banner" "not found"
python3 -c "
import json,sys
d=json.load(open('$T/out-anon/tenant.json'))
sys.exit(0 if d.get('Anonymized') is True else 1)" && ok "tenant.json is flagged Anonymized" \
  || bad "tenant.json flag" "Anonymized not true"
python3 -c "
import json,sys
d=json.load(open('$T/out-clear/tenant.json'))
sys.exit(0 if not d.get('Anonymized') else 1)" && ok "a clear render is NOT flagged anonymized" \
  || bad "clear flag" "clear render claims to be anonymized"

echo
echo "--- 9. the HTML report actually executes -----------------------------"
# The docs are rendered by PowerShell; report.html is rendered by JavaScript
# from an embedded payload, so only running the page proves it works.
if python3 $HERE/htmlcheck.py $T/out-clear/report.html $T/out-anon/report.html $T/out-thin/report.html > $T/html.txt 2>&1; then
  while read -r line; do ok "${line#PASS }"; done < <(grep '^PASS' $T/html.txt)
else
  bad "html render" "$(cat $T/html.txt)"
fi

echo
echo "======================================================================"
echo "PASS $PASS   FAIL $FAIL"
[ $FAIL -eq 0 ] || exit 1
