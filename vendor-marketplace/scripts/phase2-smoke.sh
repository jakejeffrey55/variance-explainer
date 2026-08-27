#!/usr/bin/env bash
# Phase 2 smoke test: property CRUD, CSV import (incl. stubbed providers),
# job CRUD, status transitions, and admin-only enforcement.
# Requires a running dev server and a seeded database.
set -u
BASE="${BASE:-http://localhost:3000}"
DIR=$(mktemp -d)
JAR="$DIR/admin.jar"
VJAR="$DIR/vendor.jar"
pass=0; fail=0

check() {
  if [ "$2" = "$3" ]; then echo "  PASS  $1 ($3)"; pass=$((pass+1));
  else echo "  FAIL  $1 (expected $2, got $3)"; fail=$((fail+1)); fi
}
contains() {
  if printf '%s' "$3" | grep -q -- "$2"; then echo "  PASS  $1"; pass=$((pass+1));
  else echo "  FAIL  $1 (missing '$2' in: $(printf '%s' "$3" | head -c 200))"; fail=$((fail+1)); fi
}

login() {
  local surface=$1 email=$2 pw=$3 jar=$4
  rm -f "$jar"
  local csrf
  csrf=$(curl -s -c "$jar" "$BASE/api/auth/$surface/csrf" | sed -E 's/.*"csrfToken":"([^"]*)".*/\1/')
  curl -s -o /dev/null -b "$jar" -c "$jar" -d "csrfToken=$csrf" -d "email=$email" -d "password=$pw" \
    -d "json=true" -d "callbackUrl=$BASE" "$BASE/api/auth/$surface/callback/${surface}-credentials"
}

login admin dana.reyes@cortland-example.com 'Admin123!' "$JAR"
login vendor rick@lonestarmakeready.example.com 'Vendor123!' "$VJAR"

echo "== admin-only enforcement =="
check "vendor session -> admin properties API 401" 401 \
  "$(curl -s -o /dev/null -w '%{http_code}' -b "$VJAR" "$BASE/api/admin/properties")"
check "no session -> admin jobs API 401" 401 \
  "$(curl -s -o /dev/null -w '%{http_code}' "$BASE/api/admin/jobs")"
check "vendor session -> job create 401" 401 \
  "$(curl -s -o /dev/null -w '%{http_code}' -X POST -b "$VJAR" -H 'Content-Type: application/json' \
    -d '{"propertyId":"x","title":"hack attempt","description":"should never work","category":"MAKE_READY","budgetMin":1,"budgetMax":2}' \
    "$BASE/api/admin/jobs")"

echo "== properties =="
LIST=$(curl -s -b "$JAR" "$BASE/api/admin/properties")
contains "property list returns seeded properties" "Cortland Uptown Dallas" "$LIST"

CREATED=$(curl -s -b "$JAR" -X POST -H 'Content-Type: application/json' \
  -d '{"name":"Smoke Test Commons","addressLine1":"1 Test Way","city":"Irving","state":"TX","postalCode":"75039","latitude":32.8887,"longitude":-96.9489,"unitCount":100,"isActive":true}' \
  "$BASE/api/admin/properties")
PROP_ID=$(printf '%s' "$CREATED" | grep -o '"id":"[^"]*"' | head -1 | cut -d'"' -f4)
contains "manual property created" "Smoke Test Commons" "$CREATED"

check "property update 200" 200 "$(curl -s -o /dev/null -w '%{http_code}' -b "$JAR" -X PATCH \
  -H 'Content-Type: application/json' -d '{"unitCount":120}' "$BASE/api/admin/properties/$PROP_ID")"

BAD=$(curl -s -b "$JAR" -X POST -H 'Content-Type: application/json' \
  -d '{"name":"x","addressLine1":"1 Test Way","city":"Irving","state":"TX","postalCode":"75039","latitude":999,"longitude":0}' \
  "$BASE/api/admin/properties")
contains "invalid coordinates rejected" "invalid_request" "$BAD"

echo "== CSV import (active provider) =="
# Unique ids per run so the assertions hold on a database that has already been
# smoke-tested (create vs. update vs. unchanged all have to be distinguishable).
RUN=$(date +%s)
RIDGE="Smoke Ridge $RUN"
cat > "$DIR/props.csv" <<CSV
property_id,name,address,city,state,zip,latitude,longitude,units,property_manager,manager_email,active
SMOKE-$RUN-1,$RIDGE,900 Test Blvd,Grapevine,TX,76051,32.9343,-97.0781,180,Pat Lane,pat.lane@example.com,true
SMOKE-$RUN-2,Smoke Creek $RUN,50 Sample Dr,Denton,TX,76201,33.2148,-97.1331,96,Sam Ford,sam.ford@example.com,true
CRT-DAL-0148,Cortland Uptown Dallas,2801 Cedar Springs Rd,Dallas,TX,75201,32.7997,-96.8065,$((300 + RANDOM % 99)),Alicia Gomez,alicia.gomez@cortland-example.com,true
SMOKE-$RUN-BAD,Missing Coords,12 Nowhere,Plano,TX,75024,,,50,,,true
CSV

PREVIEW=$(curl -s -b "$JAR" -F "providerKey=csv" -F "dryRun=true" -F "file=@$DIR/props.csv" \
  "$BASE/api/admin/properties/import")
contains "dry run reports creates" '"created":2' "$PREVIEW"
contains "dry run detects an update to an existing property" '"updated":1' "$PREVIEW"
contains "dry run flags the row missing coordinates" 'Latitude and longitude are required' "$PREVIEW"
contains "dry run does not write" '"dryRun":true' "$PREVIEW"

BEFORE=$(curl -s -b "$JAR" "$BASE/api/admin/properties" | grep -o "$RIDGE" | wc -l | tr -d ' ')
check "dry run really wrote nothing" 0 "$BEFORE"

IMPORTED=$(curl -s -b "$JAR" -F "providerKey=csv" -F "dryRun=false" -F "file=@$DIR/props.csv" \
  "$BASE/api/admin/properties/import")
contains "import created 2 properties" '"created":2' "$IMPORTED"
AFTER=$(curl -s -b "$JAR" "$BASE/api/admin/properties" | grep -o "$RIDGE" | wc -l | tr -d ' ')
check "imported property is readable" 1 "$AFTER"

# Re-importing the same file must be a no-op, not a duplicate.
REIMPORT=$(curl -s -b "$JAR" -F "providerKey=csv" -F "dryRun=false" -F "file=@$DIR/props.csv" \
  "$BASE/api/admin/properties/import")
contains "re-import creates nothing" '"created":0' "$REIMPORT"
contains "re-import reports rows unchanged" '"unchanged":3' "$REIMPORT"
STILL=$(curl -s -b "$JAR" "$BASE/api/admin/properties" | grep -o "$RIDGE" | wc -l | tr -d ' ')
check "no duplicate row created" 1 "$STILL"

cat > "$DIR/badheaders.csv" <<'CSV'
foo,bar
1,2
CSV
MISSING=$(curl -s -b "$JAR" -F "providerKey=csv" -F "dryRun=true" -F "file=@$DIR/badheaders.csv" \
  "$BASE/api/admin/properties/import")
contains "missing required columns reported by name" 'missing required column' "$MISSING"

echo "== stubbed providers =="
for provider in powerbi onesite; do
  CODE=$(curl -s -o "$DIR/$provider.json" -w '%{http_code}' -b "$JAR" -F "providerKey=$provider" -F "dryRun=true" \
    "$BASE/api/admin/properties/import")
  check "$provider returns 501 not-configured" 501 "$CODE"
  contains "$provider explains the remedy" 'not configured' "$(cat "$DIR/$provider.json")"
done
PROVIDERS=$(curl -s -b "$JAR" "$BASE/api/admin/providers")
contains "provider registry lists all three" 'onesite' "$PROVIDERS"

echo "== jobs =="
JOB=$(curl -s -b "$JAR" -X POST -H 'Content-Type: application/json' -d "{
  \"propertyId\":\"$PROP_ID\",
  \"title\":\"Smoke test make-ready\",
  \"description\":\"Paint, clean, and punch the unit for the smoke test.\",
  \"category\":\"MAKE_READY\",
  \"budgetMin\":1200,\"budgetMax\":2400,\"enforceBudgetCap\":true,
  \"bidDeadline\":\"$(date -u -d '+5 days' '+%Y-%m-%dT%H:%M:00Z')\",
  \"status\":\"DRAFT\"}" "$BASE/api/admin/jobs")
JOB_ID=$(printf '%s' "$JOB" | grep -o '"id":"[^"]*"' | head -1 | cut -d'"' -f4)
contains "job created as draft" '"status":"DRAFT"' "$JOB"
contains "job number allocated" '"jobNumber":"J-' "$JOB"

BADBUDGET=$(curl -s -b "$JAR" -X POST -H 'Content-Type: application/json' -d "{
  \"propertyId\":\"$PROP_ID\",\"title\":\"Bad budget job\",\"description\":\"Maximum below minimum should fail.\",
  \"category\":\"PLUMBING\",\"budgetMin\":5000,\"budgetMax\":100}" "$BASE/api/admin/jobs")
contains "budget max below min rejected" 'at least the minimum' "$BADBUDGET"

PASTDEADLINE=$(curl -s -b "$JAR" -X POST -H 'Content-Type: application/json' -d "{
  \"propertyId\":\"$PROP_ID\",\"title\":\"Past deadline job\",\"description\":\"Publishing with a past deadline should fail.\",
  \"category\":\"PLUMBING\",\"budgetMin\":100,\"budgetMax\":500,
  \"bidDeadline\":\"$(date -u -d '-2 days' '+%Y-%m-%dT%H:%M:00Z')\",\"status\":\"OPEN\"}" "$BASE/api/admin/jobs")
contains "past bid deadline rejected on publish" 'already passed' "$PASTDEADLINE"

EMG=$(curl -s -b "$JAR" -X POST -H 'Content-Type: application/json' -d "{
  \"propertyId\":\"$PROP_ID\",\"title\":\"Smoke emergency — no cooling\",
  \"description\":\"Resident reports no cooling; dispatch immediately.\",
  \"category\":\"HVAC\",\"budgetMin\":250,\"budgetMax\":1500,
  \"priority\":\"EMERGENCY\",\"emergencyCategory\":\"AC_HVAC\",\"responseDeadlineMinutes\":15}" "$BASE/api/admin/jobs")
contains "emergency job opens immediately" '"status":"OPEN"' "$EMG"
contains "emergency job is dispatched on creation" '"dispatchedAt":"' "$EMG"

EMGBID=$(curl -s -b "$JAR" -X POST -H 'Content-Type: application/json' -d "{
  \"propertyId\":\"$PROP_ID\",\"title\":\"Emergency with a bid deadline\",
  \"description\":\"An emergency job must not carry a bid deadline.\",
  \"category\":\"HVAC\",\"budgetMin\":250,\"budgetMax\":1500,\"priority\":\"EMERGENCY\",
  \"emergencyCategory\":\"LEAK\",\"bidDeadline\":\"$(date -u -d '+2 days' '+%Y-%m-%dT%H:%M:00Z')\"}" "$BASE/api/admin/jobs")
contains "emergency + bid deadline rejected" 'skip bidding' "$EMGBID"

echo "== job transitions =="
EMG_ID=$(printf '%s' "$EMG" | grep -o '"id":"[^"]*"' | head -1 | cut -d'"' -f4)
EMGCLOSE=$(curl -s -b "$JAR" -X POST -H 'Content-Type: application/json' -d '{"action":"close_bidding"}' \
  "$BASE/api/admin/jobs/$EMG_ID/status")
contains "closing bidding on an emergency rejected" 'claimed rather than bid on' "$EMGCLOSE"

check "publish draft 200" 200 "$(curl -s -o /dev/null -w '%{http_code}' -b "$JAR" -X POST \
  -H 'Content-Type: application/json' -d '{"action":"publish"}' "$BASE/api/admin/jobs/$JOB_ID/status")"
INVALID=$(curl -s -b "$JAR" -X POST -H 'Content-Type: application/json' -d '{"action":"publish"}' \
  "$BASE/api/admin/jobs/$JOB_ID/status")
contains "re-publishing an open job rejected" 'invalid_transition' "$INVALID"
check "close bidding 200" 200 "$(curl -s -o /dev/null -w '%{http_code}' -b "$JAR" -X POST \
  -H 'Content-Type: application/json' -d '{"action":"close_bidding"}' "$BASE/api/admin/jobs/$JOB_ID/status")"

echo "== activity log =="
DETAIL=$(curl -s -b "$JAR" "$BASE/api/admin/jobs/$JOB_ID")
contains "job.created logged" 'job.created' "$DETAIL"
contains "job.publish logged" 'job.publish' "$DETAIL"
contains "job.close_bidding logged" 'job.close_bidding' "$DETAIL"

echo "== pages render =="
for path in /admin /admin/jobs /admin/properties /admin/import /admin/jobs/new "/admin/jobs/$JOB_ID" "/admin/properties/$PROP_ID" /vendor; do
  jar="$JAR"; [ "$path" = "/vendor" ] && jar="$VJAR"
  check "GET $path" 200 "$(curl -s -o /dev/null -w '%{http_code}' -b "$jar" "$BASE$path")"
done

echo
echo "passed=$pass failed=$fail"
[ "$fail" -eq 0 ]
