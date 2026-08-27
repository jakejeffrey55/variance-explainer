#!/usr/bin/env bash
# Phase 1 auth enforcement checks. Every assertion is server-side.
set -u
BASE="http://localhost:3000"
DIR=$(mktemp -d)
pass=0; fail=0
check() { # check <label> <expected> <actual>
  if [ "$2" = "$3" ]; then echo "  PASS  $1 ($3)"; pass=$((pass+1));
  else echo "  FAIL  $1 (expected $2, got $3)"; fail=$((fail+1)); fi
}

login() { # login <surface> <email> <password> <jar>
  local surface=$1 email=$2 pass=$3 jar=$4
  rm -f "$jar"
  local csrf
  csrf=$(curl -s -c "$jar" "$BASE/api/auth/$surface/csrf" | sed -E 's/.*"csrfToken":"([^"]*)".*/\1/')
  curl -s -o /dev/null -b "$jar" -c "$jar" \
    -d "csrfToken=$csrf" -d "email=$email" -d "password=$pass" -d "json=true" \
    -d "callbackUrl=$BASE" \
    "$BASE/api/auth/$surface/callback/${surface}-credentials"
  grep -c "vm.$surface.session-token" "$jar" 2>/dev/null | tr -d ' '
}

code() { curl -s -o /dev/null -w "%{http_code}" -b "$1" "$BASE$2"; }
body() { curl -s -b "$1" "$BASE$2"; }

echo "== unauthenticated =="
check "GET /api/admin/me  -> 401" 401 "$(curl -s -o /dev/null -w '%{http_code}' $BASE/api/admin/me)"
check "GET /api/vendor/me -> 401" 401 "$(curl -s -o /dev/null -w '%{http_code}' $BASE/api/vendor/me)"

echo "== admin surface =="
check "admin login issues admin cookie" 1 "$(login admin dana.reyes@cortland-example.com 'Admin123!' $DIR/admin.jar)"
check "admin -> /api/admin/me  200" 200 "$(code $DIR/admin.jar /api/admin/me)"
check "admin -> /api/vendor/me 401 (cross-scope)" 401 "$(code $DIR/admin.jar /api/vendor/me)"
check "admin bad password rejected" 0 "$(login admin dana.reyes@cortland-example.com 'wrong' $DIR/bad.jar)"

echo "== vendor surface =="
check "vendor login issues vendor cookie" 1 "$(login vendor rick@lonestarmakeready.example.com 'Vendor123!' $DIR/vendor.jar)"
check "vendor -> /api/vendor/me 200" 200 "$(code $DIR/vendor.jar /api/vendor/me)"
check "vendor -> /api/admin/me  401 (cross-scope)" 401 "$(code $DIR/vendor.jar /api/admin/me)"

echo "== separate credential spaces =="
check "admin creds on vendor surface rejected" 0 "$(login vendor dana.reyes@cortland-example.com 'Admin123!' $DIR/x1.jar)"
check "vendor creds on admin surface rejected" 0 "$(login admin rick@lonestarmakeready.example.com 'Vendor123!' $DIR/x2.jar)"

echo "== account gates =="
check "suspended vendor cannot sign in" 0 "$(login vendor neil@gulfcoasthandyman.example.com 'Vendor123!' $DIR/susp.jar)"
check "pending vendor can sign in" 1 "$(login vendor ava@pinnacleprops.example.com 'Vendor123!' $DIR/pending.jar)"
echo "  pending vendor session: $(body $DIR/pending.jar /api/vendor/me)"
echo "  active vendor session:  $(body $DIR/vendor.jar /api/vendor/me)"

echo "== middleware redirects =="
check "/admin/dashboard -> login redirect" 307 "$(curl -s -o /dev/null -w '%{http_code}' $BASE/admin/dashboard)"
check "/vendor/jobs -> login redirect" 307 "$(curl -s -o /dev/null -w '%{http_code}' $BASE/vendor/jobs)"
check "vendor cookie on /admin/* still redirected" 307 "$(curl -s -o /dev/null -w '%{http_code}' -b $DIR/vendor.jar $BASE/admin/dashboard)"

echo
echo "passed=$pass failed=$fail"
[ "$fail" -eq 0 ]
