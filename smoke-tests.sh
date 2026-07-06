#!/usr/bin/env bash
# smoke-tests.sh — Smoke tests en vivo de los ejemplos contra el tenant book-test.
# Corre el main de cada ejemplo (operación de solo lectura) y registra PASS/FAIL.
# Requiere credenciales en /home/nacho/code/books/code/dataverse-validation/.env
# (bajo prefijo DATAVERSE_*). Uso:  bash smoke-tests.sh
set -u
REPO=/home/nacho/code/books/la-experiencia-sharepoint-code
ENV_FILE=/home/nacho/code/books/code/dataverse-validation/.env
PYVENV=/home/nacho/.lesp-venv/bin/python

if [ ! -f "$ENV_FILE" ]; then echo "Falta $ENV_FILE"; exit 2; fi
set -a; . "$ENV_FILE"; set +a
export TENANT_ID="$DATAVERSE_TENANT_ID" CLIENT_ID="$DATAVERSE_CLIENT_ID" CLIENT_SECRET="$DATAVERSE_CLIENT_SECRET"
export SHAREPOINT_HOSTNAME="olddogsoft1.sharepoint.com" SHAREPOINT_SITE_PATH="book-test"

cd "$REPO"
PASS=0; FAIL=0; RESULTS=""

run_smoke() {
  local label="$1"; shift
  local out err rc
  out=$(timeout 90 "$@" 2>&1); rc=$?
  # Considerar PASS si exit 0; o si la salida indica lectura exitosa y no error fatal.
  if [ $rc -eq 0 ]; then
    PASS=$((PASS+1)); RESULTS+=$(printf "PASS | %s\n" "$label")
  else
    FAIL=$((FAIL+1))
    local errline=$(echo "$out" | grep -iE "error|❌|exception" | head -1)
    RESULTS+=$(printf "FAIL | %s | rc=%s | %s\n" "$label" "$rc" "$errline")
  fi
}

# --- C# (dotnet run) ---
run_smoke "C#  cap3 sites"       dotnet run --project chapter-03-sites/csharp/Chapter03Sites.csproj --no-build -v q
run_smoke "C#  cap4 documents"   dotnet run --project chapter-04-documents/csharp/Chapter04Documents.csproj --no-build -v q
run_smoke "C#  cap5 permissions" dotnet run --project chapter-05-permissions/csharp/Chapter05Permissions.csproj --no-build -v q
run_smoke "C#  cap6 metadata"    dotnet run --project chapter-06-metadata/csharp/Chapter06Metadata.csproj --no-build -v q
run_smoke "C#  cap7 automation"  dotnet run --project chapter-07-automation/csharp/Chapter07Automation.csproj --no-build -v q

# --- Python ---
run_smoke "PY  cap3 sites"       "$PYVENV" chapter-03-sites/python/site_operations.py
run_smoke "PY  cap4 documents"   "$PYVENV" chapter-04-documents/python/document_operations.py
run_smoke "PY  cap5 permissions" "$PYVENV" chapter-05-permissions/python/permission_operations.py
run_smoke "PY  cap6 metadata"     "$PYVENV" chapter-06-metadata/python/01-termstore/explore_term_store.py
run_smoke "PY  cap7 automation"   "$PYVENV" chapter-07-automation/python/02-delta-queries/delta_queries.py

# --- JavaScript ---
run_smoke "JS  cap3 sites"       node chapter-03-sites/javascript/SiteOperations.js
run_smoke "JS  cap4 documents"   node chapter-04-documents/javascript/DocumentOperations.js
run_smoke "JS  cap5 permissions" node chapter-05-permissions/javascript/PermissionOperations.js
run_smoke "JS  cap6 metadata"     node chapter-06-metadata/javascript/01-termstore/exploreTermStore.js
run_smoke "JS  cap7 automation"   node chapter-07-automation/javascript/02-delta-queries/deltaQueries.js
run_smoke "JS  cap7 webhooks"     node chapter-07-automation/javascript/01-webhooks/webhookManager.js

# --- PowerShell (caps 3/4/5 mains reales; cap6 interactivo PnP y cap7 parametrizado, fuera del smoke headless) ---
run_smoke "PS  cap3 sites"       pwsh -NoProfile -File chapter-03-sites/powershell/SiteOperations.ps1
run_smoke "PS  cap4 documents"   pwsh -NoProfile -File chapter-04-documents/powershell/DocumentOperations.ps1
run_smoke "PS  cap5 permissions" pwsh -NoProfile -File chapter-05-permissions/powershell/PermissionOperations.ps1

# --- Java (mains; requieren reactor instalado en .m2) ---
mvn -q install -DskipTests >/dev/null 2>&1
run_smoke "JV  cap3 sites"        mvn -q -pl chapter-03-sites/java exec:java -Dexec.mainClass=com.sharepointexperience.chapter03.SiteOperations
run_smoke "JV  cap4 documents"    mvn -q -pl chapter-04-documents/java exec:java -Dexec.mainClass=com.sharepointexperience.chapter04.DocumentOperations
run_smoke "JV  cap5 permissions"  mvn -q -pl chapter-05-permissions/java exec:java -Dexec.mainClass=com.sharepointexperience.chapter05.PermissionOperations
run_smoke "JV  cap6 metadata"     mvn -q -pl chapter-06-metadata/java exec:java -Dexec.mainClass=com.sharepointexperiencia.managedmetadata.TermStoreExplorer
run_smoke "JV  cap7 automation"   mvn -q -pl chapter-07-automation/java exec:java -Dexec.mainClass=com.sharepointexperiencia.automation.WebhookManager

echo "==================== SMOKE TESTS (en vivo, tenant book-test) ===================="
echo "$RESULTS"
echo "--------------------------------------------------------------------------------"
echo "Resumen: PASS=$PASS  FAIL=$FAIL"