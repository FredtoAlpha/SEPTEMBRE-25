#!/bin/bash

# This script checks for duplicate function definitions in groupsBackend.js

function check_duplicates {
  local function_name=$1
  local count=$(grep -c "function $function_name()" groupsBackend.js)

  if [ "$count" -gt 1 ]; then
    echo "Error: Found $count definitions for function '$function_name' in groupsBackend.js"
    exit 1
  elif [ "$count" -eq 0 ]; then
    echo "Error: Found no definitions for function '$function_name' in groupsBackend.js"
    exit 1
  else
    echo "Found 1 definition for function '$function_name' in groupsBackend.js (OK)"
  fi
}

check_duplicates "getStudentsForGroups"
check_duplicates "getINTClasses"

echo "No duplicate functions found."
exit 0
