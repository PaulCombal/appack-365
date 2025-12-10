#!/usr/bin/env bash

# Original array
apps=("EXCEL.EXE" "MSACCESS.EXE" "MSPUB.EXE" "ONENOTE.EXE" "OUTLOOK.EXE" "POWERPNT.EXE" "SETLANG.EXE" "WINWORD.EXE")

# Build array of associative arrays, like your example
declare -A app_objects
index=0

for exe in "${apps[@]}"; do
    name=$(echo "$exe" | awk -F. '{print tolower($1)}')   # EXCEL.EXE → excel

    app_objects["$index,name"]="$name"
    app_objects["$index,exe"]="$exe"
    ((index++))
done

# Ensure output directory exists
mkdir -p ./C

# Process each app
for i in $(seq 0 $((index-1))); do
    name="${app_objects[$i,name]}"
    exe="${app_objects[$i,exe]}"

    out="./C/${name}.vbs"

    # Copy template and replace the marker
    sed "s/\$APP_EXE_NAME/$exe/g" ./template.vbs > "$out"
done
