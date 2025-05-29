#!/bin/bash

# Script to check for and remove Microsoft Office on macOS

office_apps=(
  "/Applications/Microsoft Word.app"
  "/Applications/Microsoft Excel.app"
  "/Applications/Microsoft PowerPoint.app"
  "/Applications/Microsoft Outlook.app"
  "/Applications/Microsoft OneNote.app"
  "/Applications/Microsoft Teams.app"
  "/Applications/OneDrive.app"
)

## Added LauchAgents/LaunchDaemons, licensing files and installation receipts
office_library_paths=(
  "~/Library/Containers/com.microsoft.*"
  "~/Library/Group Containers/UBF8T346G9.Office"
  "~/Library/Application Scripts/com.microsoft.*"
  "~/Library/Preferences/com.microsoft.*.plist"
  "~/Library/Saved Application State/com.microsoft.*.savedState"
  "~/Library/Caches/com.microsoft.*"
  "/Library/Application Support/Microsoft"
  "~/Library/Logs/Microsoft"
  "~/Library/Cookies/com.microsoft.*.binarycookies"
  "/Library/Caches/com.microsoft*"
  "/Library/LaunchAgents/com.microsoft*"
  "/Library/LaunchDaemons/com.microsoft*"
  "/Library/Microsoft"
  "/Library/PrivilegedHelperTools/com.microsoft*"
  "/Library/Preferences/com.microsoft*"
  "/private/var/db/receipts/com.microsoft*"
)

found_apps=()
removed_app_count=0
removed_file_count=0
error_count=0

echo "Checking for Microsoft Office..."

for app in "${office_apps[@]}"; do
  if [[ -d "$app" ]]; then
    echo "Found: $app"
    found_apps+=("$app")
  fi
done

if [[ ${#found_apps[@]} -gt 0 ]]; then
  echo ""
  read -p "Microsoft Office applications found. Proceed with removal? [y/n]: " -r response

  ## Let's expand the valid responses to also include Yes (case insensitive)
  if [[ "$response" =~ ^[Yy]$ || "$response" =~ ^[Yy][Ee][Ss]$ ]]; then
    echo "Removing applications..."
    for app in "${found_apps[@]}"; do

      ## Removed the -f option as including it with rm will always return 0 (which will break your error check)
      sudo rm -r "$app" >/dev/null 2>&1 # Suppress output

      if [[ $? -eq 0 ]]; then
        echo "Removed: $app"
        ((removed_app_count++))
      else
        echo "Error removing: $app"
        ((error_count++))
      fi
    done

    echo ""
    echo "Removing associated files and folders..."
    for path_pattern in "${office_library_paths[@]}"; do

      ## Unquote $path_pattern here so the find command interprets the * as a regex and not literally
      find $path_pattern -maxdepth 3 -print0 2>/dev/null | while IFS= read -r -d $'\0' item; do

        ## Removed the -f option as including it with rm will always return 0 (which will break your error check)
        sudo rm -r "$item" >/dev/null 2>&1 # Suppress output

        if [[ $? -eq 0 ]]; then
          ((removed_file_count++))
        else
          ((error_count++))
        fi
      done

    done

    echo ""
    echo "Removal Summary:"
    echo "  Applications removed: $removed_app_count"
    echo "  Associated files/folders removed: $removed_file_count"
    if [[ $error_count -gt 0 ]]; then
      echo "  Errors encountered: $error_count"

      ## Let's add messaging (with color [red]) and an exit with a value other than 0 here to indicate an error
      echo "$(tput bold; tput setaf 1)Microsoft Office removal process completed with errors.$(tput sgr 0)"
      exit 1
    fi
    ## Just a little color [green] to the success message
    echo "$(tput bold; tput setaf 2)Microsoft Office removal process completed.$(tput sgr 0)"
  else
    echo "Removal cancelled."
  fi
else
  echo "Microsoft Office not found."
fi

exit 0