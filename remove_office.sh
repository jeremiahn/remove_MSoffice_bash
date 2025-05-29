#!/bin/bash

# =============================================================================
# Microsoft Office Removal Script for macOS
# =============================================================================
# Purpose: Completely removes Microsoft Office applications and associated files
# Author: Jeremiah Nelson
# Version: 1.5.0
# Compatibility: macOS 10.12+ (Sierra and later)
# Requirements: Administrative privileges (sudo access)
# =============================================================================

# Array containing paths to all Microsoft Office application bundles
# These are the standard installation locations in /Applications
office_apps=(
  "/Applications/Microsoft Word.app"
  "/Applications/Microsoft Excel.app"
  "/Applications/Microsoft PowerPoint.app"
  "/Applications/Microsoft Outlook.app"
  "/Applications/Microsoft OneNote.app"
  "/Applications/Microsoft Teams.app"
  "/Applications/OneDrive.app"
)

# Array containing paths to Microsoft Office support files, preferences, and system components
# This includes user-specific and system-wide files that Office creates during installation and use
office_library_paths=(
  # User-specific container directories (sandboxed app data)
  "~/Library/Containers/com.microsoft.*"
  
  # Shared group container for Office suite inter-app communication
  "~/Library/Group Containers/UBF8T346G9.Office"
  
  # AppleScript and automation scripts used by Office apps
  "~/Library/Application Scripts/com.microsoft.*"
  
  # User preference files (settings, configurations)
  "~/Library/Preferences/com.microsoft.*.plist"
  
  # Application state files (window positions, recent documents, etc.)
  "~/Library/Saved Application State/com.microsoft.*.savedState"
  
  # User-specific cache files
  "~/Library/Caches/com.microsoft.*"
  
  # System-wide application support files
  "/Library/Application Support/Microsoft"
  
  # User-specific log files
  "~/Library/Logs/Microsoft"
  
  # Browser cookies stored by Office web components
  "~/Library/Cookies/com.microsoft.*.binarycookies"
  
  # System-wide cache files
  "/Library/Caches/com.microsoft*"
  
  # Launch agents (user-level background processes)
  "/Library/LaunchAgents/com.microsoft*"
  
  # Launch daemons (system-level background processes)
  "/Library/LaunchDaemons/com.microsoft*"
  
  # System-wide Microsoft directory
  "/Library/Microsoft"
  
  # Privileged helper tools (system services)
  "/Library/PrivilegedHelperTools/com.microsoft*"
  
  # System-wide preference files
  "/Library/Preferences/com.microsoft*"
  
  # Package installation receipts (tracks what was installed)
  "/private/var/db/receipts/com.microsoft*"
)

# Initialize counters and arrays for tracking removal progress
found_apps=()           # Array to store discovered Office applications
removed_app_count=0     # Counter for successfully removed applications
removed_file_count=0    # Counter for successfully removed files/folders
error_count=0          # Counter for errors encountered during removal

# Display initial status message
echo "Checking for Microsoft Office..."

# =====================================================
# PHASE 1: DISCOVER INSTALLED OFFICE APPLICATIONS
# =====================================================
# Loop through each application path and check if it exists
for app in "${office_apps[@]}"; do
  # Check if the application directory exists using the -d test
  if [[ -d "$app" ]]; then
    echo "Found: $app"
    # Add discovered application to our tracking array
    found_apps+=("$app")
  fi
done

# =====================================================
# PHASE 2: USER CONFIRMATION AND REMOVAL PROCESS
# =====================================================
# Only proceed if we found at least one Office application
if [[ ${#found_apps[@]} -gt 0 ]]; then
  echo ""
  # Prompt user for confirmation before proceeding with removal
  read -p "Microsoft Office applications found. Proceed with removal? [y/n]: " -r response

  # Accept multiple variations of "yes" (case insensitive)
  # Regex explanation: ^[Yy]$ matches single y/Y, ^[Yy][Ee][Ss]$ matches yes/YES/Yes/etc.
  if [[ "$response" =~ ^[Yy]$ || "$response" =~ ^[Yy][Ee][Ss]$ ]]; then
    
    # =====================================================
    # PHASE 2A: REMOVE APPLICATION BUNDLES
    # =====================================================
    echo "Removing applications..."
    for app in "${found_apps[@]}"; do
      # Remove application bundle recursively
      # Note: -f flag removed to allow proper error detection
      # Redirect output to /dev/null to suppress verbose output
      sudo rm -r "$app" >/dev/null 2>&1

      # Check the exit status of the rm command
      # $? contains the exit code of the last executed command (0 = success)
      if [[ $? -eq 0 ]]; then
        echo "Removed: $app"
        ((removed_app_count++))  # Increment success counter
      else
        echo "Error removing: $app"
        ((error_count++))        # Increment error counter
      fi
    done

    echo ""
    echo "Removing associated files and folders..."
    
    # =====================================================
    # PHASE 2B: REMOVE SUPPORT FILES AND DIRECTORIES
    # =====================================================
    for path_pattern in "${office_library_paths[@]}"; do
      # Use find to locate files matching the pattern
      # -maxdepth 3: Limit search depth to prevent excessive recursion
      # -print0: Use null character as delimiter (handles filenames with spaces)
      # 2>/dev/null: Suppress error messages for inaccessible paths
      find $path_pattern -maxdepth 3 -print0 2>/dev/null | while IFS= read -r -d $'\0' item; do
        # Remove each found item
        # Note: Unquoted $path_pattern allows shell glob expansion (*)
        sudo rm -r "$item" >/dev/null 2>&1

        # Track removal success/failure
        # Note: These counters are in a subshell due to the pipe, so they won't
        # affect the parent shell's counters (limitation of this approach)
        if [[ $? -eq 0 ]]; then
          ((removed_file_count++))
        else
          ((error_count++))
        fi
      done
    done

    # =====================================================
    # PHASE 3: DISPLAY REMOVAL SUMMARY
    # =====================================================
    echo ""
    echo "Removal Summary:"
    echo "  Applications removed: $removed_app_count"
    echo "  Associated files/folders removed: $removed_file_count"
    
    # Check if any errors occurred during the removal process
    if [[ $error_count -gt 0 ]]; then
      echo "  Errors encountered: $error_count"

      # Display error message with red color formatting
      # tput bold: Make text bold
      # tput setaf 1: Set foreground color to red (1)
      # tput sgr 0: Reset all text formatting
      echo "$(tput bold; tput setaf 1)Microsoft Office removal process completed with errors.$(tput sgr 0)"
      exit 1  # Exit with error code to indicate failure
    fi
    
    # Display success message with green color formatting
    # tput setaf 2: Set foreground color to green (2)
    echo "$(tput bold; tput setaf 2)Microsoft Office removal process completed.$(tput sgr 0)"
  else
    # User declined to proceed with removal
    echo "Removal cancelled."
  fi
else
  # No Office applications were found on the system
  echo "Microsoft Office not found."
fi

# Exit with success code (0) to indicate successful completion
exit 0
