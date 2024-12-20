**SYNOPSIS**   
--------------------------------
Generates an HTML email signature for Microsoft Outlook using Base64 encoded images.

**DESCRIPTION**  
This script creates Outlook signatures using Base64 encoded images instead of downloading external files.
All logos are embedded directly in the HTML, making the signatures more portable and reliable.
Includes support for multiple templates, backup functionality, and Microsoft Graph API integration.  

**REQUIREMENTS**  
  - PowerShell 5.1 or later  
  - Microsoft Graph API access  
  - Outlook 2016 or later  
  - Windows 10/11   

**PARAMETER -Logo**  
Specifies which logo to use in the signature. Available options:  
  - FoC (default)
  - H2K
  - GSINA
  - GSIPA
  - GSIT
  - GSIAM
  - GSISG
  - GSIVET

**EXAMPLE**  
`-Logo GSIPA`

**PARAMETER -User**  
Specify a User. Files will be saved to C:\Temp\Signatures

**EXAMPLE**  
`-User JDoe`

**PARAMETER -Cleanup**    
Cleans Outlook Registry keys and backs up existing signatures to `C:\Users\USERNAME\Documents\Outlook Signatures Backup` 
Deletes content from `%APPDATA%\Microsoft\Signatures`
Use this when troubleshooting signature issues.

**EXAMPLE**  
`-Cleanup`

**CHANGELOG**  
--------------------------------

**To Do**  
  - Convert Address information to new formatting with UTC Time-Zone
  - Troubleshoot Environet default settings via Registry (currently works if capitalize the E in file name)
  - Research/Test further Registry settings
  - Test on fresh machine, review virgin registry keys and settings

**Version 1.0.0**
  - Complete rebuild of syntax for PowerShell Core (v5.1) compatibility
  - Now works in TRMM
  - Redesign default logo
  - Redesign GSIVET logo
  - Removed `-Template` parameter, Logo auto-assigns appropriate template
  - Added registry cleaning to -Cleanup
  - Added multiple registry settings
	  - Disabled Outlook New
	  - Disabled Outlook New Nag notifications
	  - Delete hardcoded Outlook New switchover date
	  - Enabled Signature Cloud upload and sync
  - Improve terminal output messaging
  - Verbose logfile for troubleshooting
    
**Version 0.3.0**  
  - Complete rebuild with embedded HTML templates
  - Removed template file dependencies
  - Automatic template selection based on logo
  - Enhanced Environet account detection and handling
  - Dual mailbox signature configuration
  - Registry settings for both primary and secondary mailboxes
  - Detailed execution logging system
  - Base64 encoded logos
  - Improved resource management and cleanup
  - Simplified logo selection process
  - Enhanced error handling and logging
    
**Version 0.2.0**  
  - Removed dependency on external image files
  - Added Base64 encoded images
  - Rewrote HTML code to support New Outlook and OWA
  - Added automatic signature backup
  - Added `-Cleanup` parameter for signature management
  - Added registry settings for New Outlook and OWA
  - Added phone number formatting
  - Added Outlook version detection
  - Improved parameter validation
  - Improved formatting and efficiency of code
  - Added ability to easily add alternative templates with switch -Template
    
**Version 0.1.5**  
  - Fixed `-Logo` parameter to work correctly in Tactical RMM
  - All non Environet email accounts have logoName added to the filenames to allow for multiple signatures
  - Cleaned up Output to be more readable
  - Added  $lowercaseWords (line 233) to allow for specific words to be in lowercase
    
**Version 0.1.4**   
  - You can now define a specific user with `.\Set-OutlookSignatures.ps1 -User JDoe` or `-User JDoe` in TRMM field
	  - Run the script on your computer with `-User JDoe` and it will save to your `C:\Temp\Signatures`
  - Finished all templates
  - Fixed Registry settings, no longer greys options out and recursively scans for existing Profiles 
  - Further optimized logic flow and Functions
  - Added more helpers for converting variables to convert phone number formatting
  - Greatly improved logging and debug details saved in LogFile
  - Added token caching to reduce API calls and improve performance
  - Implemented retry logic for logo downloads with proper resource cleanup
  - Added automatic cleanup of signatures
  - Enhanced error handling with detailed JSON logging
  - Fixed line spacing issues in log output
  - Improved registry handling to properly clear existing signatures before setting new ones
  - Added WebClient implementation for better download performance
  - Added proper resource disposal for web requests
    
**Version 0.1.3**  
  - Added support for testing against a specific user.
  - Finished Environet New Email template.
  - Optimized logic flow and Functions.
  - Added new helpers for converting variables to all Uppercase or Lowercase.
  - Added more logging information for Environet.
  - Enhanced script output with more information and better readibility.
    
**Version 0.1.2**  
  - Set Microsoft Graph with registered Entra App as default method for pulling user information, falls back to AD if fails.
  - Checks if user has an Environet account
  - Began HTML template to generate Environet "New Email" signatures
    
**Version 0.1.1**  
  - Enhanced registry settings to manage Outlook's default signature configuration.
  - Added checks to remove "First Run" registry entry to ensure users can modify the default signatures.
  - Ensured `DisableRoamingSignatures` and `DisableRoamingSignaturesTemporaryToggle` are set to allow for roaming signatures.
    
**Version 0.1.0**  
  - Initial script to generate signatures dynamically based on Active Directory details.
