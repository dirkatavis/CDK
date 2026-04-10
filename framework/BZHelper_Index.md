BZHelper.vbs — Method Index
============================
Authoritative shared library for BlueZone terminal automation.
Location: framework\BZHelper.vbs
Load guard: safe to ExecuteGlobal multiple times (idempotent).
Requires: g_bzhao declared by the calling script before load; Set before first call.

Before any coding session, paste this in with:
"Here is my existing helper library. Do not duplicate any of these."

---

CONNECTION
----------
ConnectBZ()
    → Connects g_bzhao to the active BlueZone session.
    → Returns True on success, False on failure.

DisconnectBZ()
    → Cleanly disconnects and releases g_bzhao (sets to Nothing).
    → No return value.

---

SCREEN READING
--------------
BZReadScreen(length, row, col)
    → Reads `length` characters from the terminal starting at row/col (1-based).
    → Returns the screen content as a String. Max 1920 chars (full 24×80 screen).

IsTextPresent(searchText)
    → Searches the full 24×80 screen for searchText (case-insensitive).
    → Pipe-delimited multi-target: "PROMPT A|PROMPT B" returns True if either matches.
    → Returns True if any target found, False otherwise.

---

INPUT / KEYSTROKES
------------------
BZSendKey(keyValue)
    → Sends a keystroke or text string to the terminal.
    → Handles special keys (e.g. "<NumpadEnter>", "<Enter>", "<F3>") and plain text.
    → Returns True on success, False on error.

---

WAITING / TIMING
----------------
WaitMs(milliseconds)
    → Busy-waits for the specified number of milliseconds.
    → Handles midnight Timer rollover (Timer resets to 0 at midnight).
    → No return value.

WaitForPrompt(promptText, inputValue, sendEnter, timeoutMs, description)
    → Waits for promptText to appear on screen (pipe-delimited multi-target supported).
    → If found and inputValue is non-empty: sends it (auto-detects special keys via "<>").
    → If sendEnter is True: sends <NumpadEnter> after inputValue.
    → timeoutMs: 0 defaults to 5000ms. description: optional label for log messages.
    → Returns True if prompt found, False on timeout.

WaitForAnyOf(targets, timeoutMs)
    → Waits for any one of several pipe-delimited targets to appear on screen.
    → Uses IsTextPresent internally; case-insensitive.
    → timeoutMs: 0 defaults to 5000ms.
    → Returns True if any target found, False on timeout.

---

ERROR RECOVERY
--------------
BZH_RecoverFromVehidError(employeeNumber, nameConfirmText, menuOption)
    → Recovers from "PRESS RETURN TO CONTINUE" (VEHID not on file) terminal state.
    → Dismisses error → navigates to PFC function → enters employee credentials →
      selects menuOption at the ENTER OPTION menu.
    → employeeNumber:   Employee ID string (e.g. "18351"), from config EmployeeNumber.
    → nameConfirmText:  Pipe-delimited name fragment(s) to confirm (e.g. "CAMP|PASTEUR"),
                        from config EmployeeNameConfirm.
    → menuOption:       "1" = return to main RO screen (Maintenance_RO_Closer)
                        "2" = return to PFC sequence prompt (PostFinalCharges, PFC_Scrapper)
    → Returns True on success, False if any step times out.

---

BUSINESS LOGIC
--------------
BZH_GetMatchedBlacklistTerm(blacklistTermsCsv, pauseMs)
    → Scans all RO service line pages for any term in a comma-separated blacklist.
    → Pages through multi-screen ROs: "(MORE ON NEXT SCREEN)" → advance with N+Enter;
      "(END OF DISPLAY)" → last page. Returns to page 1 via B+Enter after scanning.
    → blacklistTermsCsv: comma-separated terms, case-insensitive (e.g. "VEND TO DEALER").
    → pauseMs: delay between page advances (use script's StabilityPause value).
    → Returns the first matched term as a String, or "" if no match found.

---

INTERNAL / INFRASTRUCTURE
--------------------------
BZH_Log(level, message)
    → Internal logging shim. Calls LogResult(level, message) if defined in calling script.
    → Silently no-ops if LogResult is not available — safe in any script context.
    → Not intended for direct use by calling scripts.
