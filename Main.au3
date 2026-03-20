; ============================================
; PDF to Word Fixer Pro v6.1 - MODULAR ARCHITECTURE
; Entry Point - Chay file nay de khoi dong
; ============================================

#AutoIt3Wrapper_UseX64=n
#include-once

; === INCLUDE CORE ===
#include "Config.au3"
#include "Core\WordConnection.au3"
#include "Core\EventLoop.au3"

; === INCLUDE SHARED ===
#include "Shared\Helpers.au3"
#include "Shared\WordOps.au3"
#include "Shared\UIHelpers.au3"

; === INCLUDE GUI ===
#include "GUI\MainGUI.au3"
#include "GUI\TabPDFFix.au3"
#include "GUI\TabFormat.au3"
#include "GUI\TabTools.au3"
#include "GUI\TabTOC.au3"
#include "GUI\TabCopyStyle.au3"
#include "GUI\TabAdvanced.au3"
#include "GUI\TabQuickUtils.au3"
#include "GUI\TabSmartFix.au3"
#include "GUI\TabAIFormat.au3"
#include "GUI\Dialogs.au3"
#include "GUI\HotkeyDialogs.au3"

; === INCLUDE MODULES ===
#include "Modules\PDFFix.au3"
#include "Modules\Format.au3"
#include "Modules\Tools.au3"
#include "Modules\TOC.au3"
#include "Modules\StyleHotkey.au3"
#include "Modules\CopyStyle.au3"
#include "Modules\Advanced.au3"
#include "Modules\QuickUtils.au3"
#include "Modules\SmartFix.au3"
#include "Modules\AIFormat.au3"

; === MAIN ===
Opt("GUIOnEventMode", 0)
Opt("MustDeclareVars", 0)

_InitApplication()
