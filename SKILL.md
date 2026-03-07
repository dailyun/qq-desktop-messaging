---
name: qq-desktop-messaging
description: Read recent messages from QQ desktop chats and send text or images to a specific QQ friend or group through the Windows client. Use when Codex needs to inspect QQ chat content, target a named friend or group, calibrate UI Automation or OCR selectors for QQ, or automate message sending on a logged-in Windows desktop session.
---

# QQ Desktop Messaging

## Overview

Use this skill to automate the Windows QQ desktop client through a mixed strategy: UI Automation first, OCR plus anchor-based positioning second, and clipboard plus simulated input as the final fallback. Read visible chat content, send text, send images, and inspect either the UI tree or OCR regions for calibration.

## Prerequisites

1. Run Windows QQ desktop in a logged-in interactive desktop session.
2. Keep the QQ main window visible and not covered by a modal dialog.
3. Install local `tesseract` and Chinese language data if vision mode is required.
4. Use absolute image paths when sending images.

## Mode Precedence

- `auto`: Try UI Automation first, then OCR plus anchor matching.
- `uia`: Use only the control-tree path.
- `vision`: Use only screenshot, OCR, and simulated input.

Use `auto` by default. Use `vision` when QQ updates break the control tree or when editable controls are no longer exposed.

## Commands

### Inspect QQ UI and optional screenshot metadata

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\inspect_qq_ui.ps1 -MaxDepth 5 -IncludeScreenshotMetadata
```

Returns JSON with the QQ window rectangle, the exported UI tree, and optional screenshot metadata for calibration.

### Read messages from a named friend or group

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\read_qq_messages.ps1 -ConversationName "Project Group" -Mode auto -Last 20
```

Returns JSON with:

- `mode_used`
- `conversation_name`
- `messages`
- `confidence`
- `failure_code` on error

Vision mode only returns visible on-screen message lines.

### Send plain text

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\send_qq_message.ps1 -ConversationName "Alice" -ContentType text -Message "Meeting at 8 PM" -Mode auto
```

Behavior:

- Open the target conversation.
- Verify the target title before sending.
- Use `ValuePattern` when available.
- Fall back to focusing the input region, pasting, and pressing Enter.

### Send an image

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\send_qq_message.ps1 -ConversationName "Project Group" -ContentType image -ImagePath "C:\images\status.png" -Mode auto
```

Behavior:

- Require an absolute image path.
- Focus the target input region.
- Copy the image into the Windows clipboard.
- Paste and send with Enter.

## Failure Codes

- `window_not_found`
- `conversation_not_found`
- `message_region_not_found`
- `input_region_not_found`
- `send_not_confirmed`

## Calibration Flow

1. Run `inspect_qq_ui.ps1` first.
2. If UI Automation works, tune selectors in `qq_uia_common.ps1`.
3. If UI Automation fails, inspect OCR anchors and region rules in [vision-calibration.md](/C:/Users/86159/Documents/Playground/skills/qq-desktop-messaging/references/vision-calibration.md).
4. Re-run `read_qq_messages.ps1` or `send_qq_message.ps1` in `vision` mode to isolate OCR issues.
5. Switch back to `auto` after calibration.

## Resources

### Scripts

- `scripts/inspect_qq_ui.ps1`: Export UI tree and optional screenshot metadata.
- `scripts/read_qq_messages.ps1`: Read visible messages from a target conversation.
- `scripts/send_qq_message.ps1`: Send text or an image to a target conversation.
- `scripts/qq_uia_common.ps1`: Shared UI Automation, clipboard, and input helpers.
- `scripts/qq_visual_common.ps1`: Screenshot capture, OCR, anchor matching, and visual-region helpers.

### References

- [selector-calibration.md](/C:/Users/86159/Documents/Playground/skills/qq-desktop-messaging/references/selector-calibration.md): Tune UI Automation selectors.
- [vision-calibration.md](/C:/Users/86159/Documents/Playground/skills/qq-desktop-messaging/references/vision-calibration.md): Tune OCR anchors and region boundaries.
