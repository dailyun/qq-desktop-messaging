# QQ Selector Calibration

Use this guide when QQ updates break the default heuristics in the bundled PowerShell scripts.

## Calibration Flow

1. Run `scripts/inspect_qq_ui.ps1 -MaxDepth 5 -OutputPath .\qq-ui-tree.txt`.
2. Open `qq-ui-tree.txt` and search for the visible conversation name.
3. Inspect nearby parents to identify which controls represent:
   - the conversation list
   - the active message container
   - the message input area
4. Compare the discovered control types, names, classes, and automation IDs against the heuristics in `scripts/qq_uia_common.ps1`.
5. Patch the helper functions only where the current heuristics are too broad or too weak.

## Heuristics Used By Default

- `Get-QQMainWindow`: pick the largest top-level window whose title matches `QQ`.
- `Open-ConversationByName`: scan descendants for elements whose visible name matches or contains the target conversation name, then invoke or select that element.
- `Find-BestMessageContainer`: pick the `Pane`, `List`, `Document`, or `Custom` element with the largest number of text descendants.
- `Find-BestInputElement`: prefer lower-page `Edit` or `Document` controls and boost controls whose name hints at message input.
- `Send-QQMessage`: prefer `ValuePattern`, otherwise paste from the clipboard and press Enter.

## Typical Fixes

### Conversation search is ambiguous

- Tighten `Find-CandidateElementsByName` to filter by parent control type or class name.
- Prefer exact name matches before partial matches.
- If QQ duplicates visible names, add additional checks on bounding rectangle or ancestor structure.

### Message reads return unrelated text

- Narrow `Find-BestMessageContainer` to a specific control type discovered from the live UI tree.
- Ignore candidate containers that live in the navigation sidebar or title bar.
- Filter duplicate or structural text nodes after extraction.

### Sending focuses the wrong control

- Tighten `Find-BestInputElement` using the actual class name or automation ID from the inspected tree.
- If the editor is not a `ValuePattern` target, keep the clipboard fallback and ensure the input control is focused before pasting.

## Operating Assumptions

- The machine is running Windows and the QQ desktop client.
- QQ is already logged in.
- The current session is interactive, not a locked screen or headless environment.
- No modal dialog blocks the main conversation window.
