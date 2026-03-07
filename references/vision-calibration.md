# QQ Vision Calibration

Use this guide when `vision` mode or the `auto` fallback path fails to find the target conversation, the visible message area, or the input region.

## OCR Dependencies

- Install `tesseract` and ensure it is available on `PATH`.
- Install Chinese language data and keep the default language as `chi_sim+eng` unless the QQ client uses another locale.
- Confirm `tesseract image.png stdout --psm 6 -l chi_sim+eng tsv` works before changing OCR logic.

## Region Layout Defaults

The current vision helper assumes these window regions:

- `sidebar`: left 33 percent of the QQ window
- `title`: top 14 percent of the content area on the right
- `messages`: center-right message pane from 12 percent to 68 percent of window height
- `input`: lower-right input pane from 72 percent to 94 percent of window height

Adjust these ratios in `scripts/qq_visual_common.ps1` if QQ uses a different skin or layout.

## Anchor Strategy

- Conversation targeting uses OCR on the `sidebar` region.
- Active-title confirmation uses OCR on the `title` region.
- Message extraction uses OCR line grouping in the `messages` region.
- Input focusing clicks the center of the `input` region.

Prefer matching exact conversation names first. If OCR often splits names, broaden matching with partial-text rules only after checking for duplicate chat names.

## Typical Fixes

### Conversation not found

- Increase the sidebar width ratio if QQ shows a wider navigation list.
- Change OCR language if the current client locale is not Chinese plus English.
- Tighten exact-name preference when OCR picks a nearby unread badge or preview text.

### Message region not found

- Expand the message region downward or upward.
- Change `--psm` mode if OCR merges lines too aggressively.
- Filter low-confidence words before grouping lines.

### Input region not found

- Raise the input region top boundary if QQ shows a larger toolbar.
- Prefer clicking a lower-center point if the editor sits beneath a formatting ribbon.
- Keep QQ foreground checks enabled before simulated paste and Enter.
