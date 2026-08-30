# Contributing to DocMerge

Thanks for taking an interest.

## Reporting a bug

Open an issue and include:

- macOS version
- Output of `python3 --version`
- The Python and Tk versions shown in the DocMerge footer
- What you did, what you expected, what happened
- If the app fails to launch, run it from Terminal and paste the output:
  `/Applications/DocMerge.app/Contents/MacOS/DocMerge`

## Development setup

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python3 docmerge.py
```

Test the engine without the GUI:

```bash
python3 docmerge.py out.docx sample1.txt sample2.pdf
```

## Code notes

The whole app is one file, `docmerge.py`, split into two halves: the extraction
and merge engine at the top, the Tkinter interface below. The engine has no
dependency on the GUI, which is why command line mode works.

A few constraints worth knowing before you change the UI:

- Do not use `tk.Button` with a custom `background`. macOS Aqua discards it but
  honours `foreground`, which produces invisible text. Use the `ttk` styles
  defined in `_init_styles`.
- The progress bar is drawn on a `Canvas` rather than using `ttk.Progressbar`,
  for consistent appearance across Tk versions.
- Merge work happens on a background thread and communicates with the UI
  through a `queue.Queue`, polled by `_poll`. Never touch widgets from the
  worker thread.

## Pull requests

Keep them focused on one thing. Match the existing style; the code follows PEP 8
with a 79 column limit. Include a short description of what you tested.
