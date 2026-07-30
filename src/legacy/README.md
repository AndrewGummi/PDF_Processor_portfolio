# Legacy / non-functional modules

These two files are kept for completeness but are **not part of the working
pipeline**. Both import modules from an older package layout that no longer
exists in the repository:

| File | Missing import | Status |
|------|----------------|--------|
| `gui_launcher.py`   | `python_src.main_processor`  | fails on import |
| `batch_processor.py`| `excel_utils.excel_writer`  | fails on import |

They are excluded from the documented entry points in the top-level README.
