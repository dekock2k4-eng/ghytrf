"""
sheetmind_autowrap.py — make the xlwings "workbook name = Python module" rule
disappear.

xlwings runs `import <WorkbookName>; <WorkbookName>.main()`. Normally that needs a
hand-made <WorkbookName>.py wrapper next to addin.py. This module installs a
last-resort import hook so that ANY otherwise-unresolvable top-level import (i.e.
the workbook module xlwings asks for) is satisfied automatically by a synthetic
module whose main()/run_*()/undo_last()/show_info() forward to `addin`. It also
writes a real <WorkbookName>.py the first time, so things keep working even if
this hook is ever removed.

It only activates inside the xlwings invocation pattern (`python -c "..."`), so it
never masks genuine import errors during normal development.

Loaded automatically at interpreter startup via a .pth file in site-packages
(installed by create_addin.py / install_autowrap()).
"""
import importlib.abc
import importlib.machinery
import os
import sys

_FORWARD = ("main", "run_text_command", "run_voice_command", "undo_last", "show_info")
_EXCEL_EXT = (".xlsx", ".xlsm", ".xlsb", ".xls")


def _is_workbook(name):
    """True only if `name` matches a real Excel file on disk — so we synthesise
    ONLY the workbook module xlwings asks for, never optional deps like `pytz`."""
    dirs = ["."]
    for p in sys.path:
        if not p:
            continue
        dirs.append(p)
        if os.path.isfile(p):           # xlwings may pass the workbook's full path
            dirs.append(os.path.dirname(p))
    home = os.path.expanduser("~")
    dirs += [os.path.join(home, d) for d in ("Documents", "Desktop", "Downloads", "")]
    seen = set()
    for d in dirs:
        try:
            d = os.path.abspath(d)
        except (OSError, ValueError):
            continue
        if d in seen:
            continue
        seen.add(d)
        for ext in _EXCEL_EXT:
            if os.path.isfile(os.path.join(d, name + ext)):
                return True
    return False


def _project_dir():
    """Directory that contains addin.py (added to sys.path by xlwings)."""
    for p in sys.path:
        try:
            if p and os.path.isfile(os.path.join(p, "addin.py")):
                return p
        except OSError:
            pass
    return None


def _write_wrapper(name):
    try:
        d = _project_dir()
        if not d:
            return
        wf = os.path.join(d, f"{name}.py")
        if not os.path.exists(wf):
            with open(wf, "w") as f:
                f.write('"""Auto-generated xlwings wrapper — forwards to SheetMind."""\n'
                        "import addin\n\n\ndef main():\n    addin.main()\n")
    except OSError:
        pass


class _Loader(importlib.abc.Loader):
    def __init__(self, name):
        self.name = name

    def create_module(self, spec):
        return None

    def exec_module(self, module):
        _write_wrapper(self.name)

        def _make(fn):
            def inner(*a, **k):
                import addin
                return getattr(addin, fn)()
            return inner
        for fn in _FORWARD:
            setattr(module, fn, _make(fn))


class _Finder(importlib.abc.MetaPathFinder):
    def find_spec(self, name, path=None, target=None):
        # Only synthesise a top-level name that corresponds to a real workbook
        # file. We're appended last, so we only reach here for names nothing else
        # could import — but we still require a matching .xlsx so we never shadow
        # a genuinely-missing optional dependency (e.g. pytz).
        if path is not None or "." in name or name.startswith("_"):
            return None
        if not _is_workbook(name):
            return None
        return importlib.machinery.ModuleSpec(name, _Loader(name))


def install():
    # Activate only for the xlwings `python -c "..."` call, never for dev tools.
    if not (sys.argv and sys.argv[0] == "-c"):
        return
    if not any(type(f).__name__ == "_Finder" for f in sys.meta_path):
        sys.meta_path.append(_Finder())


def install_autowrap(site_packages_dir, project_dir):
    """Write the .pth that auto-loads this hook at interpreter startup."""
    line = (f"import sys, os; _p=r'{project_dir}'; "
            f"(_p in sys.path) or sys.path.insert(0, _p); "
            f"import sheetmind_autowrap\n")
    pth = os.path.join(site_packages_dir, "sheetmind_autowrap.pth")
    with open(pth, "w") as f:
        f.write(line)
    return pth


install()
