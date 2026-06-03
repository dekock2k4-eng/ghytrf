"""
create_addin.py -- Builds SheetMind.xlam and installs it into Excel.

Run once:
    python create_addin.py

Two paths:
  Primary  (pywin32 available) -- full 4-button ribbon via COM automation
  Fallback (no pywin32)        -- copy xlwings template + patch ribbon
"""
from __future__ import annotations
import os
import sys
import shutil
import zipfile
from types import SimpleNamespace

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ADDIN_PATH = os.path.join(SCRIPT_DIR, 'SheetMind.xlam')

# ── VBA for the PRIMARY (4-button) path ────────────────────────────────────
VBA_CODE = """\
' SheetMind Add-in -- requires xlwings base add-in to be loaded

Sub RunTextCommand(control As IRibbonControl)
    Application.Run "xlwings.RunPython", "import addin; addin.run_text_command()"
End Sub

Sub RunVoiceCommand(control As IRibbonControl)
    Application.Run "xlwings.RunPython", "import addin; addin.run_voice_command()"
End Sub

Sub UndoLast(control As IRibbonControl)
    Application.Run "xlwings.RunPython", "import addin; addin.undo_last()"
End Sub

Sub ShowInfo(control As IRibbonControl)
    Application.Run "xlwings.RunPython", "import addin; addin.show_info()"
End Sub
"""

# ── Ribbon XML for the PRIMARY path (4 buttons) ────────────────────────────
RIBBON_4BTN = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabSheetMind" label="SheetMind">
        <group id="grpAI" label="AI Commands">
          <button id="btnRun"
                  label="Run Command"
                  imageMso="PlayButton"
                  size="large"
                  onAction="RunTextCommand"
                  screentip="Type a natural language command to modify the sheet"/>
          <button id="btnVoice"
                  label="Voice"
                  imageMso="AudioRecordStart"
                  size="large"
                  onAction="RunVoiceCommand"
                  screentip="Pick a voice recording -- Whisper transcribes and runs it"/>
        </group>
        <group id="grpTools" label="Tools">
          <button id="btnUndo"
                  label="Undo"
                  imageMso="Undo"
                  size="normal"
                  onAction="UndoLast"
                  screentip="Undo the last SheetMind change"/>
          <button id="btnInfo"
                  label="Sheet Info"
                  imageMso="TableInsert"
                  size="normal"
                  onAction="ShowInfo"
                  screentip="Show column names, row count, and undo depth"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>"""

# ── Ribbon XML for the FALLBACK path (single launcher button) ──────────────
RIBBON_1BTN = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabSheetMind" label="SheetMind">
        <group id="grpAI" label="AI Assistant">
          <button id="btnLaunch"
                  label="Open SheetMind"
                  imageMso="PlayButton"
                  size="large"
                  onAction="RunMain"
                  screentip="Open the SheetMind AI panel"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>"""

# ── Open XML content-type / relationship entries ───────────────────────────
_CT_ENTRY = ('<Override PartName="/customUI/customUI14.xml"'
             ' ContentType="application/xml"/>')
_REL_ENTRY = (
    '<Relationship Id="rIdSM"'
    ' Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"'
    ' Target="customUI/customUI14.xml"/>'
)


# ══════════════════════════════════════════════════════════════════════════════
# Helpers
# ══════════════════════════════════════════════════════════════════════════════

def _ensure_xlwings() -> None:
    try:
        import xlwings  # noqa: F401
    except ImportError:
        print('  Installing xlwings...')
        import subprocess
        subprocess.run([sys.executable, '-m', 'pip', 'install', 'xlwings'], check=True)


def _install_xlwings_addin() -> None:
    """Install xlwings.xlam to XLSTART using xlwings.cli directly."""
    print('Installing xlwings base add-in...')
    try:
        import xlwings.cli as xlcli
        args = SimpleNamespace(dir=None, file=None, glob=False)
        xlcli.addin_install(args)
        print('  xlwings base add-in installed.')
    except Exception as e:
        print(f'  Warning: could not auto-install xlwings add-in: {e}')
        print('  You can install it manually from the xlwings tab in Excel.')


def _try_fix_pywin32() -> bool:
    """Ensure pywin32 is installed and DLLs registered. Returns True if win32com works.

    pywin32 is Windows-only. On macOS/Linux we skip it entirely and let the caller
    fall back to the cross-platform xlwings template path."""
    if sys.platform != 'win32':
        print('  Not on Windows — skipping pywin32 (will use the xlwings template path).')
        return False

    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        pass

    print('  Installing pywin32...')
    import subprocess
    try:
        subprocess.run([sys.executable, '-m', 'pip', 'install', 'pywin32'], check=True)
    except Exception as e:  # noqa: BLE001
        print(f'  Could not install pywin32 ({e}); falling back to template path.')
        return False

    print('  Running pywin32 post-install registration...')
    scripts_dir = os.path.join(os.path.dirname(sys.executable), 'Scripts')
    post_script = os.path.join(scripts_dir, 'pywin32_postinstall.py')
    try:
        if os.path.exists(post_script):
            subprocess.run([sys.executable, post_script, '-install'], check=False)
        else:
            subprocess.run([sys.executable, '-m', 'pywin32_postinstall', '-install'], check=False)
    except Exception:  # noqa: BLE001
        pass

    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


def _patch_ribbon(xlam_path: str, ribbon_xml: str) -> None:
    """Inject (or replace) customUI/customUI14.xml in the .xlam ZIP archive."""
    tmp = xlam_path + '.sm_tmp'
    with zipfile.ZipFile(xlam_path, 'r') as zin:
        names = zin.namelist()
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                fn = item.filename
                data = zin.read(fn)

                if fn == '[Content_Types].xml':
                    text = data.decode('utf-8')
                    if 'customUI' not in text:
                        text = text.replace('</Types>', _CT_ENTRY + '</Types>')
                    zout.writestr(fn, text.encode('utf-8'))

                elif fn == '_rels/.rels':
                    text = data.decode('utf-8')
                    if 'extensibility' not in text:
                        text = text.replace('</Relationships>',
                                            _REL_ENTRY + '</Relationships>')
                    zout.writestr(fn, text.encode('utf-8'))

                elif fn == 'customUI/customUI14.xml':
                    zout.writestr(fn, ribbon_xml.encode('utf-8'))

                else:
                    zout.writestr(item, data)

            if 'customUI/customUI14.xml' not in names:
                zout.writestr('customUI/customUI14.xml', ribbon_xml.encode('utf-8'))

    os.replace(tmp, xlam_path)
    print('  Ribbon XML patched.')


# ── Primary path ─────────────────────────────────────────────────────────────

def _create_via_pywin32() -> bool:
    """Create SheetMind.xlam using Excel COM. Returns True on success."""
    try:
        import win32com.client as win32
    except ImportError:
        return False

    xl = None
    try:
        xl = win32.DispatchEx('Excel.Application')
        xl.Visible = False
        xl.DisplayAlerts = False
        wb = xl.Workbooks.Add()

        try:
            _ = wb.VBProject
        except Exception:
            print()
            print('  Excel Trust Center is blocking VBA access.')
            print('  Fix: File -> Options -> Trust Center -> Trust Center Settings')
            print('       -> Macro Settings -> check "Trust access to the VBA project object model"')
            print('  Then re-run this script.')
            wb.Close(False)
            return False

        mod = wb.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule
        mod.Name = 'SheetMindModule'
        mod.CodeModule.AddFromString(VBA_CODE.strip())

        if os.path.exists(ADDIN_PATH):
            os.remove(ADDIN_PATH)

        wb.SaveAs(ADDIN_PATH, FileFormat=55)   # xlAddIn (.xlam)
        wb.Close(False)
        print(f'  Saved: {ADDIN_PATH}')
        return True

    except Exception as e:
        print(f'  COM error: {e}')
        return False
    finally:
        if xl is not None:
            try:
                xl.Quit()
            except Exception:
                pass


# ── Fallback path ─────────────────────────────────────────────────────────────

def _create_via_template() -> bool:
    """
    Copy xlwings's built-in quickstart_addin_ribbon.xlam as our base.
    This requires no pywin32 and no subprocess calls.
    """
    try:
        import xlwings
        pkg_dir = os.path.dirname(xlwings.__file__)
        src = os.path.join(pkg_dir, 'quickstart_addin_ribbon.xlam')
        if not os.path.exists(src):
            print(f'  Template not found: {src}')
            return False

        if os.path.exists(ADDIN_PATH):
            os.remove(ADDIN_PATH)
        shutil.copy2(src, ADDIN_PATH)
        print(f'  Copied template to: {ADDIN_PATH}')

        # Create SheetMind.py -- quickstart VBA calls "import SheetMind; SheetMind.main()"
        wrapper = os.path.join(SCRIPT_DIR, 'SheetMind.py')
        with open(wrapper, 'w') as f:
            f.write('import addin\n\ndef main():\n    addin.main()\n')
        print(f'  Created wrapper: {wrapper}')
        return True

    except Exception as e:
        print(f'  Template copy failed: {e}')
        return False


# ── Instructions ──────────────────────────────────────────────────────────────

def _print_install_steps(mode: str) -> None:
    print()
    print('=' * 62)
    if mode == 'primary':
        print('  SheetMind.xlam ready  --  4-button ribbon')
    else:
        print('  SheetMind.xlam ready  --  single launcher button')
        print('  (upgrade to 4-button: enable Trust Center then re-run)')
    print('=' * 62)
    print()
    print('Install in Excel:')
    print('  1. Open Excel')
    print('  2. File -> Options -> Add-ins')
    print('  3. "Manage: Excel Add-ins" -> Go')
    print('  4. Browse -> select:')
    print(f'     {ADDIN_PATH}')
    print('  5. Tick SheetMind -> OK')
    print()
    print('You will see a "SheetMind" tab in the Excel ribbon.')
    if mode == 'primary':
        print()
        print('Ribbon buttons:')
        print('  Run Command  -> type e.g. "Add 50 to Qty where Item is Pen"')
        print('  Voice        -> pick a .wav/.mp3 file to transcribe & run')
        print('  Undo         -> revert last AI change')
        print('  Sheet Info   -> columns, row count, undo depth')
    else:
        print()
        print('Click "Open SheetMind" to launch the AI panel.')
    print()
    print('IMPORTANT: The xlwings base add-in must also be loaded in Excel.')
    print('  It is installed automatically to your XLSTART folder.')
    print('  If the SheetMind ribbon buttons show an error, open Excel and go to:')
    print('  File -> Options -> Add-ins -> Manage: Excel Add-ins -> Go')
    print('  and make sure "xlwings" is checked.')


# ══════════════════════════════════════════════════════════════════════════════
# Main
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    print('SheetMind -- Excel Add-in Installer')
    print('-' * 40)

    print('\nEnsuring xlwings is installed...')
    _ensure_xlwings()

    _install_xlwings_addin()

    print('\nChecking pywin32...')
    have_pywin32 = _try_fix_pywin32()

    if have_pywin32:
        print('  pywin32 OK -- using primary path (4-button ribbon)')
        print('\nCreating SheetMind.xlam via COM...')
        ok = _create_via_pywin32()
        if ok:
            print('Patching ribbon...')
            _patch_ribbon(ADDIN_PATH, RIBBON_4BTN)
            _print_install_steps('primary')
            sys.exit(0)
        print('\nCOM path failed. Falling back to template copy...')
    else:
        print('  pywin32 unavailable -- using fallback path (single-button ribbon)')

    print('\nCreating SheetMind.xlam from xlwings template...')
    ok = _create_via_template()
    if not ok:
        print('\nInstallation failed. Please check the errors above.')
        sys.exit(1)

    print('Patching ribbon...')
    _patch_ribbon(ADDIN_PATH, RIBBON_1BTN)
    _print_install_steps('fallback')
