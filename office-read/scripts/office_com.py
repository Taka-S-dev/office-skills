"""Driving Word, Excel and PowerPoint through COM from Python, safely.

run_script() runs a PowerShell script in a job with a time limit: a hidden dialog (repair,
password, conversion) makes an Office process wait forever, so on the limit the job is stopped
and the hidden instances it started are ended. The script and every path travel in environment
variables, never in the command line.

convert_legacy() turns a .doc / .xls / .ppt into its x-format in a temporary folder, with
Office when it is installed (macros disabled, read-only) and LibreOffice otherwise.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

LIMIT_S = 240

_GUARD = """
$t0 = Get-Date
$job = Start-Job -ScriptBlock ([scriptblock]::Create($env:OFFICE_SCRIPT))
if (Wait-Job $job -Timeout LIMIT) { Receive-Job $job 2>$null }
else {
  Stop-Job $job
  # only hidden instances started by this call: a window the user opened meanwhile is left alone
  Get-Process IMAGE -ErrorAction SilentlyContinue |
    Where-Object { $_.StartTime -ge $t0 -and $_.MainWindowHandle -eq 0 } | Stop-Process -Force
  Write-Output 'TIMEOUT'
}
Remove-Job $job -Force
"""


def run_script(script: str, image: str, env: dict[str, str], limit: int = LIMIT_S) -> str:
    """stdout of the script; 'TIMEOUT' in it when Office had to be stopped."""
    wrapper = _GUARD.replace("LIMIT", str(int(limit))).replace("IMAGE", image)
    r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", wrapper],
                       capture_output=True, text=True, timeout=limit + 60,
                       env={**os.environ, **env, "OFFICE_SCRIPT": script})
    if "TIMEOUT" in r.stdout:
        print(f"[office] {image.title()} did not answer within {limit} s (a hidden dialog?); stopped it", file=sys.stderr)
    return r.stdout


_RELEASE = """
  foreach ($o in @($doc, $app)) { if ($o) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) } }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
"""

# every Open argument is given: COM from PowerShell cannot skip one. A placeholder password makes a
# protected file fail at once instead of waiting on a prompt nobody can see.
_CONVERT = {
    ".doc": ("WINWORD", """
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject Word.Application
try {
  $app.Visible = $false; $app.DisplayAlerts = 0; $app.AutomationSecurity = 3
  $doc = $app.Documents.Open($env:OFFICE_SRC, $false, $true, $false, '_', '_', $false, '_', '_', 0, 65001, $false, $false, 0, $true)
  $doc.SaveAs2($env:OFFICE_OUT, 16)
  $doc.Close(0)
  Write-Output 'OK'
} finally { $app.Quit()
""" + _RELEASE + "}", ".docx", "docx"),
    ".xls": ("EXCEL", """
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject Excel.Application
try {
  $app.Visible = $false; $app.DisplayAlerts = $false; $app.AutomationSecurity = 3
  $doc = $app.Workbooks.Open($env:OFFICE_SRC, 0, $true, 5, '_', '_', $true)
  $doc.SaveAs($env:OFFICE_OUT, 51)
  $doc.Close($false)
  Write-Output 'OK'
} finally { $app.Quit()
""" + _RELEASE + "}", ".xlsx", "xlsx"),
    ".ppt": ("POWERPNT", """
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject PowerPoint.Application
try {
  $app.AutomationSecurity = 3
  $doc = $app.Presentations.Open($env:OFFICE_SRC, -1, 0, 0)
  $doc.SaveAs($env:OFFICE_OUT, 24)
  $doc.Close()
  Write-Output 'OK'
} finally { $app.Quit()
""" + _RELEASE + "}", ".pptx", "pptx"),
}


def convert_legacy(src: Path) -> tuple[Path, Path, str] | None:
    """(converted file, temporary folder to remove afterwards, how) or None."""
    ext = src.suffix.lower()
    if ext not in _CONVERT:
        return None
    image, script, new_ext, lo_fmt = _CONVERT[ext]
    tmp = Path(tempfile.mkdtemp(prefix="office_legacy_"))
    out = tmp / (src.stem + new_ext)
    if sys.platform == "win32":
        try:
            run_script(script, image, {"OFFICE_SRC": str(src.resolve()), "OFFICE_OUT": str(out.resolve())})
        except (subprocess.TimeoutExpired, OSError):
            pass
        if out.exists() and out.stat().st_size:
            return out, tmp, f"converted to {new_ext} by Microsoft Office"
    so = shutil.which("soffice") or (r"C:\Program Files\LibreOffice\program\soffice.exe"
                                      if Path(r"C:\Program Files\LibreOffice\program\soffice.exe").exists() else None)
    if so:
        try:
            subprocess.run([so, "--headless", "--convert-to", lo_fmt, "--outdir", str(tmp), str(src.resolve())],
                           capture_output=True, text=True, timeout=LIMIT_S)
        except (subprocess.TimeoutExpired, OSError):
            pass
        if out.exists() and out.stat().st_size:
            return out, tmp, f"converted to {new_ext} by LibreOffice"
    shutil.rmtree(tmp, ignore_errors=True)
    return None
