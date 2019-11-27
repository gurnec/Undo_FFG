#!python3.8-32

import sys, os, winreg, venv
from subprocess import run, DEVNULL
from pathlib import Path

if len(sys.argv) == 3 and sys.argv[1] == '--sign':
    run(('signtool', '/?'), stderr=DEVNULL, check=True)  # verify that signtool is in the path
    cert_name = sys.argv[2]
elif len(sys.argv) > 1:
    sys.exit(f'Usage: {Path(sys.argv[0]).name} [--sign CERTIFICATE-NAME]')
else:
    cert_name = None
    import atexit, msvcrt
    atexit.register(lambda: (print('\nPress any key to exit ...', end='', flush=True), msvcrt.getch()))

working_dir = Path(__file__).parent
vc_redist = working_dir / 'VC_redist.x86.exe'
assert vc_redist.is_file()

try:
    with winreg.OpenKeyEx(winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\NSIS',
                          access=winreg.KEY_QUERY_VALUE | winreg.KEY_WOW64_32KEY) as regkey:
        makensis = Path(winreg.QueryValueEx(regkey, None)[0]) / 'makensis.exe'
except OSError:
    program_files = os.getenv('ProgramFiles(x86)') or os.getenv('ProgramFiles')
    assert program_files
    makensis = Path(program_files) / r'NSIS\makensis.exe'
assert makensis.is_file()

print('Building venv ...')
working_dir = working_dir.parent
os.chdir(working_dir)
venv.EnvBuilder(upgrade=True, with_pip=True).create(r'build\venv')
scripts_dir = working_dir / r'build\venv\Scripts'

print()
run([scripts_dir/'pip'] + 'install --upgrade --upgrade-strategy eager '
    'https://github.com/pyinstaller/pyinstaller/archive/develop.zip'.split(), check=True)

print()
run([scripts_dir/'pyinstaller'] + '--windowed --add-binary Undo_MoM2e.ico;. -i Undo_MoM2e.ico '
    r'--version-file installer\file_version_info.txt --noconfirm Undo_MoM2e.pyw'.split(), check=True)

working_dir /= r'dist\Undo_MoM2e'
working_dir.joinpath('libcrypto-1_1.dll').unlink()
working_dir.joinpath('libssl-1_1.dll').unlink()
working_dir.joinpath('unicodedata.pyd').unlink()

if cert_name:
    print()
    sign_args_sha1   = 'signtool', 'sign', '/v' , '/n', cert_name
    sign_args_sha256 = sign_args_sha1[:]
    sign_args_sha1   += '/t', 'http://timestamp.verisign.com/scripts/timstamp.dll'
    sign_args_sha256 += '/fd', 'sha256', '/tr', 'http://sha256timestamp.ws.symantec.com/sha256/timestamp', '/td', 'sha256', '/as'
    filename_to_sign = r'dist\Undo_MoM2e\Undo_MoM2e.exe',
    run(sign_args_sha1   + filename_to_sign, check=True)
    run(sign_args_sha256 + filename_to_sign, check=True)

print()
os.chdir('installer')
run((makensis, 'Undo_MoM2e.nsi'), check=True)

if cert_name:
    print()
    filename_to_sign = 'Undo_v2.1_for_FFG_setup.exe',
    run(sign_args_sha1   + filename_to_sign, check=True)
    run(sign_args_sha256 + filename_to_sign, check=True)

print('\nBuild succeeded.', end='')
