#!python3.6-32

import sys, os, atexit, msvcrt
from subprocess import run, PIPE
from pathlib import Path

if len(sys.argv) == 3 and sys.argv[1] == '--sign':
    run(('signtool', '/?'), stderr=PIPE, check=True)  # verify that signtool is in the path
    cert_name = sys.argv[2]
elif len(sys.argv) > 1:
    sys.exit(f'Usage: {Path(sys.argv[0]).name} [--sign CERTIFICATE-NAME]')
else:
    cert_name = None

atexit.register(lambda: (print('\nPress any key to exit ...', end='', flush=True), msvcrt.getch()))

vc_redist = Path(__file__).parent / 'vc_redist.x86.exe'
assert vc_redist.is_file()

program_files = os.getenv('ProgramFiles(x86)')or os.getenv('ProgramFiles')
assert program_files
makensis = Path(program_files) / r'NSIS\makensis.exe'
assert makensis.is_file()

output = run(('py', '-3.6-32', '-c', 'import sys; sys.stdout.buffer.write(sys.executable.encode("utf-8"))'),
             stdout=PIPE, encoding='utf-8', check=True)
scripts = Path(output.stdout).parent / r'Scripts'

run((str(scripts / 'pip3.6'), 'install', '--upgrade', '--upgrade-strategy', 'only-if-needed', 'pyinstaller'), check=True)

os.chdir('..')
run((str(scripts / 'pyinstaller'),
     '--windowed',
     '--add-data', 'Undo_MoM2e.ico;.',
     '-i', 'Undo_MoM2e.ico',
     '--version-file', r'installer\file_version_info.txt',
     'Undo_MoM2e.pyw'),
     check=True)

if cert_name:
    sign_args_sha1   = 'signtool', 'sign', '/v' , '/n', cert_name
    sign_args_sha256 = sign_args_sha1[:]
    sign_args_sha1   += '/t', 'http://timestamp.verisign.com/scripts/timstamp.dll'
    sign_args_sha256 += '/fd', 'sha256', '/tr', 'http://sha256timestamp.ws.symantec.com/sha256/timestamp', '/td', 'sha256', '/as'
    filename_to_sign = r'dist\Undo_MoM2e\Undo_MoM2e.exe',
    run(sign_args_sha1   + filename_to_sign, check=True)
    run(sign_args_sha256 + filename_to_sign, check=True)

os.chdir('installer')
run((str(makensis), 'Undo_MoM2e.nsi'), check=True)

if cert_name:
    filename_to_sign = 'Undo_v1.0_for_MoM2e_setup.exe',
    run(sign_args_sha1   + filename_to_sign, check=True)
    run(sign_args_sha256 + filename_to_sign, check=True)

print('\nBuild succeeded.', end='')
