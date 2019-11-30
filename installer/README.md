## Building the installer ##

 1. Visit the Python download page here: <https://www.python.org/downloads/windows/>, and click the link for the latest **Python 3.8** release. Download and run the `Windows x86 web-based installer`. (Although Undo for FFG Games can run under 64-bit Python, these instructions and the installer were written for 32-bit Python.)

 2. Download and install the latest 3.x version of the Nullsoft Scriptable Install System from here: <http://nsis.sourceforge.net/Download>.

 3. Download `VC_redist.x86.exe` (Microsoft Visual C++ 2015-2019 Redistributable (x86)) into this directory from here:
<https://aka.ms/vs/16/release/vc_redist.x86.exe>.

 4. Double-click the `build_installer.py` file in this directory. The built installer (`Undo_v3.0_for_FFG_setup.exe`) will be placed in this directory.
