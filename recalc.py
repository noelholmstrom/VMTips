"""Recalculate an Excel workbook with LibreOffice headless and return a recalced .xlsx path."""
import subprocess, os, tempfile, shutil, glob

def recalc(src_path, profile_dir=None):
    workdir = tempfile.mkdtemp(prefix='recalc_')
    if profile_dir is None:
        profile_dir = os.path.join(workdir, 'profile')
    os.makedirs(profile_dir, exist_ok=True)
    # Force "always recalculate" on load for OOXML + ODF
    cfg_dir = os.path.join(profile_dir, 'user')
    os.makedirs(cfg_dir, exist_ok=True)
    with open(os.path.join(cfg_dir, 'registrymodifications.xcu'), 'w') as f:
        f.write('''<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
 <item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
 <item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>''')
    env = dict(os.environ, HOME=workdir)
    cmd = ['soffice', '-env:UserInstallation=file://'+profile_dir,
           '--headless','--norestore','--convert-to','xlsx:Calc MS Excel 2007 XML',
           '--outdir', workdir, src_path]
    subprocess.run(cmd, env=env, check=True, capture_output=True, timeout=180)
    out = glob.glob(os.path.join(workdir, '*.xlsx'))
    if not out:
        raise RuntimeError('recalc produced no output')
    return out[0]

if __name__ == '__main__':
    import sys
    print(recalc(sys.argv[1]))
