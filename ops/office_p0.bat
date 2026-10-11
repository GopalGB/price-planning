@echo off
setlocal
cd /d "%~dp0\.."
for /f "delims=" %%P in ('py -3 ops\install_env.py') do set "VENV_PY=%%P"
if not defined VENV_PY ( echo install failed & exit /b 1 )
"%VENV_PY%" -m pip install --no-index --find-links vendor\wheels --require-hashes -r requirements-dev.lock || exit /b 1
if not exist results-private mkdir results-private
set "XLSM=Price Planning Template Spring 2027.xlsm"
"%VENV_PY%" ops\find_inputs.py "%XLSM%" > results-private\inputs.txt
if errorlevel 1 exit /b 1
"%VENV_PY%" -c "import json;m=json.load(open('ops/bench_manifest.json'));l=[x.strip() for x in open('results-private/inputs.txt',encoding='utf-8') if x.strip()];m['candidate_inputs']=l[::max(1,len(l)//50)][:50];json.dump(m,open('results-private/manifest.json','w'),indent=2)"
if errorlevel 1 exit /b 1
"%VENV_PY%" ops\bench_real.py --xlsm "%XLSM%" --manifest results-private\manifest.json
if errorlevel 1 exit /b 1
"%VENV_PY%" ops\bench_real.py --xlsm "%XLSM%" --manifest results-private\manifest.json
if errorlevel 1 exit /b 1
echo Done. Commit .planning\excel-grid\results\*.json on a branch and push.
