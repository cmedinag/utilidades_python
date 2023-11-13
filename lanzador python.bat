echo off
set script="%cd%\nombre_de_tu_fichero_con_el_codigo_python.py"
set CDIR=%cd%

::Si existe el python instalado directamente, vamos por esa vía
if exist "C:\Program Files\python39-32\python.exe" (
   "C:\Program Files\python39-32\python.exe" %script%
   goto finalizar
)

::Si no existe el python directamente, buscamos la instalación de Anaconda en las posibles carpetas
if exist "C:\Program Files\Anaconda\python.exe" (
	set pyexe="C:\Program Files\Anaconda\python.exe"
	set pydir="C:\Program Files\Anaconda"
) else (
	if exist "C:\Program Files\Anaconda3\python.exe" (
		set pyexe="C:\Program Files\Anaconda3\python.exe"
		set pydir="C:\Program Files\Anaconda3"
	) else (
		if exist %USERPROFILE%\Anaconda3\python.exe (
			set pyexe=%USERPROFILE%\Anaconda3\python.exe
			set pydir=%USERPROFILE%\Anaconda3
		) else (
			if exist "C:\ProgramData\Anaconda3\python.exe" (
				set pyexe="C:\ProgramData\Anaconda3\python.exe"
				set pydir="C:\ProgramData\Anaconda3"
			) else (
				if exist %USERPROFILE%\Anaconda\python.exe (
					set pyexe=%USERPROFILE%\Anaconda\python.exe
					set pydir=%USERPROFILE%\Anaconda
				)
			)
		)
	)
)
%pyexe% %pydir%\cwp.py %pydir% %pyexe% %script%

:finalizar
pause
