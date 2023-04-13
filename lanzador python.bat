echo off
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
%pyexe% %pydir%\cwp.py %pydir% %pyexe% "%cd%\script.py"

pause


