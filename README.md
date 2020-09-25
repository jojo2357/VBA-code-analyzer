In order to analyze .bas files, put the exe in the folder that contains the files that you want to analyze and run.

If you would like to analyze bas files in all subfolders for, say, a project, change line one from `SHELL _HIDE "dir /b *.bas>filenames.txt"` to `SHELL _HIDE "dir /b /s *.bas>filenames.txt"` and recompile in QB64