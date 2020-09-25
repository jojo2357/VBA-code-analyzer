SHELL _HIDE "dir /b *.bas>filenames.txt"

OPEN "filenames.txt" FOR INPUT AS #1
OPEN "MetaData.csv" FOR OUTPUT AS #3

PRINT #3, "Name, lines, source, comment lines, empty lines, total comments, , source %, line comment %, empty line %, % lines commented"

outFormat$ = "& _, & _, & _, & _, & _, & _,_, ##.#_%_, ##.#_%_, ##.#_%_, ##.#_%"

DO
  LINE INPUT #1, filename$
  OPEN filename$ FOR INPUT AS #2
  comments = 0
  commentLines = 0
  sourceLines = 0
  lines = 0
  emptyLines = 0
  DO
    lines = lines + 1
    LINE INPUT #2, lineIn$
    lineIn$ = RTRIM$(LTRIM$(lineIn$))
    IF LEN(lineIn$) > 0 AND MID$(lineIn$, 1, 1) <> "'" AND NOT INSTR(lineIn$, "REM") THEN
      sourceLines = sourceLines + 1
    END IF
    IF LEN(lineIn$) = 0 THEN
      emptyLines = emptyLines + 1
    END IF
    IF INSTR(lineIn$, "'") THEN
      comments = comments + 1
      IF MID$(lineIn$, 1, 1) = "'" OR INSTR(lineIn$, "REM") THEN
        commentLines = commentLines + 1
      END IF
    END IF
  LOOP UNTIL EOF(2)
  CLOSE #2
  totalcomments = totalcomments + comments
  totalcommentLines = totalcommentLines + commentLines
  totalsourceLines = totalsourceLines + sourceLines
  totallines = totallines + lines
  totalemptyLines = totalemptyLines + emptyLines
  PRINT USING outFormat$; filename$; STR$(lines); STR$(sourceLines); STR$(commentLines); STR$(comments); STR$(emptyLines); 100 * sourceLines / lines; 100 * commentLines / lines; 100 * emptyLines / lines; 100 * comments / lines
  PRINT #3, USING outFormat$; filename$; STR$(lines); STR$(sourceLines); STR$(commentLines); STR$(comments); STR$(emptyLines); 100 * sourceLines / lines; 100 * commentLines / lines; 100 * emptyLines / lines; 100 * comments / lines
LOOP UNTIL EOF(1)

CLOSE #1

PRINT #3, USING outFormat$; "Totals: "; STR$(totallines); STR$(totalsourceLines); STR$(totalcommentLines); STR$(totalemptyLines); STR$(totalcomments); 100 * totalsourceLines / totallines; 100 * totalcommentLines / totallines; 100 * totalemptyLines / totallines; 100 * totalcomments / totallines

CLOSE #3

SYSTEM
