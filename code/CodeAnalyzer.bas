SHELL _HIDE "dir /b *.bas>filenames.txt"

OPEN "filenames.txt" FOR INPUT AS #1
OPEN "MetaData.csv" FOR OUTPUT AS #3

PRINT #3, "Name, lines, source, comment lines, total comments, empty lines"

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
  PRINT #3, filename$; ", "; lines; ", "; sourceLines; ", "; commentLines; ", "; comments; ", "; emptyLines
  PRINT filename$; ", "; lines; ", "; sourceLines; ", "; commentLines; ", "; comments; ", "; emptyLines
LOOP UNTIL EOF(1)

CLOSE #1

PRINT filename$, lines, sourceLines, commentLines
