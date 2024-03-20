set config=%1

call .\.venv\scripts\activate
.\.venv\scripts\python -m cast_arg.main -c %config% 