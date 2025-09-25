# hooks/rth_lockdown.py
# Полный локдаун путей: убираем CWD, PYTHONPATH, user site.
import os, sys

# a) Сброс переменных окружения, которые могут ломать импорт
for k in ("PYTHONPATH", "PYTHONHOME"):
    if k in os.environ:
        del os.environ[k]
os.environ["PYTHONNOUSERSITE"] = "1"

# b) Определяем разрешённые корни (только бандл/папка exe)
def abspath(p): 
    try: return os.path.abspath(p)
    except Exception: return p

allowed = []
if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
    allowed.append(abspath(sys._MEIPASS))
exe_dir = abspath(os.path.dirname(sys.executable)) if getattr(sys, "frozen", False) else abspath(os.getcwd())
allowed.append(exe_dir)

# c) Полностью переопределяем sys.path на минимальный набор
sys.path[:] = [p for p in allowed if p]

# d) Блокируем user site
try:
    import site
    site.ENABLE_USER_SITE = False
except Exception:
    pass

# e) Меняем рабочую папку на MEIPASS (чтобы CWD не мешал)
try:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        os.chdir(sys._MEIPASS)
except Exception:
    pass
