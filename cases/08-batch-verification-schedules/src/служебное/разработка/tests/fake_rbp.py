"""Подставной Revit Batch Processor для тестов запускающих скриптов. Запускается и на Linux (скрипт с shebang), и на Windows:
там файл без PE-заголовка запустить нельзя («system cannot find all the information required», «not a valid Win32 application»),
поэтому рядом создаётся .cmd, который вызывает Python. Передавать в -RbpExe нужно значение, которое вернул install()."""
import os
import stat
import sys


def install(folder, name, python_source):
    """Создаёт подставной RBP из Python-кода; возвращает путь для -RbpExe."""
    if os.name == "nt":
        impl = os.path.join(folder, name + "_impl.py")
        with open(impl, "w", encoding="utf-8", newline="") as f:
            f.write(python_source)
        cmd = os.path.join(folder, name + ".cmd")
        with open(cmd, "w", encoding="ascii", newline="") as f:
            f.write('@echo off\r\n"%s" "%s" %%*\r\n' % (sys.executable, impl))
        return cmd
    path = os.path.join(folder, name + ".exe")
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write("#!/usr/bin/env python3\n" + python_source)
    os.chmod(path, os.stat(path).st_mode | stat.S_IEXEC)
    return path
