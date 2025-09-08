#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
键盘输入处理模块
支持跨平台的单字符输入检测
"""

import sys
import termios
import tty

class _Getch:
    """Gets a single character from stdin."""
    
    def __init__(self):
        try:
            self.impl = _GetchUnix()
        except ImportError:
            self.impl = _GetchWindows()
    
    def __call__(self):
        return self.impl()

class _GetchUnix:
    def __init__(self):
        pass
    
    def __call__(self):
        import sys, tty, termios
        fd = sys.stdin.fileno()
        old_settings = termios.tcgetattr(fd)
        try:
            tty.setcbreak(fd)
            ch = sys.stdin.read(1)
        finally:
            termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)
        return ch

class _GetchWindows:
    def __init__(self):
        import msvcrt
    
    def __call__(self):
        import msvcrt
        return msvcrt.getch().decode('utf-8')

getch = _Getch()