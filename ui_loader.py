"""
ui_loader.py
Loads the main window from the .ui file using PyQt5's uic module.
"""
from PyQt5 import uic, QtWidgets
import os

def load_main_window():
    ui_path = os.path.join(os.path.dirname(__file__), 'main_window_5.ui')
    return uic.loadUi(ui_path)
