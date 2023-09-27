import sys
import tkinter as tk

class Redirector:
    def __init__(self, widget):
        self.widget = widget

    def write(self, message):
        self.widget.insert('end', message)