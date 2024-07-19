import tkinter as tk
from tkinter import ttk

class PlaceholderTab:
    def __init__(self, parent, tab_name):
        self.frame = ttk.Frame(parent)
        label = ttk.Label(self.frame, text=f"Content for {tab_name}")
        label.pack(pady=20)
