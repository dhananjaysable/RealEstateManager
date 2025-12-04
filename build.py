import PyInstaller.__main__
import customtkinter
import os
import sys

# Get customtkinter path to include its data
customtkinter_path = os.path.dirname(customtkinter.__file__)

# Define paths for our scripts
current_dir = os.path.dirname(os.path.abspath(__file__))
res_path = os.path.join(current_dir, 'reslivemain')
manage_path = os.path.join(current_dir, 'resvaduvlive')

print("Building Real Estate Manager...")

PyInstaller.__main__.run([
    'modern_gui_app.py',
    '--name=RealEstateManager',
    '--noconfirm',
    '--onedir',
    '--windowed',
    '--clean',
    # Add customtkinter data
    f'--add-data={customtkinter_path};customtkinter/',
    # Add our script folders
    f'--add-data={res_path};reslivemain/',
    f'--add-data={manage_path};resvaduvlive/',
    # Hidden imports might be needed since we import dynamically
    '--hidden-import=pandas',
    '--hidden-import=openpyxl',
    '--hidden-import=PIL',
    '--hidden-import=tqdm',
])

print("Build complete. Check dist/RealEstateManager folder.")
