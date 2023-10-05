from cx_Freeze import setup, Executable

# Define your main script
main_script = 'gui.py'

# Create an executable
executables = [Executable(main_script, base="Win32GUI")]  # Use "Win32GUI" for Windows

# Additional options (e.g., include files or packages)
options = {
    'build_exe': {
        'packages': ["tkinter","docx","pandas","os","subprocess","re","win32com.client","tkinter.filedialog","main","utils"],
        'include_files': ["Images\logo.png"],
    },
}

setup(
    name='Word To Excel Converter',
    version='2.3',
    author="caphefalumi",
    author_email="dangduytoan13l@gmail.com",
    url="https://github.com/Dangduytoan12l/WordToExcel",
    
    description='Converts questions into Excel to import to Quizizz',
    executables=executables,
    options=options,
)
