#!/usr/bin/env python3
"""
Python example to use the Rust dynamic library for Excel sheet classification
"""

import ctypes
import json
import os
from ctypes import c_char_p


def load_rust_library():
    """
    Load the Rust dynamic library and set up function signatures
    """
    # Add the library directory to the system library path
    lib_dir = os.path.join(
        os.path.dirname(__file__), 
        "target", "release"
    )
    
    # On Linux, we may need to modify LD_LIBRARY_PATH
    if os.name == 'posix':
        current_ld_path = os.environ.get('LD_LIBRARY_PATH', '')
        new_ld_path = f"{lib_dir}:{current_ld_path}" if current_ld_path else lib_dir
        # Note: setting LD_LIBRARY_PATH after the process starts has no effect on dlopen
        # The environment variable must be set before the Python process starts
    
    # Path to the dynamic library
    lib_path = os.path.join(lib_dir, "liblayout_view.so")
    
    if not os.path.exists(lib_path):
        raise FileNotFoundError(f"Library not found: {lib_path}")
    
    # Load the library
    lib = ctypes.CDLL(lib_path)
    
    # Define function signatures
    lib.classify_excel_sheets_c.argtypes = [c_char_p]
    lib.classify_excel_sheets_c.restype = ctypes.POINTER(ctypes.c_char)  # Return pointer to char
    
    lib.free_c_string.argtypes = [ctypes.POINTER(ctypes.c_char)]
    lib.free_c_string.restype = None
    
    return lib


def classify_excel_sheets(xlsx_path):
    """
    Classify Excel sheets using the Rust library
    
    Args:
        xlsx_path (str): Path to the Excel file
    
    Returns:
        list: List of classified sheets or None if error occurred
    """
    lib = load_rust_library()
    
    # Convert Python string to C string
    c_path = ctypes.c_char_p(xlsx_path.encode('utf-8'))
    
    # Call the Rust function
    result_ptr = lib.classify_excel_sheets_c(c_path)
    
    if not result_ptr:
        return None
    
    try:
        # Convert the result to Python string
        # Use ctypes.c_char_p(result_ptr).value to get the string value
        result_bytes = ctypes.cast(result_ptr, ctypes.c_char_p).value
        if result_bytes is None:
            return None
        
        result_str = result_bytes.decode('utf-8')
        
        # Parse the JSON result
        parsed_result = json.loads(result_str)
        
        # Free the allocated string in Rust
        lib.free_c_string(result_ptr)
        
        return parsed_result
    except json.JSONDecodeError:
        # Even if JSON parsing fails, we still need to free the string
        lib.free_c_string(result_ptr)
        return None
    except Exception:
        # In case of any other error, free the string and re-raise
        lib.free_c_string(result_ptr)
        raise


def main():
    """
    Main function to demonstrate the usage
    """
    import sys
    
    if len(sys.argv) != 2:
        print("Usage: python use_layout_view.py <xlsx_file_path>")
        sys.exit(1)
    
    xlsx_path = sys.argv[1]
    
    if not os.path.exists(xlsx_path):
        print(f"Error: File does not exist: {xlsx_path}")
        sys.exit(1)
    
    print(f"Classifying Excel file: {xlsx_path}")
    
    results = classify_excel_sheets(xlsx_path)
    
    if results is None:
        print("Error: Failed to classify Excel sheets")
        sys.exit(1)
    
    print("Classification results:")
    print(json.dumps(results, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()