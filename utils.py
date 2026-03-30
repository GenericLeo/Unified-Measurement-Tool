"""
Utility Functions Module
Handles file operations, folder selection, and image discovery
"""

import os
import re
import tkinter as tk
from tkinter import filedialog
from typing import Dict, List, Optional, Tuple
import cv2


def select_folder(title: str = "Select folder") -> str:
    """
    Open a folder selection dialog.
    
    Args:
        title: Dialog title
    
    Returns:
        Selected folder path or empty string if cancelled
    """
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title=title)
    root.destroy()
    return folder_path


def select_file(title: str = "Select file", filetypes: Optional[List[Tuple[str, str]]] = None) -> str:
    """Open a file selection dialog and return the selected path."""
    root = tk.Tk()
    root.withdraw()
    if filetypes is None:
        filetypes = [("All files", "*.*")]
    file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    root.destroy()
    return file_path


def find_images_recursively(folder_path: str) -> List[str]:
    """
    Find all image files in a folder and its subdirectories.
    
    Args:
        folder_path: Root folder to search
    
    Returns:
        List of relative paths to image files
    """
    image_exts = ('.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp')
    found = []
    for root, _dirs, files in os.walk(folder_path):
        for f in files:
            if f.lower().endswith(image_exts):
                found.append(os.path.relpath(os.path.join(root, f), folder_path))
    return found


def categorize_data_files(folder_path: str, recursive: bool = True) -> Tuple[List[str], List[str]]:
    """Scan a folder for TIFF and TXT files, returning relative paths."""
    tiff_files: List[str] = []
    txt_files: List[str] = []

    if recursive:
        for root, _dirs, files in os.walk(folder_path):
            for name in files:
                lower = name.lower()
                rel = os.path.relpath(os.path.join(root, name), folder_path)
                if lower.endswith((".tif", ".tiff")):
                    tiff_files.append(rel)
                elif lower.endswith(".txt"):
                    txt_files.append(rel)
    else:
        for name in os.listdir(folder_path):
            full_path = os.path.join(folder_path, name)
            if not os.path.isfile(full_path):
                continue
            lower = name.lower()
            if lower.endswith((".tif", ".tiff")):
                tiff_files.append(name)
            elif lower.endswith(".txt"):
                txt_files.append(name)

    return sorted(tiff_files), sorted(txt_files)


def categorize_mask_files(folder_path: str) -> List[str]:
    """Scan a folder for mask image files."""
    files: List[str] = []
    for name in os.listdir(folder_path):
        full_path = os.path.join(folder_path, name)
        if not os.path.isfile(full_path):
            continue
        if name.lower().endswith((".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp")):
            files.append(name)
    return sorted(files)


def get_pixel_size(txt_file_path: str) -> Tuple[Optional[float], Optional[float]]:
    """Extract PixelSize from metadata TXT and return (nm_per_px, um_per_px)."""
    try:
        content: Optional[str] = None
        for encoding in ["utf-16-le", "utf-16", "utf-8", "latin-1"]:
            try:
                with open(txt_file_path, "r", encoding=encoding, errors="ignore") as handle:
                    maybe = handle.read()
                if "PixelSize" in maybe:
                    content = maybe
                    break
                if content is None:
                    content = maybe
            except (UnicodeError, UnicodeDecodeError):
                continue

        if not content:
            return None, None

        match = re.search(r"PixelSize=([0-9.]+)", content)
        if not match:
            return None, None

        nm_value = float(match.group(1))
        return nm_value, nm_value / 1000.0
    except Exception:
        return None, None


def build_pixel_size_map(base_folder: str, txt_rel_paths: List[str]) -> Dict[str, Tuple[Optional[float], Optional[float]]]:
    """Build mapping: txt relative path -> (nm_per_px, um_per_px)."""
    mapping: Dict[str, Tuple[Optional[float], Optional[float]]] = {}
    for rel in txt_rel_paths:
        full_path = os.path.join(base_folder, rel)
        mapping[rel] = get_pixel_size(full_path)
    return mapping


def process_images(folder: str, images: List[str], crop_right_half: bool = False,
                   crop_left_half: bool = False,
                   target_color: tuple = (128, 0, 128), tolerance: int = 20) -> List:
    """
    Process a list of images by applying color filtering.
    
    Args:
        folder: Base folder containing images
        images: List of relative image paths
        crop_right_half: Whether to crop to right half before processing
        crop_left_half: Whether to crop to left half before processing
        target_color: RGB color to filter
        tolerance: Color matching tolerance
    
    Returns:
        List of processed binary images (or None for failed images)
    """
    from image_processor import color_to_bw, crop_image
    
    processed = []
    for rel in images:
        path = os.path.join(folder, rel)
        img = cv2.imread(path, cv2.IMREAD_COLOR)
        if img is None:
            processed.append(None)
            continue
        
        # Crop if requested
        img = crop_image(
            img,
            crop_right_half=crop_right_half,
            crop_left_half=crop_left_half,
        )
        
        # Apply color filtering
        bw = color_to_bw(img, target_color, tolerance)
        processed.append(bw)
    
    return processed


def open_file_in_system(file_path: str) -> None:
    """
    Open a file using the system's default application.
    
    Args:
        file_path: Path to the file to open
    """
    if os.path.exists(file_path):
        os.system(f'open "{file_path}"')  # macOS
        # For cross-platform, could use: import subprocess; subprocess.run(['open', file_path])


def rgb_to_hex(rgb: tuple) -> str:
    """
    Convert RGB tuple to hex color string.
    
    Args:
        rgb: Tuple of (r, g, b) values
    
    Returns:
        Hex color string (e.g., "#ff00ff")
    """
    return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
