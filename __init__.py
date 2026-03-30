"""
Unified Measurement Tool Package
A comprehensive image measurement application
"""

__version__ = "1.0.0"
__author__ = "Layer Measurements Project"
__all__ = ['image_processor', 'measurement_engine', 'utils', 'main_app']

import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as messagebox
import tkinter.filedialog as filedialog
import tkinter.colorchooser as colorchooser
import os
import numpy as np
from PIL import Image, ImageTk

class ImageProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("Unified Measurement Tool")
        self.root.geometry("1600x1000")
        
        # Application state
        self.folder_path = ""
        self.images_list = []
        self.processed_images = []
        self.thumbnail_refs = []
        self.selected_color = (128, 0, 128)  # Default purple
        self.measurement_mode = tk.StringVar(value="vertical")
        
        # Build UI
        self.build_ui()

    def build_ui(self):
        """Build the complete user interface"""
        
        # Top frame - Folder selection and settings
        top_frame = tk.Frame(self.root)
        top_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)
        
        self.folder_var = tk.StringVar(value="No folder selected")
        self.images_var = tk.StringVar(value="No images indexed")
        
        tk.Button(top_frame, text="Choose Image Folder", command=self.choose_folder, 
                 width=25, font=("Arial", 10, "bold")).pack(pady=5)
        tk.Label(top_frame, textvariable=self.folder_var, wraplength=600).pack(pady=5)
        tk.Label(top_frame, textvariable=self.images_var).pack(pady=5)
        
        # Middle frame - Main content area
        middle_frame = tk.Frame(self.root)
        middle_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Create the three-column layout
        self.create_image_list(middle_frame)
        self.create_preview_section(middle_frame)
        self.create_segments_section(middle_frame)
        
        # Bottom frame - Controls
        bottom_frame = tk.Frame(self.root)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)
        
        self.create_controls(bottom_frame)

    def create_image_list(self, parent):
        """Create the left panel with original images list"""
        left_frame = tk.Frame(parent)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        tk.Label(left_frame, text="Original Images", font=("Arial", 10, "bold")).pack()
        
        # Listbox with scrollbar
        list_container = tk.Frame(left_frame)
        list_container.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(list_container, orient="vertical")
        self.listbox = tk.Listbox(list_container, width=35, height=20, 
                                   yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.listbox.yview)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Bind double-click to open image
        self.listbox.bind('<Double-Button-1>', self.open_image)
        
        # Processed images status
        tk.Label(left_frame, text="Processed Status", font=("Arial", 10, "bold")).pack(pady=(10,0))
        
        status_container = tk.Frame(left_frame)
        status_container.pack(fill=tk.BOTH, expand=True)
        
        status_scrollbar = tk.Scrollbar(status_container, orient="vertical")
        self.processed_listbox = tk.Listbox(status_container, width=35, height=20, 
                                            yscrollcommand=status_scrollbar.set)
        status_scrollbar.config(command=self.processed_listbox.yview)
        
        status_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.processed_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def create_preview_section(self, parent):
        """Create the middle panel with processed image previews"""
        preview_right_frame = tk.Frame(parent)
        preview_right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        tk.Label(preview_right_frame, text="Preview (Processed)", 
                font=("Arial", 10, "bold")).pack()
        
        # Canvas with scrollbar for thumbnails
        canvas_container = tk.Frame(preview_right_frame)
        canvas_container.pack(fill=tk.BOTH, expand=True)
        
        self.preview_canvas = tk.Canvas(canvas_container, width=300)
        preview_scrollbar = tk.Scrollbar(canvas_container, orient="vertical", 
                                            command=self.preview_canvas.yview)
        self.preview_canvas.configure(yscrollcommand=preview_scrollbar.set)
        
        preview_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.preview_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Frame inside canvas to hold thumbnails
        self.preview_frame = tk.Frame(self.preview_canvas)
        self.preview_canvas_window = self.preview_canvas.create_window(
            (0, 0), window=self.preview_frame, anchor="nw")
        
        # Update scrollregion when frame changes size
        self.preview_frame.bind("<Configure>", self.on_preview_frame_configure)
        self.preview_canvas.bind("<Configure>", self.on_preview_canvas_configure)

    def create_segments_section(self, parent):
        """Create the right panel with segmented regions"""
        right_frame = tk.Frame(parent)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)
        
        tk.Label(right_frame, text="Segmented Regions", 
                font=("Arial", 10, "bold")).pack()
        
        # Canvas with scrollbar for segmented thumbnails
        canvas_container = tk.Frame(right_frame)
        canvas_container.pack(fill=tk.BOTH, expand=True)
        
        self.segment_canvas = tk.Canvas(canvas_container, width=300)
        segment_scrollbar = tk.Scrollbar(canvas_container, orient="vertical", 
                                            command=self.segment_canvas.yview)
        self.segment_canvas.configure(yscrollcommand=segment_scrollbar.set)
        
        segment_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.segment_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Frame inside canvas to hold segmented thumbnails
        self.segment_frame = tk.Frame(self.segment_canvas)
        self.segment_canvas_window = self.segment_canvas.create_window(
            (0, 0), window=self.segment_frame, anchor="nw")
        
        # Update scrollregion when frame changes size
        self.segment_frame.bind("<Configure>", self.on_segment_frame_configure)
        self.segment_canvas.bind("<Configure>", self.on_segment_canvas_configure)

    def create_controls(self, parent):
        """Create the control panel at the bottom"""
        
        # Measurement mode selection
        mode_frame = tk.LabelFrame(parent, text="Measurement Mode", padx=10, pady=10)
        mode_frame.pack(pady=5, fill=tk.X)
        
        tk.Radiobutton(mode_frame, text="Vertical Measurements", 
                        variable=self.measurement_mode, value="vertical").pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(mode_frame, text="Horizontal Measurements", 
                        variable=self.measurement_mode, value="horizontal").pack(side=tk.LEFT, padx=10)
        
        # Color selection
        color_frame = tk.LabelFrame(parent, text="Color Filtering", padx=10, pady=10)
        color_frame.pack(pady=5, fill=tk.X)
        
        # Color display
        self.color_display = tk.Frame(color_frame, width=30, height=30, 
                                      bg=rgb_to_hex(self.selected_color))
        self.color_display.pack(side=tk.LEFT, padx=5)
        
        tk.Button(color_frame, text="Choose Color", command=self.choose_color, 
                 width=15).pack(side=tk.LEFT, padx=5)
        
        self.color_label = tk.Label(color_frame, text=f"Target Color: RGB{self.selected_color}")
        self.color_label.pack(side=tk.LEFT, padx=5)
        
        # Tolerance setting
        tk.Label(color_frame, text="Tolerance:").pack(side=tk.LEFT, padx=5)
        self.tolerance_var = tk.IntVar(value=20)
        tolerance_scale = tk.Scale(color_frame, from_=5, to=50, orient=tk.HORIZONTAL, 
                                    variable=self.tolerance_var, length=150)
        tolerance_scale.pack(side=tk.LEFT, padx=5)
        
        # Processing options
        options_frame = tk.LabelFrame(parent, text="Processing Options", padx=10, pady=10)
        options_frame.pack(pady=5, fill=tk.X)
        
        self.crop_var = tk.BooleanVar(value=False)
        tk.Checkbutton(options_frame, text="Crop to right half only", 
                        variable=self.crop_var).pack(side=tk.LEFT, padx=10)
        
        # Action buttons
        button_frame = tk.Frame(parent)
        button_frame.pack(pady=5)
        
        tk.Button(button_frame, text="Process Images", command=self.process_all, 
                 width=20, bg="#4CAF50", fg="white", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Save to CSV", command=self.save_to_csv, 
                 width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Save to Excel", command=self.save_to_excel, 
                 width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="View Statistics", command=self.show_statistics, 
                 width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Close", command=self.root.destroy, 
                 width=10).pack(side=tk.LEFT, padx=5)

# Event handlers

def choose_folder(self):
    """Handle folder selection"""
    path = select_folder()
    if not path:
        return
    
    self.folder_path = path
    self.folder_var.set(path)
    
    # Find all images
    self.images_list = find_images_recursively(path)
    self.images_var.set(f"Indexed {len(self.images_list)} images")
    
    # Update listbox
    self.listbox.delete(0, tk.END)
    for img in self.images_list:
        self.listbox.insert(tk.END, img)
    
    # Clear processed list and previews
    self.processed_listbox.delete(0, tk.END)
    self.clear_previews()

def choose_color(self):
    """Handle color picker"""
    initial_hex = rgb_to_hex(self.selected_color)
    color = colorchooser.askcolor(
        title="Choose target color to filter",
        initialcolor=initial_hex
    )
    if color[0]:  # color[0] is RGB tuple
        self.selected_color = tuple(int(c) for c in color[0])
        hex_color = rgb_to_hex(self.selected_color)
        self.color_display.config(bg=hex_color)
        self.color_label.config(text=f"Target Color: RGB{self.selected_color}")

def open_image(self, event):
    """Handle double-click on image to open it"""
    selection = self.listbox.curselection()
    if not selection:
        return
    img_name = self.listbox.get(selection[0])
    if not os.path.isdir(self.folder_path):
        return
    img_path = os.path.join(self.folder_path, img_name)
    if os.path.exists(img_path):
        os.system(f'open "{img_path}"')

def process_all(self):
    """Process all images with current settings"""
    if not os.path.isdir(self.folder_path):
        messagebox.showerror("Error", "No valid folder selected.")
        return
    
    # Process with current settings
    self.processed_images = process_images(
        self.folder_path,
        self.images_list,
        crop_right_half=self.crop_var.get(),
        target_color=self.selected_color,
        tolerance=self.tolerance_var.get()
    )
    
    # Update processed listbox
    self.processed_listbox.delete(0, tk.END)
    success_count = 0
    for i, (img_name, bw) in enumerate(zip(self.images_list, self.processed_images)):
        if bw is not None:
            self.processed_listbox.insert(tk.END, f"✓ {img_name}")
            self.processed_listbox.itemconfig(i, fg="green")
            success_count += 1
        else:
            self.processed_listbox.insert(tk.END, f"✗ {img_name}")
            self.processed_listbox.itemconfig(i, fg="red")
    
    self.images_var.set(f"Processed {success_count}/{len(self.processed_images)} images successfully.")
    
    # Generate previews
    self.generate_previews()

def generate_previews(self):
    """Generate thumbnail previews for processed images"""
    self.clear_previews()
    
    orientation = self.measurement_mode.get()
    
    for i, (img_name, bw) in enumerate(zip(self.images_list, self.processed_images)):
        if bw is not None:
            # Create processed thumbnail
            thumb_frame = tk.Frame(self.preview_frame, relief=tk.RIDGE, borderwidth=2)
            thumb_frame.pack(fill=tk.X, padx=5, pady=5)
            
            tk.Label(thumb_frame, text=img_name, wraplength=280, 
                    font=("Arial", 8)).pack(pady=2)
            
            pil = Image.fromarray(bw)
            pil.thumbnail((280, 280))
            tkimg = ImageTk.PhotoImage(pil)
            img_label = tk.Label(thumb_frame, image=tkimg)
            img_label.image = tkimg  # type: ignore  # Keep reference to prevent GC
            img_label.pack(pady=2)
            
            self.thumbnail_refs.append(tkimg)
            
            # Create segmented version
            segmented_rgb = draw_segments_on_image(bw, orientation=orientation)
            
            seg_thumb_frame = tk.Frame(self.segment_frame, relief=tk.RIDGE, borderwidth=2)
            seg_thumb_frame.pack(fill=tk.X, padx=5, pady=5)
            
            tk.Label(seg_thumb_frame, text=img_name, wraplength=280, 
                    font=("Arial", 8)).pack(pady=2)
            
            seg_pil = Image.fromarray(segmented_rgb)
            seg_pil.thumbnail((280, 280))
            seg_tkimg = ImageTk.PhotoImage(seg_pil)
            seg_img_label = tk.Label(seg_thumb_frame, image=seg_tkimg)
            seg_img_label.image = seg_tkimg  # type: ignore  # Keep reference to prevent GC
            seg_img_label.pack(pady=2)
            
            self.thumbnail_refs.append(seg_tkimg)

def clear_previews(self):
    """Clear all preview thumbnails"""
    for widget in self.preview_frame.winfo_children():
        widget.destroy()
    for widget in self.segment_frame.winfo_children():
        widget.destroy()
    self.thumbnail_refs.clear()

def save_to_csv(self):
    """Save measurements to CSV file"""
    if not self.processed_images:
        messagebox.showwarning("No Data", "Please process images first.")
        return
    
    csv_path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        title="Save Measurements"
    )
    
    if not csv_path:
        return
    
    try:
        if self.measurement_mode.get() == "vertical":
            save_vertical_segments_to_csv(self.images_list, self.processed_images, csv_path)
        else:
            save_horizontal_segments_to_csv(self.images_list, self.processed_images, csv_path)
        
        messagebox.showinfo("Success", f"Measurements saved to:\n{csv_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save CSV:\n{str(e)}")

def save_to_excel(self):
    """Save measurements to Excel file"""
    if not self.processed_images:
        messagebox.showwarning("No Data", "Please process images first.")
        return
    
    excel_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        title="Save Measurements"
    )
    
    if not excel_path:
        return
    
    try:
        if self.measurement_mode.get() == "vertical":
            save_vertical_segments_to_excel(self.images_list, self.processed_images, excel_path)
        else:
            save_horizontal_segments_to_excel(self.images_list, self.processed_images, excel_path)
        
        messagebox.showinfo("Success", f"Measurements saved to:\n{excel_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save Excel:\n{str(e)}")

def show_statistics(self):
    """Show statistics for the current measurements"""
    if not self.processed_images:
        messagebox.showwarning("No Data", "Please process images first.")
        return
    
    # Calculate statistics for all images
    total_stats = {
        'count': 0, 'min': float('inf'), 'max': 0, 
        'mean': 0, 'median': 0, 'std_dev': 0
    }
    all_lengths = []
    
    for bw in self.processed_images:
        if bw is not None:
            if self.measurement_mode.get() == "vertical":
                segments = analyze_all_vertical_segments(bw)
            else:
                segments = analyze_all_horizontal_segments(bw)
            
            for coord, seg_list in segments.items():
                for start, end in seg_list:
                    length = end - start + 1
                    all_lengths.append(length)
    
    if all_lengths:
        total_stats['count'] = len(all_lengths)
        total_stats['min'] = min(all_lengths)
        total_stats['max'] = max(all_lengths)
        total_stats['mean'] = np.mean(all_lengths)
        total_stats['median'] = np.median(all_lengths)
        total_stats['std_dev'] = np.std(all_lengths)
    
    # Display statistics in a message box
    stats_text = f"""Measurement Statistics ({self.measurement_mode.get().capitalize()})
```
<userPrompt>
Provide the fully rewritten file, incorporating the suggested code change. You must produce the complete file.
</userPrompt>
