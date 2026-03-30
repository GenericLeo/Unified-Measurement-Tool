# Quick Start Guide - Unified Measurement Tool

## Getting Started in 3 Steps

### Step 1: Setup Environment (First Time Only)

Open Terminal and run:

```bash
cd "/Users/leosoler/Layer Measurements/UnifiedMeasurementTool"
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Step 2: Run the Application

#### Option A: Use the launcher script (Recommended)
```bash
./run_app.sh
```

#### Option B: Manual launch
```bash
source .venv/bin/activate
python main_app.py
```

### Step 3: Use the Tool

1. **Click "Choose Image Folder"** - Select your folder with images
2. **Select Measurement Mode** - Choose Vertical or Horizontal
3. **Pick Target Color** - Click "Choose Color" and select the color you want to measure
4. **Adjust Tolerance** - Slide between 5-50 (start with 20)
5. **Click "Process Images"** - Watch the magic happen!
6. **Export Results** - Save to CSV or Excel

## What This Tool Does

This unified tool combines your previous measurement applications into one powerful tool:

### From untitled-2.py:
- Purple color filtering
- Vertical segment detection
- Basic measurements

### From GUI_CM_working_Nov20.py:
- Advanced color picker (any color!)
- Tolerance adjustment
- Professional GUI layout
- CSV and Excel export

### NEW Features:
- **Horizontal measurements** in addition to vertical
- **Statistics viewer** for quick analysis
- **Modular code** - easier to maintain and extend
- **Better organization** - clean folder structure

## Common Use Cases

### Measuring Vertical Layer Thickness
1. Set mode to "Vertical"
2. Choose your layer color
3. Adjust tolerance until you see good filtering in preview
4. Process and export

### Measuring Horizontal Features
1. Set mode to "Horizontal"
2. Same color selection process
3. Different segment orientation

### Batch Processing Multiple Samples
1. Put all images in one folder (can be in subfolders)
2. Process once
3. Export all results to Excel with separate sheets

## Tips for Best Results

### Color Selection
- Click directly on the target color in the color picker
- Or use RGB values if you know them
- Purple (default): RGB(128, 0, 128)

### Tolerance Settings
- **5-10**: Very precise, only exact color matches
- **15-20**: Good for most cases (recommended start)
- **25-35**: Include color variations
- **40-50**: Very inclusive, may catch unwanted colors

### Cropping
- Use "Crop to right half" if you only need to measure one side
- Reduces processing time
- Useful for symmetric samples

## Troubleshooting

### "No module named cv2"
Run: `pip install opencv-python`

### "No images indexed"
- Check that folder contains: .png, .jpg, .jpeg, .tif, .tiff files
- Subfolders are automatically searched

### Segments not detected
- Adjust tolerance higher
- Verify target color is correct
- Check that processed image shows white pixels

## File Structure After Setup

```
UnifiedMeasurementTool/
├── .venv/                  # Virtual environment (created during setup)
├── main_app.py            # Main application
├── image_processor.py     # Image processing
├── measurement_engine.py  # Measurements and export
├── utils.py              # Utilities
├── requirements.txt      # Dependencies
├── README.md            # Full documentation
├── QUICKSTART.md        # This file
├── run_app.sh           # Launcher script
└── __init__.py          # Package file
```

## Next Steps

Once you're comfortable with basic usage:

1. Read the full [README.md](README.md) for advanced features
2. Try both vertical and horizontal measurements
3. Experiment with different tolerance settings
4. Export data and analyze in Excel or other tools

## Need Help?

- Check the full README.md
- Review the code comments in each module
- Test with sample images first

---

**Version 1.0** - Created March 2026
