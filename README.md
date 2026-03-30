# Unified Measurement Tool

A comprehensive image measurement application that combines vertical and horizontal measurement capabilities with customizable color filtering.

## Features

### Measurement Capabilities
- **Vertical Measurements**: Analyze vertical segments in images (column-by-column analysis)
- **Horizontal Measurements**: Analyze horizontal segments in images (row-by-row analysis)
- **Batch Processing**: Process multiple images simultaneously
- **Customizable Color Filtering**: Select any color to filter and measure
- **Adjustable Tolerance**: Fine-tune color matching sensitivity

### Processing Options
- Crop images (right half) before processing
- Visual previews of processed images
- Highlighted segment visualization
- Real-time statistics

### Export Capabilities
- **CSV Export**: Save measurements with detailed segment information
- **Excel Export**: Create separate sheets for each image
- **Statistics View**: View aggregate statistics for all measurements

## Installation

### Prerequisites
- Python 3.8 or higher
- Virtual environment (recommended)

### Setup

1. Navigate to the UnifiedMeasurementTool folder:
   ```bash
   cd "UnifiedMeasurementTool"
   ```

2. Create a virtual environment (if not already created):
   ```bash
   python -m venv .venv
   ```

3. Activate the virtual environment:
   - macOS/Linux:
     ```bash
     source .venv/bin/activate
     ```
   - Windows:
     ```
     .venv\Scripts\activate
     ```

4. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Running the Application

```bash
python main_app.py
```

Or use the provided launcher script:
```bash
./run_app.sh
```

## Releases

This repository includes a GitHub Actions release workflow that builds platform-specific installers for:
- Windows: `.exe`
- macOS: `.dmg`

### How to Publish a Release

1. Update [version.py](version.py) with the version you want to release.
2. Commit and push your changes to `main`.
3. Create and push a matching tag:

```bash
git tag v0.1.0
git push origin v0.1.0
```

4. GitHub Actions will automatically:
   - build the Windows executable
   - build the macOS application and package it as a DMG
   - upload both assets to the GitHub release for that tag

### Expected Release Asset Names

- `Unified-Measurement-Tool-0.1.0-Windows.exe`
- `Unified-Measurement-Tool-0.1.0-macOS.dmg`

These names match the updater logic used by the app.

### Workflow

1. **Choose Image Folder**: Click the button to select a folder containing your images
2. **Configure Settings**:
   - Select measurement mode (Vertical or Horizontal)
   - Choose target color to filter using the color picker
   - Adjust color tolerance (5-50)
   - Enable cropping if needed
3. **Process Images**: Click "Process Images" to analyze all images
4. **View Results**: 
   - Review processed images in the preview panel
   - View highlighted segments in the segments panel
5. **Export Data**:
   - Save to CSV for detailed segment data
   - Save to Excel for organized multi-sheet reports
   - View Statistics for aggregate measurements

## Project Structure

```
UnifiedMeasurementTool/
├── main_app.py              # Main GUI application
├── image_processor.py       # Image processing and segment detection
├── measurement_engine.py    # Measurement calculations and export
├── utils.py                 # Utility functions
├── requirements.txt         # Python dependencies
├── README.md               # This file
└── run_app.sh              # Launcher script (macOS)
```

## Module Documentation

### image_processor.py
- Color filtering (RGB to binary conversion)
- Vertical and horizontal segment detection
- Image cropping and transformations
- Segment visualization

### measurement_engine.py
- Segment statistics calculations
- CSV export functionality
- Excel export with multi-sheet support
- Physical measurement conversion (future)

### utils.py
- File and folder operations
- Image discovery and loading
- Batch image processing
- System integration utilities

## Features Comparison

This unified tool combines features from:
- **untitled-2.py**: Core purple color filtering and vertical segment analysis
- **GUI_CM_working_Nov20.py**: Advanced color picker, tolerance control, and enhanced UI

### New Features
- Horizontal measurement mode
- Cleaner modular architecture
- Enhanced statistics display
- Improved error handling
- Better code organization

## Tips

- Use **higher tolerance** (30-50) for colors with variation
- Use **lower tolerance** (5-15) for precise color matching
- Double-click images in the list to open them in system viewer
- Process smaller batches first to test settings
- Save frequently to avoid data loss

## Troubleshooting

### Images Not Loading
- Ensure the folder contains supported formats: PNG, JPG, JPEG, TIF, TIFF, BMP
- Check file permissions

### Color Not Filtering Correctly
- Adjust tolerance setting
- Try different color selection if auto-detection fails
- Ensure image has the target color visible

### Export Errors
- Ensure you have write permissions for the target folder
- Check that Excel/CSV file is not open in another program

## Future Enhancements

- [ ] Scale factor input for physical unit conversion
- [ ] Area measurements in addition to linear
- [ ] Custom export templates
- [ ] Image annotation tools
- [ ] Measurement comparison tools
- [ ] Batch export with automatic naming

## Requirements

See [requirements.txt](requirements.txt) for full dependency list.

Core dependencies:
- opencv-python (image processing)
- numpy (numerical operations)
- Pillow (image handling)
- openpyxl (Excel export)
- tkinter (GUI - included with Python)

## License

For internal use - Layer Measurements Project

## Version

Version 1.0 - March 2026
