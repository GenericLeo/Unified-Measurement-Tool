"""
Image Processing Module for Unified Measurement Tool
Handles color filtering, image transformations, and segment detection
"""

import cv2
import numpy as np
from math import pi, sqrt
from typing import Dict, List, Optional, Tuple


def rgb_to_hsv_range(rgb_color: Tuple[int, int, int], tolerance: int = 20) -> Tuple[np.ndarray, np.ndarray]:
    """
    Convert RGB color to HSV range for color filtering.
    
    Args:
        rgb_color: RGB color tuple (r, g, b) where values are 0-255
        tolerance: Tolerance for hue range (default 20)
    
    Returns:
        Tuple of (lower_hsv, upper_hsv) arrays for cv2.inRange
    """
    # Convert RGB to BGR for OpenCV
    r, g, b = rgb_color
    bgr_array = np.array([[[b, g, r]]], dtype=np.uint8)
    hsv_color = cv2.cvtColor(bgr_array, cv2.COLOR_BGR2HSV)[0][0]
    
    # Create range around the target hue
    hue = hsv_color[0]
    lower_hsv = np.array([max(0, hue - tolerance), 50, 50], dtype=np.uint8)
    upper_hsv = np.array([min(179, hue + tolerance), 255, 255], dtype=np.uint8)
    
    return lower_hsv, upper_hsv


def color_to_bw(img_bgr: np.ndarray, target_color: Tuple[int, int, int] = (128, 0, 128), 
                tolerance: int = 20) -> np.ndarray:
    """
    Convert pixels of a specific color to white, everything else to black.
    
    Args:
        img_bgr: Input image in BGR format
        target_color: RGB color to filter (default purple)
        tolerance: Color tolerance for matching (default 20)
    
    Returns:
        Binary image with target color as white pixels
    """
    # Convert to HSV for better color filtering
    hsv = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2HSV)
    
    # Get HSV range for the target color
    lower_hsv, upper_hsv = rgb_to_hsv_range(target_color, tolerance)
    
    # Create mask for the target color
    mask_hsv = cv2.inRange(hsv, lower_hsv, upper_hsv)
    
    # Additional RGB-based filtering for better accuracy
    b, g, r = cv2.split(img_bgr)
    target_r, target_g, target_b = target_color
    
    # Create RGB tolerance mask
    r_mask = np.abs(r.astype(int) - target_r) <= tolerance
    g_mask = np.abs(g.astype(int) - target_g) <= tolerance
    b_mask = np.abs(b.astype(int) - target_b) <= tolerance
    rgb_mask = (r_mask & g_mask & b_mask).astype(np.uint8) * 255
    
    # Combine HSV and RGB masks
    combined = cv2.bitwise_or(mask_hsv, rgb_mask)
    
    # Apply morphological operations to clean up the mask
    kernel = np.ones((3, 3), np.uint8)
    combined = cv2.morphologyEx(combined, cv2.MORPH_OPEN, kernel, iterations=1)
    combined = cv2.morphologyEx(combined, cv2.MORPH_CLOSE, kernel, iterations=1)
    
    return combined


def purple_to_bw(img_bgr: np.ndarray) -> np.ndarray:
    """
    Convert purple pixels to white, everything else to black.
    Legacy function for backward compatibility.
    
    Args:
        img_bgr: Input image in BGR format
    
    Returns:
        Binary image with purple pixels as white
    """
    return color_to_bw(img_bgr, target_color=(128, 0, 128), tolerance=20)


def crop_image(img: np.ndarray, crop_right_half: bool = False, 
               crop_left_half: bool = False) -> np.ndarray:
    """
    Crop image based on specified options.
    
    Args:
        img: Input image
        crop_right_half: Keep only right half
        crop_left_half: Keep only left half
    
    Returns:
        Cropped image
    """
    if crop_right_half:
        h, w = img.shape[:2]
        return img[:, w//2:]
    elif crop_left_half:
        h, w = img.shape[:2]
        return img[:, :w//2]
    return img


def find_vertical_segments(bw_image: np.ndarray, x: int) -> List[Tuple[int, int]]:
    """
    Find all vertical segments of white pixels in a specific column.
    
    Args:
        bw_image: Binary image (white pixels = 255, black = 0)
        x: The column (x-coordinate) to analyze
        
    Returns:
        List of tuples (start_y, end_y) representing segments of white pixels
    """
    if x < 0 or x >= bw_image.shape[1]:
        return []
    
    column = bw_image[:, x]
    segments = []
    in_segment = False
    start_y = None
    
    for y in range(len(column)):
        if column[y] == 255:  # White pixel
            if not in_segment:
                in_segment = True
                start_y = y
        else:  # Black pixel
            if in_segment:
                segments.append((start_y, y - 1))
                in_segment = False
                start_y = None
    
    # Handle case where segment extends to bottom of image
    if in_segment:
        segments.append((start_y, len(column) - 1))
    
    return segments


def find_horizontal_segments(bw_image: np.ndarray, y: int) -> List[Tuple[int, int]]:
    """
    Find all horizontal segments of white pixels in a specific row.
    
    Args:
        bw_image: Binary image (white pixels = 255, black = 0)
        y: The row (y-coordinate) to analyze
        
    Returns:
        List of tuples (start_x, end_x) representing segments of white pixels
    """
    if y < 0 or y >= bw_image.shape[0]:
        return []
    
    row = bw_image[y, :]
    segments = []
    in_segment = False
    start_x = None
    
    for x in range(len(row)):
        if row[x] == 255:  # White pixel
            if not in_segment:
                in_segment = True
                start_x = x
        else:  # Black pixel
            if in_segment:
                segments.append((start_x, x - 1))
                in_segment = False
                start_x = None
    
    # Handle case where segment extends to right edge
    if in_segment:
        segments.append((start_x, len(row) - 1))
    
    return segments


def analyze_all_vertical_segments(bw_image: np.ndarray) -> dict:
    """
    Analyze all columns in the image and find vertical white segments.
    
    Args:
        bw_image: Binary image (white pixels = 255, black = 0)
        
    Returns:
        Dictionary mapping x-coordinates to list of segments
        Example: {0: [(10, 20), (50, 60)], 1: [(15, 25)], ...}
    """
    width = bw_image.shape[1]
    all_segments = {}
    
    for x in range(width):
        segments = find_vertical_segments(bw_image, x)
        if segments:  # Only store columns that have segments
            all_segments[x] = segments
    
    return all_segments


def analyze_all_horizontal_segments(bw_image: np.ndarray) -> dict:
    """
    Analyze all rows in the image and find horizontal white segments.
    
    Args:
        bw_image: Binary image (white pixels = 255, black = 0)
        
    Returns:
        Dictionary mapping y-coordinates to list of segments
        Example: {0: [(10, 20), (50, 60)], 1: [(15, 25)], ...}
    """
    height = bw_image.shape[0]
    all_segments = {}
    
    for y in range(height):
        segments = find_horizontal_segments(bw_image, y)
        if segments:  # Only store rows that have segments
            all_segments[y] = segments
    
    return all_segments


def draw_segments_on_image(bw_image: np.ndarray, orientation: str = 'vertical') -> np.ndarray:
    """
    Draw red highlights on identified segments.
    
    Args:
        bw_image: Binary image (white pixels = 255, black = 0)
        orientation: 'vertical' or 'horizontal' - direction to analyze segments
        
    Returns:
        RGB image with segments highlighted in red
    """
    # Convert grayscale to RGB
    rgb_image = cv2.cvtColor(bw_image, cv2.COLOR_GRAY2RGB)
    
    if orientation == 'vertical':
        # Find all vertical segments
        all_segments = analyze_all_vertical_segments(bw_image)
        
        # Draw red lines on segments
        for x, segments in all_segments.items():
            for start_y, end_y in segments:
                rgb_image[start_y:end_y+1, x] = [255, 0, 0]  # Red in RGB
    
    elif orientation == 'horizontal':
        # Find all horizontal segments
        all_segments = analyze_all_horizontal_segments(bw_image)
        
        # Draw red lines on segments
        for y, segments in all_segments.items():
            for start_x, end_x in segments:
                rgb_image[y, start_x:end_x+1] = [255, 0, 0]  # Red in RGB
    
    return rgb_image


def _get_named_color_ranges() -> Dict[str, List[Tuple[np.ndarray, np.ndarray]]]:
    """Return HSV ranges for supported named colors."""
    return {
        "red": [
            (np.array([0, 40, 40]), np.array([10, 255, 255])),
            (np.array([170, 40, 40]), np.array([180, 255, 255])),
        ],
        "green": [(np.array([35, 40, 40]), np.array([85, 255, 255]))],
        "blue": [(np.array([100, 40, 40]), np.array([130, 255, 255]))],
        "yellow": [(np.array([20, 40, 40]), np.array([35, 255, 255]))],
        "cyan": [(np.array([85, 40, 40]), np.array([100, 255, 255]))],
        "magenta": [(np.array([140, 40, 40]), np.array([170, 255, 255]))],
        "orange": [(np.array([10, 40, 40]), np.array([20, 255, 255]))],
        "purple": [(np.array([130, 40, 40]), np.array([140, 255, 255]))],
    }


def color_name_to_bw_mask(color_bgr: np.ndarray, target_color: str = "green") -> np.ndarray:
    """Convert named-color regions in a mask image to white pixels."""
    hsv = cv2.cvtColor(color_bgr, cv2.COLOR_BGR2HSV)
    ranges = _get_named_color_ranges()

    target = target_color.strip().lower()
    if target not in ranges:
        target = "green"

    mask_hsv = np.zeros(hsv.shape[:2], dtype=np.uint8)
    for lower, upper in ranges[target]:
        mask_hsv = cv2.bitwise_or(mask_hsv, cv2.inRange(hsv, lower, upper))

    kernel = np.ones((3, 3), np.uint8)
    mask_hsv = cv2.morphologyEx(mask_hsv, cv2.MORPH_OPEN, kernel, iterations=1)
    mask_hsv = cv2.morphologyEx(mask_hsv, cv2.MORPH_CLOSE, kernel, iterations=1)
    return mask_hsv


def detect_baseline_region(color_bgr: np.ndarray, baseline_color: str = "auto") -> Tuple[Optional[np.ndarray], Optional[np.ndarray]]:
    """Detect baseline/interface curve and return (y-array per x, baseline mask)."""
    hsv = cv2.cvtColor(color_bgr, cv2.COLOR_BGR2HSV)
    ranges = _get_named_color_ranges()

    target = baseline_color.strip().lower()
    baseline_mask = np.zeros(hsv.shape[:2], dtype=np.uint8)

    if target == "auto":
        for color_ranges in ranges.values():
            for lower, upper in color_ranges:
                baseline_mask = cv2.bitwise_or(baseline_mask, cv2.inRange(hsv, lower, upper))
    elif target in ranges:
        for lower, upper in ranges[target]:
            baseline_mask = cv2.bitwise_or(baseline_mask, cv2.inRange(hsv, lower, upper))
    else:
        return None, None

    kernel = np.ones((5, 5), np.uint8)
    baseline_mask = cv2.morphologyEx(baseline_mask, cv2.MORPH_CLOSE, kernel, iterations=2)
    baseline_mask = cv2.morphologyEx(baseline_mask, cv2.MORPH_OPEN, kernel, iterations=1)

    height, width = baseline_mask.shape
    _ = height
    baseline_y_array = np.full(width, -1, dtype=np.float32)

    for x in range(width):
        y_positions = np.where(baseline_mask[:, x] > 0)[0]
        if len(y_positions) > 0:
            baseline_y_array[x] = float(np.max(y_positions))

    if np.all(baseline_y_array == -1):
        return None, baseline_mask

    valid_indices = np.where(baseline_y_array != -1)[0]
    invalid_indices = np.where(baseline_y_array == -1)[0]
    if len(valid_indices) > 0 and len(invalid_indices) > 0:
        baseline_y_array[invalid_indices] = np.interp(
            invalid_indices,
            valid_indices,
            baseline_y_array[valid_indices],
        )

    return baseline_y_array, baseline_mask


def calculate_interface_distances(
    baseline_y_array: np.ndarray,
    image_height: int,
    image_width: int,
    pixel_size_nm: Optional[float],
    direction: str = "bottom",
) -> List[Dict[str, object]]:
    """Calculate baseline-to-top or baseline-to-bottom distances across x positions."""
    rows: List[Dict[str, object]] = []
    step = max(1, image_width // 100) if image_width > 100 else 1

    for x in range(0, image_width, step):
        if x >= len(baseline_y_array):
            break

        baseline_y = baseline_y_array[x]
        if baseline_y == -1:
            continue

        if direction == "top":
            distance_px = baseline_y
            px_key = "Distance to Top (px)"
            um_key = "Distance to Top (um)"
        else:
            distance_px = image_height - baseline_y
            px_key = "Distance to Bottom (px)"
            um_key = "Distance to Bottom (um)"

        row: Dict[str, object] = {
            "X Position (px)": x,
            "Interface Y (px)": baseline_y,
            px_key: distance_px,
        }

        if pixel_size_nm is not None:
            px_to_um = pixel_size_nm / 1000.0
            row.update(
                {
                    "X Position (um)": x * px_to_um,
                    "Interface Y (um)": baseline_y * px_to_um,
                    um_key: distance_px * px_to_um,
                }
            )

        rows.append(row)

    return rows


def contour_metrics(binary: np.ndarray, contour: np.ndarray, pixel_size_nm: Optional[float], baseline_y_array: Optional[np.ndarray] = None) -> Dict[str, object]:
    """Compute ellipsoidal/contour feature metrics, with optional baseline normalization."""
    moments = cv2.moments(contour)
    if moments["m00"] == 0:
        centroid_x_px = 0.0
        centroid_y_px = 0.0
    else:
        centroid_x_px = moments["m10"] / moments["m00"]
        centroid_y_px = moments["m01"] / moments["m00"]

    area_px2 = cv2.contourArea(contour)
    perimeter_px = cv2.arcLength(contour, True)
    x, y, w, h = cv2.boundingRect(contour)
    aspect_ratio = float(w) / float(h) if h != 0 else 0.0
    extent = float(area_px2) / float(w * h) if (w * h) != 0 else 0.0
    equivalent_diameter_px = sqrt(4 * area_px2 / pi) if area_px2 > 0 else 0.0

    if len(contour) >= 5:
        (_, axes, angle_deg) = cv2.fitEllipse(contour)
        minor_axis_px = min(axes[0], axes[1])
        major_axis_px = max(axes[0], axes[1])
    else:
        major_axis_px = max(binary.shape[0], binary.shape[1])
        minor_axis_px = min(binary.shape[0], binary.shape[1])
        angle_deg = 0.0

    row: Dict[str, object] = {
        "Area (px^2)": area_px2,
        "Perimeter (px)": perimeter_px,
        "Equivalent Diameter (px)": equivalent_diameter_px,
        "width (px)": w,
        "Height (px)": h,
        "Aspect Ratio": aspect_ratio,
        "Extent": extent,
        "Centroid X (px)": centroid_x_px,
        "Centroid Y (px)": centroid_y_px,
        "Major Axis (px)": major_axis_px,
        "Minor Axis (px)": minor_axis_px,
        "Orientation (deg)": angle_deg,
        "X (px)": x,
        "Y (px)": y,
    }

    if baseline_y_array is not None and len(baseline_y_array) > 0:
        centroid_x_int = int(round(centroid_x_px))
        centroid_x_int = max(0, min(centroid_x_int, len(baseline_y_array) - 1))
        baseline_y_at_centroid = baseline_y_array[centroid_x_int]

        x_int = max(0, min(x, len(baseline_y_array) - 1))
        baseline_y_at_x = baseline_y_array[x_int]

        normalized_centroid_y_px = centroid_y_px - baseline_y_at_centroid
        normalized_y_px = y - baseline_y_at_x
        row.update(
            {
                "Baseline Y at Centroid (px)": baseline_y_at_centroid,
                "Normalized Centroid Y (px)": normalized_centroid_y_px,
                "Normalized Y (px)": normalized_y_px,
            }
        )

    if pixel_size_nm is not None:
        px_to_um = pixel_size_nm / 1000.0
        row.update(
            {
                "Area (um^2)": (area_px2 * (px_to_um ** 2)),
                "Perimeter (um)": (perimeter_px * px_to_um),
                "Equivalent Diameter (um)": (equivalent_diameter_px * px_to_um),
                "width (um)": (w * px_to_um),
                "Height (um)": (h * px_to_um),
                "Centroid X (um)": (centroid_x_px * px_to_um),
                "Centroid Y (um)": (centroid_y_px * px_to_um),
                "Major Axis (um)": (major_axis_px * px_to_um),
                "Minor Axis (um)": (minor_axis_px * px_to_um),
            }
        )

        if baseline_y_array is not None and len(baseline_y_array) > 0:
            row.update(
                {
                    "Baseline Y at Centroid (um)": baseline_y_at_centroid * px_to_um,
                    "Normalized Centroid Y (um)": normalized_centroid_y_px * px_to_um,
                    "Normalized Y (um)": normalized_y_px * px_to_um,
                }
            )

    return row


def measure_mask_image(
    mask_path: str,
    pixel_size_nm: Optional[float],
    feature_color: str = "green",
    baseline_color: str = "auto",
    measure_interface_distances: bool = False,
    distance_direction: str = "bottom",
) -> Tuple[List[Dict[str, object]], Optional[np.ndarray], Optional[np.ndarray], Optional[np.ndarray], List[Dict[str, object]]]:
    """Measure carbide features from mask image and optionally interface distance samples."""
    color = cv2.imread(mask_path, cv2.IMREAD_COLOR)
    if color is None:
        return [], None, None, None, []

    baseline_y_array, baseline_mask = detect_baseline_region(color, baseline_color)

    baseline_viz = color.copy()
    if baseline_y_array is not None:
        points: List[Tuple[int, int]] = []
        for x in range(len(baseline_y_array)):
            points.append((x, int(round(baseline_y_array[x]))))

        if len(points) > 1:
            points_array = np.array(points, dtype=np.int32)
            cv2.polylines(baseline_viz, [points_array], False, (255, 255, 0), 2)

        if baseline_mask is not None:
            red_overlay = np.zeros_like(baseline_viz)
            red_overlay[:] = (0, 0, 255)
            baseline_viz = np.where(
                baseline_mask[:, :, np.newaxis] > 0,
                cv2.addWeighted(baseline_viz, 0.7, red_overlay, 0.3, 0),
                baseline_viz,
            )

    bw_features = color_name_to_bw_mask(color, feature_color)
    if int(bw_features.sum()) == 0:
        return [], color, baseline_viz, color.copy(), []

    features_viz = color.copy()
    contours, _ = cv2.findContours(bw_features, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cv2.drawContours(features_viz, contours, -1, (0, 255, 0), 2)

    rows: List[Dict[str, object]] = []
    for contour in contours:
        rows.append(contour_metrics(bw_features, contour, pixel_size_nm, baseline_y_array))

    interface_rows: List[Dict[str, object]] = []
    if measure_interface_distances and baseline_y_array is not None:
        interface_rows = calculate_interface_distances(
            baseline_y_array,
            color.shape[0],
            color.shape[1],
            pixel_size_nm,
            distance_direction,
        )

    return rows, color, baseline_viz, features_viz, interface_rows
