# Montana Dot Mapper

A Python application for creating dot maps of Montana specimen locations with topographic background.

## Features

- **Dot Map Visualization**: Shows individual specimen locations as red dots on a map
- **Topographic Background**: Includes terrain features like mountains, rivers, and flat areas using OpenTopoMap
- **Species Filtering**: Filter data by Family, Genus, and Species
- **High-Resolution Export**: Save maps as high-quality TIFF files
- **User-Friendly Interface**: Simple GUI with dropdown selections and file browsing

## Requirements

- Python 3.7+
- pandas >= 1.3.0
- geopandas >= 0.10.0
- matplotlib >= 3.4.0
- shapely >= 1.7.0
- openpyxl >= 3.0.0
- numpy >= 1.21.0
- contextily >= 1.2.0

## Installation

1. Clone or download this repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```bash
   python montana_dot_mapper.py
   ```

2. Load an Excel file containing specimen data with the following columns:
   - `lat`: Latitude coordinates
   - `lat_dir`: Latitude direction (N/S)
   - `long`: Longitude coordinates
   - `long_dir`: Longitude direction (E/W)
   - `family`: Taxonomic family
   - `genus`: Taxonomic genus
   - `species`: Taxonomic species
   - `year`: Collection year

3. Select Family, Genus, and Species from the dropdown menus

4. Click "Generate Dot Map" to create a map showing specimen locations

5. Use "Download Dot Map" to save the map as a high-resolution TIFF file

## Map Features

- **Red Dots**: Each dot represents a specimen found at that location
- **County Boundaries**: Montana county borders are shown in black
- **Topographic Background**: Terrain features including mountains, rivers, and elevation
- **Legend**: Shows the total number of specimens found
- **Title**: Displays the selected taxonomic hierarchy and specimen count

## File Structure

```
MontanaDotMapper/
├── montana_dot_mapper.py      # Main application
├── requirements.txt           # Python dependencies
├── README.md                  # This file
├── app_icon.ico              # Application icon
├── shapefiles/               # Montana county shapefiles
│   ├── cb_2021_us_county_5m.shp
│   ├── cb_2021_us_county_5m.dbf
│   └── ...
└── venv/                     # Virtual environment (if used)
```

## Notes

- The application requires an internet connection to load the topographic background maps
- If the background map fails to load, the application will fall back to a simple county boundary display
- All coordinates are automatically converted to the appropriate coordinate system for display
- The application filters out coordinates that are outside Montana's boundaries

## Troubleshooting

- **Background map not loading**: Check your internet connection
- **No dots appearing**: Verify that your coordinates are within Montana's boundaries
- **Import errors**: Make sure all dependencies are installed correctly
- **File loading errors**: Ensure your Excel file has all required columns

## License

This application is provided as-is for educational and research purposes. 