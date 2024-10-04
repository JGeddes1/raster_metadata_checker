import os
import sys
import openpyxl
import folium
from geopy.geocoders import Nominatim
from pyproj import Transformer
from PyQt5.QtWidgets import (QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, 
                             QHBoxLayout, QFileDialog, QListWidget, QMessageBox, QGridLayout)
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import QUrl
import geopandas as gpd
from shapely.geometry import Point


# Initialize geocoder
geolocator = Nominatim(user_agent="geoapiExercises")
# Transformer for OSGB to WGS84 (Lat/Lon)
transformer = Transformer.from_crs("epsg:27700", "epsg:4326", always_xy=True)
# Transformer for OSGB (EPSG:27700) to WGS84 (EPSG:4326)
wgs84_to_osgb = Transformer.from_crs("epsg:4326", "epsg:27700", always_xy=True)  # Lon/Lat to British National Grid

class FileChecker(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('File Checker')
        self.setGeometry(100, 100, 800, 600)

        # Directory selection
        self.directory_label = QLabel('Select Directory:', self)
        self.directory_path = QLineEdit(self)
        self.browse_directory_button = QPushButton('Browse', self)
        self.browse_directory_button.clicked.connect(self.browse_directory)

        # Excel file selection
        self.excel_label = QLabel('Select Excel File:', self)
        self.excel_file_path = QLineEdit(self)
        self.browse_excel_button = QPushButton('Browse', self)
        self.browse_excel_button.clicked.connect(self.browse_excel_file)

        # Second Excel file selection
        self.excel_label2 = QLabel('Select Second Excel File:', self)
        self.excel_file_path2 = QLineEdit(self)
        self.browse_excel_button2 = QPushButton('Browse', self)
        self.browse_excel_button2.clicked.connect(self.browse_excel_file2)

        # Check files button
        self.check_button = QPushButton('Check Files', self)
        self.check_button.clicked.connect(self.check_files)

        # Listboxes for missing files and keywords
        self.missing_directory_label = QLabel('Missing from directory:', self)
        self.missing_directory_listbox = QListWidget(self)
        self.missing_directory_listbox.setFixedHeight(300)  # Set fixed height

        self.missing_metadata_label = QLabel('Not listed in metadata:', self)
        self.missing_metadata_listbox = QListWidget(self)
        self.missing_metadata_listbox.setFixedHeight(300)  # Set fixed height

        self.keyword_label = QLabel('Subject Keywords:', self)
        self.keyword_listbox = QListWidget(self)
        self.keyword_listbox.setFixedHeight(300)  # Set fixed height

        # Map display using QWebEngineView
        self.map_view = QWebEngineView()
        self.map_view.setMinimumSize(500, 300)  # Set minimum size for the map

        # Status label
        self.status_label = QLabel('', self)
        self.status_label_Lat_Long = QLabel('', self)

        # Layouts
        main_layout = QVBoxLayout()

        directory_layout = QHBoxLayout()
        directory_layout.addWidget(self.directory_label)
        directory_layout.addWidget(self.directory_path)
        directory_layout.addWidget(self.browse_directory_button)
        main_layout.addLayout(directory_layout)

        excel_layout = QHBoxLayout()
        excel_layout.addWidget(self.excel_label)
        excel_layout.addWidget(self.excel_file_path)
        excel_layout.addWidget(self.browse_excel_button)
        main_layout.addLayout(excel_layout)

        excel_layout2 = QHBoxLayout()
        excel_layout2.addWidget(self.excel_label2)
        excel_layout2.addWidget(self.excel_file_path2)
        excel_layout2.addWidget(self.browse_excel_button2)
        main_layout.addLayout(excel_layout2)

        main_layout.addWidget(self.check_button)

        listbox_layout = QGridLayout()
        listbox_layout.addWidget(self.missing_directory_label, 0, 0)
        listbox_layout.addWidget(self.missing_directory_listbox, 1, 0)
        listbox_layout.addWidget(self.missing_metadata_label, 0, 1)
        listbox_layout.addWidget(self.missing_metadata_listbox, 1, 1)
        listbox_layout.addWidget(self.keyword_label, 0, 2)
        listbox_layout.addWidget(self.keyword_listbox, 1, 2)
        main_layout.addLayout(listbox_layout)

        # Add map view below the list boxes with some padding
        main_layout.addWidget(self.map_view)
        main_layout.addWidget(self.status_label)
        main_layout.addWidget(self.status_label_Lat_Long)

        main_layout.setSpacing(10)  # Set spacing between widgets
        self.setLayout(main_layout)




    def browse_directory(self):
        directory = QFileDialog.getExistingDirectory(self, 'Select Directory')
        if directory:
            self.directory_path.setText(directory)

    def browse_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, 'Select Excel File', '', 'Excel files (*.xlsx *.xls)')
        if file_path:
            self.excel_file_path.setText(file_path)

    def browse_excel_file2(self):
        file_path, _ = QFileDialog.getOpenFileName(self, 'Select Second Excel File', '', 'Excel files (*.xlsx *.xls)')
        if file_path:
            self.excel_file_path2.setText(file_path)

    def read_excel_file(self, file_path):
        filenames = []
        coordinates = []  # Store coordinates as well
        global subjectkeywords
        subjectkeywords = set()  # Initialize here

        workbook = openpyxl.load_workbook(file_path)
        worksheet = workbook.active
        for row in worksheet.iter_rows(values_only=True):
            if row[0] and not str(row[0]).lower() in ["filename", "filename"]:
                filenames.append(row[0].strip().lower())
            
            # Check for keywords and add them
            for index in range(2, 5):  # Assuming keywords are in columns 3, 4, and 5 (index 2, 3, 4)
                if row[index] and isinstance(row[index], str):  # Check if it's a string
                    subjectkeywords.add(row[index].strip().lower())

            try:
                # Check if lat/lon (rows 17/18) are provided, use them if available
                if row[17] is not None and row[18] is not None:  # Adjust index for lat/lon (17/18 correspond to indices 16/17)
                    latitude = float(row[18])
                    longitude = float(row[17])
                    print(f"Lat/Lon for {row[0]}: Latitude={latitude}, Longitude={longitude}")
                    coordinates.append((latitude, longitude))
                elif row[19] is not None and row[20] is not None:  # Fall back to OSGB if lat/lon are missing
                    osgb_easting = float(row[19])  # Adjust index for easting
                    osgb_northing = float(row[20])  # Adjust index for northing
                    print(f"OSGB Coordinates for {row[0]}: Easting={osgb_easting}, Northing={osgb_northing}")
                    
                    # Convert to Latitude/Longitude
                    longitude, latitude = transformer.transform(osgb_easting, osgb_northing)
                    print(f"Converted Coordinates for {row[0]}: Latitude={latitude}, Longitude={longitude}")
                    coordinates.append((latitude, longitude))  # Append in (latitude, longitude) order
                else:
                    coordinates.append((None, None))  # Append None for missing coordinates
            except (ValueError, IndexError) as e:
                print(f"Invalid coordinates for {row[0]}: {e}")
                coordinates.append((None, None))  # Append None for invalid data
        return filenames, coordinates


       

    def list_files(self, directory):
        image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif', '.dng']
        return [filename.lower() for filename in os.listdir(directory) if os.path.isfile(os.path.join(directory, filename)) and os.path.splitext(filename)[1].lower() in image_extensions]

    def find_missing_files(self, directory_files, excel_filenames):
        missing_in_directory = [filename for filename in excel_filenames if filename not in directory_files]
        missing_in_metadata = [filename for filename in directory_files if filename not in excel_filenames]
        return missing_in_directory, missing_in_metadata

    def check_files(self):
            directory = self.directory_path.text()
            excel_file1 = self.excel_file_path.text()
            excel_file2 = self.excel_file_path2.text()

            if not directory or not excel_file1:
                QMessageBox.warning(self, 'Error', 'Please select a directory and at least one Excel file.')
                return

            # Define keyword_remove_list here
            keyword_remove_list = ['subject keyword 3', 'subject keyword 2', 'subject keyword 1']  # Example keywords to remove

            # Read Excel files and get filenames and coordinates
            excel_filenames1, coordinates1 = self.read_excel_file(excel_file1)
            excel_filenames2, coordinates2 = self.read_excel_file(excel_file2) if excel_file2 else ([], [])
            excel_filenames = list(set(excel_filenames1 + excel_filenames2))

            # List files in directory
            directory_files = self.list_files(directory)

            # Find missing files
            missing_in_directory, missing_in_metadata = self.find_missing_files(directory_files, excel_filenames)

            # Clear listboxes
            self.missing_directory_listbox.clear()
            self.missing_metadata_listbox.clear()
            self.keyword_listbox.clear()

            # Populate listboxes
            for file in missing_in_directory:
                self.missing_directory_listbox.addItem(file)
            for file in missing_in_metadata:
                self.missing_metadata_listbox.addItem(file)
            
            # Add keywords to listbox, excluding the ones in keyword_remove_list
            for keyword in subjectkeywords:
                if keyword not in keyword_remove_list:
                    self.keyword_listbox.addItem(keyword)

            # Set status
            if not missing_in_directory and not missing_in_metadata:
                self.status_label.setText("No missing files found! All good!")
                self.status_label.setStyleSheet("color: green;")
            else:
                self.status_label.setText("Missing files detected. Check the lists above.")
                self.status_label.setStyleSheet("color: red;")

            # Process unique coordinates
            self.process_unique_coordinates(excel_filenames, coordinates1 + coordinates2)


    def process_unique_coordinates(self, filenames, coordinates):
        # Track unique coordinates
        processed_coordinates = set()

        unique_filenames = []
        unique_coordinates = []

        for filename, (lat, lon) in zip(filenames, coordinates):
            if lat is not None and lon is not None and (lat, lon) not in processed_coordinates:
                processed_coordinates.add((lat, lon))
                unique_filenames.append(filename)
                unique_coordinates.append((lat, lon))

        # Create the map only with unique coordinates
        self.create_map(unique_filenames, unique_coordinates)

        # Display unique lat/lon in label
        if unique_coordinates:
            coordinates_text = "\n".join([f"Lat: {lat}, Lon: {lon}" for lat, lon in unique_coordinates])
            self.status_label_Lat_Long.setText(coordinates_text)
        else:
            self.status_label_Lat_Long.setText("No valid coordinates found.")



    def find_region_with_shapefile(self, lat, lon):
        # Load UK region shapefile (adjust file path)
        shapefile_path = 'C:/Users/jgg513/Downloads/NUTS1/NUTS1_Jan_2018_SGCB_in_the_UK.shp'
        regions = gpd.read_file(shapefile_path)
        
        # Convert Lat/Lon to British National Grid (EPSG:27700)
        osgb_x, osgb_y = wgs84_to_osgb.transform(lon, lat)
        
        # Create a point object in the shapefile's CRS
        point = Point(osgb_x, osgb_y)
        
        # Find the region that contains the point
        region = regions[regions.contains(point)]
        
        if not region.empty:
            return region.iloc[0]['nuts118nm']  # Adjust based on the shapefile column name
        else:
            return "Unknown Region"

    def create_map(self, filenames, coordinates):
            valid_coordinates = [(lat, lon) for lat, lon in coordinates if lat is not None and lon is not None]
            
            if valid_coordinates:
                initial_lat, initial_lon = valid_coordinates[0]
            else:
                initial_lat, initial_lon = 0, 0

            folium_map = folium.Map(location=[initial_lat, initial_lon], zoom_start=13)

            for filename, (lat, lon) in zip(filenames, coordinates):
                if lat is not None and lon is not None:
                    region = self.find_region_with_shapefile(lat, lon)  # Use self to call the method
                    popup_text = f"{filename}<br>Region: {region}"
                    folium.Marker(location=[lat, lon], popup=popup_text).add_to(folium_map)

            map_path = os.path.abspath('raster_metadata_map.html')
            folium_map.save(map_path)

            if os.path.exists(map_path):
                print(f"Map saved at {map_path}")
            else:
                print("Error: Map not saved!")

            self.map_view.setUrl(QUrl.fromLocalFile(map_path))
            print(f"Map URL set to {map_path}")




if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = FileChecker()
    ex.show()
    sys.exit(app.exec_())
