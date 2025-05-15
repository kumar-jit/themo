import sys
import os
import ctypes
import winreg
import requests
import tempfile
import urllib.parse  # For URL encoding search queries

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QComboBox,
    QSpinBox, QMessageBox, QFrame, QTabWidget,
    QListWidget, QListWidgetItem, QLineEdit,  # Added QLineEdit for search
    QSizePolicy
)
from PyQt6.QtGui import QPixmap, QFont, QImageReader, QPainter, QColor
from PyQt6.QtCore import Qt, QTimer, QSize, QByteArray, QBuffer, QIODevice

# Windows API constants for setting wallpaper
SPI_SETDESKWALLPAPER = 0x0014
SPIF_UPDATEINIFILE = 0x01
SPIF_SENDWININICHANGE = 0x02

# Supported image extensions for folder scanning (local)
IMAGE_EXTENSIONS = ('.png', '.jpg', '.jpeg', '.bmp', '.gif')

# --- Pixabay API Configuration ---
PIXABAY_API_KEY = "50316370-ada398d289a6bffb2e36f55aa"
PIXABAY_API_URL = "https://pixabay.com/api/"


def set_wallpaper_windows(image_path, style_name="Stretch"):
    if not image_path or not os.path.exists(image_path):
        QMessageBox.critical(None, "Error", f"Image path is invalid or does not exist: {image_path}")
        return False

    abs_image_path = os.path.abspath(image_path)
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Control Panel\Desktop", 0, winreg.KEY_SET_VALUE)
        style_value_reg, tile_value_reg = "2", "0"  # Default Stretch
        if style_name == "Fill":
            style_value_reg = "10"
        elif style_name == "Fit":
            style_value_reg = "6"
        elif style_name == "Stretch":
            style_value_reg = "2"
        elif style_name == "Tile":
            style_value_reg, tile_value_reg = "0", "1"
        elif style_name == "Center":
            style_value_reg, tile_value_reg = "0", "0"

        winreg.SetValueEx(key, "WallpaperStyle", 0, winreg.REG_SZ, style_value_reg)
        winreg.SetValueEx(key, "TileWallpaper", 0, winreg.REG_SZ, tile_value_reg)
        winreg.CloseKey(key)
    except Exception as e:
        QMessageBox.critical(None, "Registry Error", f"Error setting wallpaper style: {e}")
        return False

    try:
        if not ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, abs_image_path,
                                                          SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE):
            error_code = ctypes.get_last_error()
            error_message = ctypes.WinError(error_code).strerror
            QMessageBox.critical(None, "API Error",
                                 f"SystemParametersInfoW failed: {error_message} (Code: {error_code})")
            return False
        return True
    except Exception as e:
        QMessageBox.critical(None, "API Call Error", f"Error calling SystemParametersInfoW: {e}")
        return False


class ImageCardWidget(QWidget):
    """Custom widget to display an image thumbnail and name in a card format."""

    def __init__(self, pixabay_hit_data, parent=None):  # Takes data from a Pixabay API hit
        super().__init__(parent)
        self.pixabay_hit_data = pixabay_hit_data
        self.thumbnail_pixmap = None
        # Use previewURL for thumbnail, tags for name
        self.thumbnail_url = pixabay_hit_data.get('previewURL')
        image_name = pixabay_hit_data.get('tags', 'Pixabay Image')

        self.setMinimumSize(150, 180)
        self.setMaximumSize(180, 200)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(5)

        self.thumbnail_label = QLabel()
        self.thumbnail_label.setFixedSize(140, 100)
        self.thumbnail_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.thumbnail_label.setStyleSheet("border: 1px solid #ccc; background-color: #f0f0f0;")
        self.thumbnail_label.setText("Loading...")
        layout.addWidget(self.thumbnail_label)

        self.name_label = QLabel(image_name)
        self.name_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.name_label.setWordWrap(True)
        font = self.name_label.font()
        font.setPointSize(8)
        self.name_label.setFont(font)
        layout.addWidget(self.name_label)

        self.setLayout(layout)

    def load_thumbnail(self):
        if not self.thumbnail_url:
            self.thumbnail_label.setText("No Thumb URL")
            return

        try:
            response = requests.get(self.thumbnail_url, stream=True, timeout=10)
            response.raise_for_status()

            image_data_bytes = response.content

            pixmap = QPixmap()
            if pixmap.loadFromData(image_data_bytes):
                self.thumbnail_pixmap = pixmap.scaled(
                    self.thumbnail_label.size(),
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation
                )
                self.thumbnail_label.setPixmap(self.thumbnail_pixmap)
            else:
                self.thumbnail_label.setText("Load Error")
        except requests.exceptions.RequestException as e:
            self.thumbnail_label.setText("DL Error")
            print(f"Error downloading thumbnail {self.pixabay_hit_data.get('tags', 'Unknown')}: {e}")
        except Exception as e:
            self.thumbnail_label.setText("Proc. Error")
            print(f"Error processing thumbnail {self.pixabay_hit_data.get('tags', 'Unknown')}: {e}")


class WallpaperApp(QWidget):
    def __init__(self):
        super().__init__()
        self.current_image_path = None
        self.local_image_folder_paths = []
        self.current_folder_index = 0
        self.slideshow_timer = QTimer(self)
        self.slideshow_timer.timeout.connect(self.next_slide)

        self.active_temp_online_file = None
        self.pixabay_search_results = []  # Stores hits from Pixabay API

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("TheMo Desktop Wallpaper Changer (Windows)")
        self.setMinimumSize(600, 700)

        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(10, 10, 10, 10)

        self.preview_display_label = QLabel("Select an image to preview")
        self.preview_display_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_display_label.setFrameShape(QFrame.Shape.StyledPanel)
        self.preview_display_label.setMinimumSize(520, 300)
        self.preview_display_label.setMaximumSize(580, 350)
        font = self.preview_display_label.font()
        font.setPointSize(12)
        self.preview_display_label.setFont(font)
        main_layout.addWidget(self.preview_display_label, alignment=Qt.AlignmentFlag.AlignCenter)

        controls_layout = QHBoxLayout()
        style_label = QLabel("Wallpaper Style:")
        self.style_combo = QComboBox()
        self.style_combo.addItems(["Fill", "Fit", "Stretch", "Tile", "Center"])
        self.style_combo.setCurrentText("Fill")
        controls_layout.addWidget(style_label)
        controls_layout.addWidget(self.style_combo)

        self.apply_wallpaper_button = QPushButton("Set as Wallpaper")
        self.apply_wallpaper_button.clicked.connect(self.apply_current_wallpaper)
        self.apply_wallpaper_button.setEnabled(False)
        self.apply_wallpaper_button.setFixedHeight(35)
        font = self.apply_wallpaper_button.font()
        font.setPointSize(10)
        self.apply_wallpaper_button.setFont(font)
        controls_layout.addWidget(self.apply_wallpaper_button)
        main_layout.addLayout(controls_layout)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._create_local_tab(), "Local Wallpaper")
        self.tabs.addTab(self._create_online_tab(), "Online Wallpaper (Pixabay)")  # Updated Tab Title
        main_layout.addWidget(self.tabs)

        self.setLayout(main_layout)
        self.show()

    def _create_local_tab(self):
        local_tab_widget = QWidget()
        layout = QVBoxLayout(local_tab_widget)
        layout.setSpacing(10)

        selection_layout = QHBoxLayout()
        self.browse_file_button = QPushButton("Browse Image File")
        self.browse_file_button.clicked.connect(self.browse_local_file)
        selection_layout.addWidget(self.browse_file_button)

        self.browse_folder_button = QPushButton("Browse Image Folder")
        self.browse_folder_button.clicked.connect(self.browse_local_folder)
        selection_layout.addWidget(self.browse_folder_button)
        layout.addLayout(selection_layout)

        self.local_selected_path_label = QLabel("No local file or folder selected.")
        self.local_selected_path_label.setWordWrap(True)
        self.local_selected_path_label.setStyleSheet("color: gray;")
        layout.addWidget(self.local_selected_path_label)

        slideshow_group_layout = QVBoxLayout()
        timer_layout = QHBoxLayout()
        timer_label = QLabel("Slideshow Interval (minutes):")
        self.timer_spinbox = QSpinBox()
        self.timer_spinbox.setRange(1, 1440)
        self.timer_spinbox.setValue(5)
        timer_layout.addWidget(timer_label)
        timer_layout.addWidget(self.timer_spinbox)
        slideshow_group_layout.addLayout(timer_layout)

        slideshow_buttons_layout = QHBoxLayout()
        self.start_slideshow_button = QPushButton("Start Slideshow")
        self.start_slideshow_button.clicked.connect(self.start_slideshow)
        self.start_slideshow_button.setEnabled(False)
        slideshow_buttons_layout.addWidget(self.start_slideshow_button)

        self.stop_slideshow_button = QPushButton("Stop Slideshow")
        self.stop_slideshow_button.clicked.connect(self.stop_slideshow)
        self.stop_slideshow_button.setEnabled(False)
        slideshow_buttons_layout.addWidget(self.stop_slideshow_button)
        slideshow_group_layout.addLayout(slideshow_buttons_layout)
        layout.addLayout(slideshow_group_layout)

        layout.addStretch()
        return local_tab_widget

    def _create_online_tab(self):
        online_tab_widget = QWidget()
        layout = QVBoxLayout(online_tab_widget)
        layout.setSpacing(10)

        # Search bar and button
        search_layout = QHBoxLayout()
        self.pixabay_search_input = QLineEdit()
        self.pixabay_search_input.setPlaceholderText("Enter search term (e.g., nature, cats)")
        search_layout.addWidget(self.pixabay_search_input)

        self.search_pixabay_button = QPushButton("Search Images")
        self.search_pixabay_button.clicked.connect(self.search_pixabay_images)
        search_layout.addWidget(self.search_pixabay_button)
        layout.addLayout(search_layout)

        self.online_images_listwidget = QListWidget()
        self.online_images_listwidget.setViewMode(QListWidget.ViewMode.IconMode)
        self.online_images_listwidget.setResizeMode(QListWidget.ResizeMode.Adjust)
        self.online_images_listwidget.setMovement(QListWidget.Movement.Static)
        self.online_images_listwidget.setSpacing(10)
        self.online_images_listwidget.itemClicked.connect(self.handle_online_image_card_selection)
        layout.addWidget(self.online_images_listwidget)

        self.online_status_label = QLabel("Status: Enter a search term and click 'Search Images'.")
        self.online_status_label.setWordWrap(True)
        layout.addWidget(self.online_status_label)

        return online_tab_widget

    def browse_local_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Image File", "",
                                                   f"Image Files ({' '.join(['*' + ext for ext in IMAGE_EXTENSIONS])})")
        if file_path:
            self.stop_slideshow_if_active()
            self.cleanup_temp_online_file()
            self.current_image_path = file_path
            self.local_image_folder_paths = []
            self.update_shared_preview(file_path)
            self.local_selected_path_label.setText(f"Selected file: {file_path}")
            self.local_selected_path_label.setStyleSheet("color: black;")
            self.apply_wallpaper_button.setEnabled(True)
            self.start_slideshow_button.setEnabled(False)

    def browse_local_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Image Folder")
        if folder_path:
            self.stop_slideshow_if_active()
            self.cleanup_temp_online_file()
            self.local_image_folder_paths = self.scan_folder(folder_path)
            if self.local_image_folder_paths:
                self.current_folder_index = 0
                self.current_image_path = self.local_image_folder_paths[0]
                self.update_shared_preview(self.current_image_path)
                self.local_selected_path_label.setText(
                    f"Selected folder: {folder_path} ({len(self.local_image_folder_paths)} images)")
                self.local_selected_path_label.setStyleSheet("color: black;")
                self.apply_wallpaper_button.setEnabled(True)
                self.start_slideshow_button.setEnabled(True)
            else:
                self.current_image_path = None
                self.update_shared_preview(None)
                self.local_selected_path_label.setText(f"No supported images found in: {folder_path}")
                self.local_selected_path_label.setStyleSheet("color: red;")
                self.apply_wallpaper_button.setEnabled(False)
                self.start_slideshow_button.setEnabled(False)
                QMessageBox.information(self, "No Images", "No supported image files found in the selected folder.")

    def scan_folder(self, folder_path):
        images = []
        for item in os.listdir(folder_path):
            item_path = os.path.join(folder_path, item)
            if os.path.isfile(item_path) and item.lower().endswith(IMAGE_EXTENSIONS):
                images.append(item_path)
        return sorted(images)

    def update_shared_preview(self, image_path=None):
        if image_path and os.path.exists(image_path):
            pixmap = QPixmap(image_path)
            if pixmap.isNull():
                self.preview_display_label.setText("Cannot load preview")
                self.apply_wallpaper_button.setEnabled(False)
            else:
                scaled_pixmap = pixmap.scaled(self.preview_display_label.size(),
                                              Qt.AspectRatioMode.KeepAspectRatio,
                                              Qt.TransformationMode.SmoothTransformation)
                self.preview_display_label.setPixmap(scaled_pixmap)
                self.apply_wallpaper_button.setEnabled(True if self.current_image_path else False)
        else:
            self.preview_display_label.setText("Select an image to preview")
            font = self.preview_display_label.font()
            font.setPointSize(12)
            self.preview_display_label.setFont(font)
            self.preview_display_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.apply_wallpaper_button.setEnabled(False)

    def apply_current_wallpaper(self):
        if not self.current_image_path:
            QMessageBox.warning(self, "No Image", "Please select an image first (either local or online).")
            return

        style = self.style_combo.currentText()
        if set_wallpaper_windows(self.current_image_path, style):
            QMessageBox.information(self, "Success",
                                    f"Wallpaper set to:\n{os.path.basename(self.current_image_path)}\nStyle: {style}")

    def start_slideshow(self):
        if not self.local_image_folder_paths:
            QMessageBox.warning(self, "No Folder", "Please select a local folder with images for the slideshow.")
            return

        interval_minutes = self.timer_spinbox.value()
        if interval_minutes <= 0:
            QMessageBox.warning(self, "Invalid Interval", "Slideshow interval must be greater than 0 minutes.")
            return

        self.cleanup_temp_online_file()

        self.current_image_path = self.local_image_folder_paths[self.current_folder_index]
        self.update_shared_preview(self.current_image_path)
        self.apply_current_wallpaper()

        self.slideshow_timer.start(interval_minutes * 60 * 1000)
        self.start_slideshow_button.setEnabled(False)
        self.stop_slideshow_button.setEnabled(True)
        self.browse_file_button.setEnabled(False)
        self.browse_folder_button.setEnabled(False)
        self.tabs.setTabEnabled(1, False)
        QMessageBox.information(self, "Slideshow Started",
                                f"Slideshow started with an interval of {interval_minutes} minutes.")

    def stop_slideshow_if_active(self):
        if self.slideshow_timer.isActive():
            self.slideshow_timer.stop()
            self.start_slideshow_button.setEnabled(True if self.local_image_folder_paths else False)
            self.stop_slideshow_button.setEnabled(False)
            self.browse_file_button.setEnabled(True)
            self.browse_folder_button.setEnabled(True)
            self.tabs.setTabEnabled(1, True)
            return True
        return False

    def stop_slideshow(self):
        if self.stop_slideshow_if_active():
            QMessageBox.information(self, "Slideshow Stopped", "Slideshow has been stopped.")
        else:
            self.stop_slideshow_button.setEnabled(False)

    def next_slide(self):
        if not self.local_image_folder_paths or not self.slideshow_timer.isActive():
            return

        self.current_folder_index = (self.current_folder_index + 1) % len(self.local_image_folder_paths)
        self.current_image_path = self.local_image_folder_paths[self.current_folder_index]
        self.update_shared_preview(self.current_image_path)
        self.local_selected_path_label.setText(f"Slideshow: {self.current_image_path}")
        self.apply_current_wallpaper()

    # --- Online Wallpaper Methods (Pixabay) ---
    def search_pixabay_images(self):
        search_term = self.pixabay_search_input.text().strip()
        if not search_term:
            QMessageBox.information(self, "Search Term Required", "Please enter a search term for Pixabay.")
            self.online_status_label.setText("Status: Please enter a search term.")
            return

        self.online_status_label.setText(f"Status: Searching Pixabay for '{search_term}'...")
        QApplication.processEvents()

        encoded_search_term = urllib.parse.quote_plus(search_term)
        params = {
            'key': PIXABAY_API_KEY,
            'q': encoded_search_term,
            'image_type': 'photo',
            'safesearch': 'true',
            'per_page': 50  # Fetch a decent number of results
        }

        try:
            response = requests.get(PIXABAY_API_URL, params=params, timeout=15)
            response.raise_for_status()  # Raise an exception for HTTP errors (4xx or 5xx)

            data = response.json()
            self.pixabay_search_results = data.get('hits', [])

            self.online_images_listwidget.clear()
            if not self.pixabay_search_results:
                self.online_status_label.setText(f"Status: No images found on Pixabay for '{search_term}'.")
                QMessageBox.information(self, "No Results", f"No images found on Pixabay for '{search_term}'.")
                return

            for hit_data in self.pixabay_search_results:
                card_widget = ImageCardWidget(hit_data)  # Pass the whole hit data

                list_item = QListWidgetItem(self.online_images_listwidget)
                list_item.setSizeHint(card_widget.sizeHint())
                # Store the full hit data with the QListWidgetItem for later retrieval
                list_item.setData(Qt.ItemDataRole.UserRole, hit_data)

                self.online_images_listwidget.addItem(list_item)
                self.online_images_listwidget.setItemWidget(list_item, card_widget)

                card_widget.load_thumbnail()
                QApplication.processEvents()

            self.online_status_label.setText(
                f"Status: Found {len(self.pixabay_search_results)} images for '{search_term}'. Select one.")

        except requests.exceptions.RequestException as e:
            self.online_status_label.setText(f"Status: Error searching Pixabay: {e}")
            QMessageBox.critical(self, "Pixabay API Error", f"Failed to search images on Pixabay: {e}")
        except Exception as e:
            self.online_status_label.setText(f"Status: An unexpected error occurred: {e}")
            QMessageBox.critical(self, "Error", f"An unexpected error occurred: {e}")

    def handle_online_image_card_selection(self, list_item):
        self.stop_slideshow_if_active()

        # Retrieve the full hit data stored in the item
        pixabay_hit_data = list_item.data(Qt.ItemDataRole.UserRole)

        # Prefer webformatURL, fallback to largeImageURL, then previewURL if others are missing
        image_url = pixabay_hit_data.get('webformatURL') or \
                    pixabay_hit_data.get('largeImageURL') or \
                    pixabay_hit_data.get('previewURL')

        image_name = pixabay_hit_data.get('tags', 'Pixabay Image')

        if not image_url:
            QMessageBox.warning(self, "No Image URL", "Selected Pixabay image has no usable URL.")
            self.online_status_label.setText("Status: Selected image has no valid URL.")
            return

        self.online_status_label.setText(f"Status: Downloading '{image_name}' from Pixabay...")
        self.preview_display_label.setText(f"Downloading '{image_name}'...")
        QApplication.processEvents()

        try:
            response = requests.get(image_url, stream=True, timeout=20)  # Increased timeout
            response.raise_for_status()

            self.cleanup_temp_online_file()

            # Determine file suffix from URL or default
            _, suffix = os.path.splitext(image_url.split('?')[0])  # Split '?' to handle query params in URL
            if not suffix or len(suffix) > 5 or len(suffix) < 2:
                # Try to get type from Pixabay data if available
                img_type = pixabay_hit_data.get('type', 'photo')  # 'type' in hit is 'photo', 'illustration' etc.
                if img_type == 'jpeg' or img_type == 'jpg':
                    suffix = ".jpg"
                elif img_type == 'png':
                    suffix = ".png"
                else:
                    suffix = ".img"  # Default if type is not helpful for extension

            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=suffix, prefix="themo_pixabay_")
            self.active_temp_online_file = temp_file.name

            with temp_file as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)

            self.current_image_path = self.active_temp_online_file
            self.update_shared_preview(self.current_image_path)
            self.apply_wallpaper_button.setEnabled(True)
            self.online_status_label.setText(f"Status: Previewing '{image_name}'. Ready to set.")
            self.local_selected_path_label.setText("Online image (Pixabay) selected.")
            self.local_selected_path_label.setStyleSheet("color: gray;")

        except requests.exceptions.RequestException as e:
            self.online_status_label.setText(f"Status: Failed to download '{image_name}': {e}")
            self.preview_display_label.setText("Download failed")
            QMessageBox.critical(self, "Download Error", f"Failed to download image from Pixabay: {e}")
            self.current_image_path = None
            self.apply_wallpaper_button.setEnabled(False)
            self.cleanup_temp_online_file(complain=False)

    def cleanup_temp_online_file(self, complain=True):
        if self.active_temp_online_file:
            try:
                if os.path.exists(self.active_temp_online_file):
                    os.remove(self.active_temp_online_file)
            except OSError as e:
                if complain: print(f"Warning: Could not delete temp file {self.active_temp_online_file}: {e}")
            finally:
                self.active_temp_online_file = None

    def closeEvent(self, event):
        self.stop_slideshow_if_active()
        self.cleanup_temp_online_file(complain=False)
        event.accept()


if __name__ == '__main__':
    if os.name != 'nt':
        app_dummy = QApplication(sys.argv)
        QMessageBox.warning(None, "OS Compatibility",
                            "This application uses Windows-specific features to change the wallpaper. "
                            "Wallpaper setting may not function on other operating systems.")

    app = QApplication(sys.argv)
    ex = WallpaperApp()
    sys.exit(app.exec())
