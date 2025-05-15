# TheMo Desktop Wallpaper Changer (Windows)

TheMo is a Python-based desktop application for Windows that allows users to change their desktop wallpaper. Users can select images from their local machine (single file or a folder for a slideshow) or search for and apply high-quality images from Pixabay.

## Features

* **Local Wallpaper Management:**
    * Select a single image file to set as wallpaper.
    * Select a folder of images for an automatic slideshow.
    * Configurable slideshow interval (in minutes).
* **Online Wallpaper Discovery (via Pixabay API):**
    * Search for images on Pixabay using keywords.
    * View search results in an intuitive card format with thumbnails.
    * Download and preview selected online images.
* **Wallpaper Display Styles:**
    * Choose how the wallpaper fits the screen: Fill, Fit, Stretch, Tile, or Center.
* **User Interface:**
    * Tabbed interface for easy navigation between local and online sources.
    * Image preview area for the currently selected image.
    * User-friendly controls and status messages.
* **Windows Specific:**
    * Utilizes Windows API for robust wallpaper setting.

## Technologies & Requirements

* **Python 3.x**
* **PyQt6:** For the graphical user interface.
* **Requests:** For making HTTP requests to the Pixabay API and downloading images.
* **Operating System:** Windows (due to reliance on `ctypes` and `winreg` for wallpaper manipulation).

## Setup and Usage

1.  **Clone the Repository (or download the files):**
    ```bash
    git clone [<your-repository-url>](https://github.com/kumar-jit/themo)
    cd TheMoWallpaperChanger
    ```

2.  **Install Dependencies:**
    Make sure you have Python 3 installed. Then, install the required libraries using pip:
    ```bash
    pip install PyQt6 requests
    ```

3.  **Pixabay API Key:**
    The application uses the Pixabay API to fetch online images. The API key is currently hardcoded in `api_handler.py`:
    ```python
    # TheMoWallpaperChanger/api_handler.py
    PIXABAY_API_KEY = "****************************"
    ```
    This is a public key provided in the Pixabay API documentation for testing/example purposes. For more extensive use or if you encounter rate limits, you might want to register for your own free API key at [Pixabay](https://pixabay.com/api/docs/) and replace the existing one.

4.  **Run the Application:**
    Execute the `main.py` script:
    ```bash
    python main.py
    ```

5.  **Using the Application:**
    * **Local Wallpaper Tab:**
        * Click "Browse Image File" to select a single image.
        * Click "Browse Image Folder" to select a folder for a slideshow. Set the interval and click "Start Slideshow".
    * **Online Wallpaper (Pixabay) Tab:**
        * Enter a search term (e.g., "mountains", "beach sunset") in the search bar.
        * Click "Search Images".
        * Results will appear as cards. Click on a card to preview the image.
    * **Common Controls:**
        * Select a "Wallpaper Style" from the dropdown.
        * Click "Set as Wallpaper" to apply the currently previewed image.

## Screenshot
![image](https://github.com/user-attachments/assets/49a52656-9dbf-42b5-867f-b2abfd73eb25)

