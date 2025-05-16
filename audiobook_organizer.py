import sys
import os
import shutil
import re
import requests
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                               QLabel, QLineEdit, QPushButton, QCheckBox, QGroupBox,
                               QTableWidget, QTableWidgetItem, QMessageBox, QStatusBar, QFileDialog,
                               QListWidget, QListWidgetItem, QComboBox, QDialog, QFormLayout)
from PySide6.QtCore import Qt, QObject, Signal, QThread
from mutagen.easyid3 import EasyID3
from mutagen.mp4 import MP4

def sanitize_filename(name):
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        name = name.replace(char, '_')
    return name.strip()

def extract_metadata(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    metadata = {}
    try:
        if ext == '.mp3':
            audio = EasyID3(file_path)
            metadata['artist'] = audio.get('artist', ['Unknown'])[0]
            metadata['title'] = audio.get('title', ['Unknown'])[0]
            metadata['album'] = audio.get('album', ['Unknown'])[0]
            metadata['tracknumber'] = audio.get('tracknumber', ['0'])[0].split('/')[0]
            metadata['year'] = audio.get('date', ['Unknown'])[0]
            metadata['genre'] = audio.get('genre', ['Unknown'])[0]
        elif ext in ['.m4a', '.m4b']:
            audio = MP4(file_path)
            metadata['artist'] = audio.get('\xa9ART', ['Unknown'])[0]
            metadata['title'] = audio.get('\xa9nam', ['Unknown'])[0]
            metadata['album'] = audio.get('\xa9alb', ['Unknown'])[0]
            metadata['tracknumber'] = str(audio.get('trkn', [(0,0)])[0][0])
            metadata['year'] = audio.get('\xa9day', ['Unknown'])[0]
            metadata['genre'] = audio.get('\xa9gen', ['Unknown'])[0]
    except Exception as e:
        print(f"Error reading metadata from {file_path}: {e}")
        metadata = {k: 'Unknown' for k in ['artist', 'title', 'album', 'tracknumber', 'year', 'genre']}
    metadata['ext'] = ext
    return metadata

def generate_new_path(file_path, pattern, output_dir, metadata):
    sanitized_metadata = {k: sanitize_filename(v) for k, v in metadata.items()}
    try:
        relative_path = pattern.format(**sanitized_metadata)
    except KeyError as e:
        raise ValueError(f"Invalid placeholder {e} in pattern")
    new_path = os.path.join(output_dir, relative_path)
    return new_path

def extract_title_and_author_from_filename(filename):
    base_name = os.path.splitext(filename)[0]
    patterns = [
        r'(?P<author>.+?) - (?P<title>.+)',
        r'(?P<title>.+?) by (?P<author>.+)',
        r'(?P<title>.+?) \((?P<author>.+?)\)',
        r'(?P<author>.+?): (?P<title>.+)',
        r'(?P<author>.+?)_+(?P<title>.+)',
        r'(?P<title>.+?)_+\((?P<author>.+?)\)',
    ]
    for pattern in patterns:
        match = re.fullmatch(pattern, base_name)
        if match:
            author = match.group('author').strip() if match.group('author') else ''
            title = match.group('title').strip() if match.group('title') else ''
            title = re.sub(r'(_|\s)\d+$', '', title).strip()
            if author and title:
                return title, author
    title = re.sub(r'(_|\s)\d+$', '', base_name).strip()
    author = ''
    print(f"No author extracted from filename: {filename}")
    return title, author

def search_open_library(title, author=''):
    try:
        query = f'title:"{title}"'
        if author:
            query += f'+author:"{author}"'
        url = f"https://openlibrary.org/search.json?q={query.replace(' ', '+')}"
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        print(f"Open Library API Response for {title}: {data.get('docs', [])[:5]}")
        matches = []
        for doc in data.get('docs', [])[:5]:
            title = doc.get('title')
            if not title or title == 'Unknown':
                print(f"Skipping match for {title}: Missing or invalid title")
                continue
            authors = doc.get('author_name')
            if not authors or not any(author.strip() and author != 'Unknown' for author in authors):
                print(f"Skipping match for {title}: Missing or invalid authors")
                continue
            olid = doc.get('key').split('/')[-1]
            work_data = get_open_library_metadata(olid)
            if not work_data:
                print(f"Skipping match for {title}: Missing or invalid metadata")
                continue
            author_str = ', '.join([author.strip() for author in authors if author.strip() and author != 'Unknown'])
            year = doc.get('first_publish_year')
            display_text = f"{title} by {author_str}" + (f" ({year})" if year else "")
            matches.append((display_text, {'source': 'Open Library', 'olid': olid}))
        return matches
    except Exception as e:
        print(f"Error searching Open Library: {e}")
        return []

def search_open_library_manual(title, author, series):
    try:
        query_parts = []
        if title:
            query_parts.append(f'title:"{title}"')
        if author:
            query_parts.append(f'author:"{author}"')
        if series:
            query_parts.append(f'series:"{series}"')
        query = ' '.join(query_parts)
        url = f"https://openlibrary.org/search.json?q={query.replace(' ', '+')}"
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        print(f"Open Library Manual Search Response: {data.get('docs', [])[:5]}")
        matches = []
        for doc in data.get('docs', [])[:5]:
            book_title = doc.get('title')
            if not book_title or book_title == 'Unknown':
                continue
            authors = doc.get('author_name')
            if not authors or not any(a and a != 'Unknown' for a in authors):
                continue
            olid = doc.get('key').split('/')[-1]
            work_data = get_open_library_metadata(olid)
            if not work_data:
                continue
            author_str = ', '.join([a for a in authors if a and a != 'Unknown'])
            year = doc.get('first_publish_year')
            display_text = f"{book_title} by {author_str}" + (f" ({year})" if year else "")
            matches.append((display_text, {'source': 'Open Library', 'olid': olid}))
        return matches
    except requests.exceptions.RequestException as e:
        print(f"Error in manual Open Library search: {e}")
        return []

def search_google_books(title, author='', api_key=None):
    try:
        search_query = f'intitle:"{title}"'
        if author:
            search_query += f'+inauthor:"{author}"'
        url = f"https://www.googleapis.com/books/v1/volumes?q={search_query}"
        if api_key:
            url += f"&key={api_key}"
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        print(f"Google Books API Response for {search_query}: {data.get('items', [])[:5]}")
        matches = []
        for item in data.get('items', [])[:5]:
            volumeInfo = item.get('volumeInfo', {})
            title = volumeInfo.get('title')
            if not title or title == 'Unknown':
                print(f"Skipping match for {title}: Missing or invalid title")
                continue
            authors = volumeInfo.get('authors')
            if not authors or not any(author.strip() and author != 'Unknown' for author in authors):
                print(f"Skipping match for {title}: Missing or invalid authors")
                continue
            publishedDate = volumeInfo.get('publishedDate', 'Unknown')
            year_match = re.search(r'\d{4}', publishedDate) if publishedDate else None
            year = year_match.group(0) if year_match else None
            display_text = f"{title} by {', '.join([author.strip() for author in authors if author.strip() and author != 'Unknown'])}" + (f" ({year})" if year else "")
            metadata_dict = {
                'title': title,
                'authors': [author.strip() for author in authors if author.strip() and author != 'Unknown'],
                'publishedDate': publishedDate,
                'series': title,
                'source': 'Google Books'
            }
            matches.append((display_text, {'source': 'Google Books', 'metadata': metadata_dict}))
        return matches
    except Exception as e:
        print(f"Error searching Google Books: {e}")
        return []

def search_google_books_manual(title, author, api_key=None):
    try:
        query_parts = []
        if title:
            query_parts.append(f'intitle:"{title}"')
        if author:
            query_parts.append(f'inauthor:"{author}"')
        query = '+'.join(query_parts)
        url = f"https://www.googleapis.com/books/v1/volumes?q={query}"
        if api_key:
            url += f"&key={api_key}"
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        print(f"Google Books Manual Search Response: {data.get('items', [])[:5]}")
        matches = []
        for item in data.get('items', [])[:5]:
            volumeInfo = item.get('volumeInfo', {})
            book_title = volumeInfo.get('title')
            if not book_title or book_title == 'Unknown':
                continue
            authors = volumeInfo.get('authors')
            if not authors or not any(a and a != 'Unknown' for a in authors):
                continue
            publishedDate = volumeInfo.get('publishedDate', 'Unknown')
            year_match = re.search(r'\d{4}', publishedDate) if publishedDate else None
            year = year_match.group(0) if year_match else None
            display_text = f"{book_title} by {', '.join([a for a in authors if a and a != 'Unknown'])}" + (f" ({year})" if year else "")
            metadata_dict = {
                'title': book_title,
                'authors': [a for a in authors if a and a != 'Unknown'],
                'publishedDate': publishedDate,
                'series': book_title,
                'source': 'Google Books'
            }
            matches.append((display_text, {'source': 'Google Books', 'metadata': metadata_dict}))
        return matches
    except requests.exceptions.RequestException as e:
        print(f"Error in manual Google Books search: {e}")
        return []

def get_open_library_metadata(olid):
    try:
        url = f"https://openlibrary.org/works/{olid}.json"
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        print(f"Open Library API Response for OLID {olid}: {data}")
        title = data.get('title')
        if not title or title == 'Unknown':
            print(f"No valid title found for OLID {olid}")
            return None
        authors_data = data.get('authors')
        if not authors_data or not any(author.get('name') and author.get('name') != 'Unknown' for author in authors_data):
            print(f"No valid authors found for OLID {olid}")
            return None
        authors = [author.get('name') for author in authors_data if author.get('name') and author.get('name') != 'Unknown']
        if not authors:
            print(f"No valid authors found for OLID {olid}")
            return None
        publishedDate = data.get('first_publish_date', 'Unknown')
        series = data.get('series')
        series_name = series[0].get('name') if series else ''
        return {
            'title': title,
            'authors': authors,
            'publishedDate': publishedDate,
            'series': series_name,
            'source': 'Open Library'
        }
    except Exception as e:
        print(f"Error fetching metadata for OLID {olid}: {e}")
        return None

def update_metadata(file_path, book_metadata, set_title):
    ext = os.path.splitext(file_path)[1].lower()
    try:
        print(f"Attempting to update metadata for {file_path}: Authors = {book_metadata['authors']}, Series = {book_metadata.get('series', book_metadata['title'])}, Title = {book_metadata['title']}")
        if ext == '.mp3':
            audio = EasyID3(file_path)
            if not audio.get('artist') or audio.get('artist')[0] == 'Unknown':
                if book_metadata['authors'] and book_metadata['authors'][0] != 'Unknown':
                    audio['artist'] = [', '.join(book_metadata['authors'])]
                else:
                    print(f"No valid author for {file_path}, skipping artist update")
                    return False
            if not audio.get('album') or audio.get('album')[0] == 'Unknown':
                audio['album'] = [book_metadata.get('series', book_metadata['title'])]
            if set_title and (not audio.get('title') or audio.get('title')[0] == 'Unknown'):
                audio['title'] = [book_metadata['title']]
            if book_metadata['publishedDate'] and book_metadata['publishedDate'] != 'Unknown':
                if not audio.get('date') or audio.get('date')[0] == 'Unknown':
                    audio['date'] = [book_metadata['publishedDate']]
            audio.save()
        elif ext in ['.m4a', '.m4b']:
            audio = MP4(file_path)
            if '\xa9ART' not in audio or not audio['\xa9ART'] or audio['\xa9ART'][0] == 'Unknown':
                if book_metadata['authors'] and book_metadata['authors'][0] != 'Unknown':
                    audio['\xa9ART'] = book_metadata['authors']
                else:
                    print(f"No valid author for {file_path}, skipping artist update")
                    return False
            if '\xa9alb' not in audio or not audio['\xa9alb'] or audio['\xa9alb'][0] == 'Unknown':
                audio['\xa9alb'] = [book_metadata.get('series', book_metadata['title'])]
            if set_title and ('\xa9nam' not in audio or not audio['\xa9nam'] or audio['\xa9nam'][0] == 'Unknown'):
                audio['\xa9nam'] = [book_metadata['title']]
            if book_metadata['publishedDate'] and book_metadata['publishedDate'] != 'Unknown':
                if '\xa9day' not in audio or not audio['\xa9day'] or audio['\xa9day'][0] == 'Unknown':
                    audio['\xa9day'] = [book_metadata['publishedDate']]
            audio.save()
        updated_metadata = extract_metadata(file_path)
        print(f"Updated metadata for {file_path}: Artist = {updated_metadata['artist']}, Album = {updated_metadata['album']}, Title = {updated_metadata['title']}")
        if updated_metadata['artist'] == 'Unknown' and book_metadata['authors'] and book_metadata['authors'][0] != 'Unknown':
            print(f"Warning: Artist still 'Unknown' for {file_path} despite update attempt")
            return False
        return True
    except Exception as e:
        print(f"Error updating metadata for {file_path}: {e}")
        return False

class ManualSearchDialog(QDialog):
    def __init__(self, source, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Manual Metadata Search")
        self.source = source
        layout = QFormLayout(self)
        self.title_input = QLineEdit(self)
        self.author_input = QLineEdit(self)
        self.series_input = QLineEdit(self)
        layout.addRow("Title:", self.title_input)
        layout.addRow("Author:", self.author_input)
        layout.addRow("Series:", self.series_input)
        if source == "Google Books":
            self.series_input.setEnabled(False)
            self.series_input.setPlaceholderText("Not supported by Google Books")
        buttons = QHBoxLayout()
        ok_button = QPushButton("OK", self)
        cancel_button = QPushButton("Cancel", self)
        ok_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)
        buttons.addWidget(ok_button)
        buttons.addWidget(cancel_button)
        layout.addRow(buttons)
        self.setLayout(layout)

    def get_inputs(self):
        return {
            'title': self.title_input.text().strip(),
            'author': self.author_input.text().strip(),
            'series': self.series_input.text().strip()
        }

class MetadataWorker(QObject):
    progress_signal = Signal(str)
    results_signal = Signal(dict)

    def __init__(self):
        super().__init__()
        self.input_dir = ""
        self.selected_extensions = []
        self.source = ""
        self.api_key = ""

    def set_params(self, input_dir, selected_extensions, source, api_key):
        self.input_dir = input_dir
        self.selected_extensions = selected_extensions
        self.source = source
        self.api_key = api_key

    def process_files(self):
        if not self.input_dir or not self.selected_extensions:
            return
        all_files = []
        for root, _, filenames in os.walk(self.input_dir):
            for filename in filenames:
                if os.path.splitext(filename)[1].lower() in self.selected_extensions:
                    all_files.append(os.path.join(root, filename))
        total_files = len(all_files)
        metadata_matches = {}
        for idx, file_path in enumerate(all_files):
            metadata = extract_metadata(file_path)
            if metadata['artist'] == 'Unknown' or metadata['title'] == 'Unknown' or metadata['album'] == 'Unknown':
                title, author = extract_title_and_author_from_filename(os.path.basename(file_path))
                matches = []
                if self.source == "Open Library":
                    matches = search_open_library(title, author)
                    if not matches:
                        print(f"No valid metadata from Open Library for {title}, trying Google Books")
                        matches.extend(search_google_books(title, author, self.api_key))
                elif self.source == "Google Books":
                    matches = search_google_books(title, author, self.api_key)
                    if not matches:
                        print(f"No valid metadata from Google Books for {title}, trying Open Library")
                        matches.extend(search_open_library(title, author))
                metadata_matches[file_path] = matches
            self.progress_signal.emit(f"Processed {idx+1}/{total_files} files")
        self.results_signal.emit(metadata_matches)

class AudiobookOrganizer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Audiobook File Organizer")
        self.setGeometry(100, 100, 800, 600)
        self.metadata_matches = {}

        self.input_dir_label = QLabel("Input Directory:")
        self.input_dir_text = QLineEdit()
        self.input_dir_button = QPushButton("Browse...")
        self.input_dir_button.clicked.connect(self.select_input_directory)

        self.output_dir_label = QLabel("Output Directory:")
        self.output_dir_text = QLineEdit()
        self.output_dir_button = QPushButton("Browse...")
        self.output_dir_button.clicked.connect(self.select_output_directory)
        self.same_as_input_checkbox = QCheckBox("Use same as input directory")
        self.same_as_input_checkbox.stateChanged.connect(self.toggle_output_dir)

        self.file_types_group = QGroupBox("File Types")
        self.file_types_layout = QHBoxLayout()
        self.file_types = {
            '.mp3': QCheckBox('.mp3'),
            '.m4a': QCheckBox('.m4a'),
            '.m4b': QCheckBox('.m4b'),
            '.aac': QCheckBox('.aac'),
        }
        for checkbox in self.file_types.values():
            self.file_types_layout.addWidget(checkbox)
            checkbox.setChecked(True)
        self.file_types_group.setLayout(self.file_types_layout)

        self.pattern_label = QLabel("Path Pattern (e.g., {artist}/{album}/{title}/{title}.{ext}):")
        self.pattern_text = QLineEdit("{artist}/{album}/{title}/{title}.{ext}")
        self.placeholders_label = QLabel("Available placeholders: {artist}, {title}, {album}, {tracknumber}, {year}, {genre}, {ext}")

        self.metadata_group = QGroupBox("Metadata Matching")
        self.metadata_layout = QVBoxLayout()
        
        self.metadata_source_label = QLabel("Metadata Source:")
        self.metadata_source_combo = QComboBox()
        self.metadata_source_combo.addItems(["Open Library", "Google Books"])
        self.google_api_key_label = QLabel("Google Books API Key:")
        self.google_api_key_text = QLineEdit()
        source_layout = QHBoxLayout()
        source_layout.addWidget(self.metadata_source_label)
        source_layout.addWidget(self.metadata_source_combo)
        source_layout.addWidget(self.google_api_key_label)
        source_layout.addWidget(self.google_api_key_text)
        self.metadata_layout.addLayout(source_layout)
        
        self.missing_metadata_label = QLabel("Files with Missing Metadata:")
        self.missing_metadata_list = QListWidget()
        self.missing_metadata_list.itemSelectionChanged.connect(self.update_match_combo)
        self.match_combo = QComboBox()
        self.set_title_checkbox = QCheckBox("Set title to book title")
        self.apply_button = QPushButton("Apply")
        self.apply_button.clicked.connect(self.apply_match)
        self.skip_button = QPushButton("Skip")
        self.skip_button.clicked.connect(self.skip_file)
        self.manual_search_button = QPushButton("Manual Search")
        self.manual_search_button.clicked.connect(self.perform_manual_search)
        self.next_button = QPushButton("Next")
        self.next_button.clicked.connect(self.next_file)
        self.previous_button = QPushButton("Previous")
        self.previous_button.clicked.connect(self.previous_file)
        self.match_all_button = QPushButton("Match All")
        self.match_all_button.clicked.connect(self.match_all)
        match_controls = QHBoxLayout()
        match_controls.addWidget(self.match_combo)
        match_controls.addWidget(self.apply_button)
        match_controls.addWidget(self.skip_button)
        match_controls.addWidget(self.manual_search_button)
        navigation = QHBoxLayout()
        navigation.addWidget(self.previous_button)
        navigation.addWidget(self.next_button)
        self.metadata_layout.addWidget(self.missing_metadata_label)
        self.metadata_layout.addWidget(self.missing_metadata_list)
        self.metadata_layout.addWidget(self.set_title_checkbox)
        self.metadata_layout.addLayout(match_controls)
        self.metadata_layout.addLayout(navigation)
        self.metadata_layout.addWidget(self.match_all_button)
        self.metadata_group.setLayout(self.metadata_layout)

        self.preview_button = QPushButton("Preview")
        self.preview_button.clicked.connect(self.preview_changes)
        self.preview_table = QTableWidget()
        self.preview_table.setColumnCount(2)
        self.preview_table.setHorizontalHeaderLabels(["Original Path", "New Path"])
        self.preview_table.horizontalHeader().setStretchLastSection(True)

        self.execute_button = QPushButton("Rename and Organize")
        self.execute_button.clicked.connect(self.execute_changes)

        self.status_bar = QStatusBar()

        self.help_button = QPushButton("Help")
        self.help_button.clicked.connect(self.show_help)

        central_widget = QWidget()
        layout = QVBoxLayout()
        input_dir_row = QHBoxLayout()
        input_dir_row.addWidget(self.input_dir_label)
        input_dir_row.addWidget(self.input_dir_text)
        input_dir_row.addWidget(self.input_dir_button)
        layout.addLayout(input_dir_row)

        output_dir_row = QHBoxLayout()
        output_dir_row.addWidget(self.output_dir_label)
        output_dir_row.addWidget(self.output_dir_text)
        output_dir_row.addWidget(self.output_dir_button)
        output_dir_row.addWidget(self.same_as_input_checkbox)
        layout.addLayout(output_dir_row)

        layout.addWidget(self.file_types_group)
        layout.addWidget(self.pattern_label)
        layout.addWidget(self.pattern_text)
        layout.addWidget(self.placeholders_label)
        layout.addWidget(self.metadata_group)
        layout.addWidget(self.preview_button)
        layout.addWidget(self.preview_table)
        layout.addWidget(self.execute_button)
        layout.addWidget(self.help_button)
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)
        self.setStatusBar(self.status_bar)

    def select_input_directory(self):
        dir_path = QFileDialog.getExistingDirectory(self, "Select Input Directory")
        if dir_path:
            self.input_dir_text.setText(dir_path)
            selected_extensions = [ext for ext, cb in self.file_types.items() if cb.isChecked()]
            if not selected_extensions:
                self.status_bar.showMessage("Please select at least one file type")
                return
            source = self.metadata_source_combo.currentText()
            api_key = self.google_api_key_text.text()
            self.metadata_worker = MetadataWorker()
            self.metadata_worker.set_params(dir_path, selected_extensions, source, api_key)
            self.metadata_thread = QThread()
            self.metadata_worker.moveToThread(self.metadata_thread)
            self.metadata_worker.progress_signal.connect(self.update_status_bar)
            self.metadata_worker.results_signal.connect(self.populate_metadata_list)
            self.metadata_thread.started.connect(self.metadata_worker.process_files)
            self.metadata_thread.start()

    def select_output_directory(self):
        dir_path = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if dir_path:
            self.output_dir_text.setText(dir_path)

    def toggle_output_dir(self, state):
        if state == Qt.Checked:
            self.output_dir_text.setText(self.input_dir_text.text())
            self.output_dir_text.setEnabled(False)
            self.output_dir_button.setEnabled(False)
        else:
            self.output_dir_text.setEnabled(True)
            self.output_dir_button.setEnabled(True)

    def update_status_bar(self, message):
        self.status_bar.showMessage(message)

    def populate_metadata_list(self, metadata_matches):
        self.missing_metadata_list.clear()
        self.metadata_matches = metadata_matches
        for file_path in metadata_matches.keys():
            item = QListWidgetItem(os.path.basename(file_path))
            item.setData(Qt.UserRole, file_path)
            self.missing_metadata_list.addItem(item)
        if metadata_matches:
            self.status_bar.showMessage(f"Found {len(metadata_matches)} files with missing metadata")
        else:
            self.status_bar.showMessage("No files with missing metadata found")
        self.metadata_thread.quit()
        self.metadata_thread.wait()

    def update_match_combo(self):
        self.match_combo.clear()
        selected_items = self.missing_metadata_list.selectedItems()
        if selected_items:
            file_path = selected_items[0].data(Qt.UserRole)
            matches = self.metadata_matches.get(file_path, [])
            self.match_combo.addItem("No match", None)
            for display_text, data_dict in matches:
                self.match_combo.addItem(display_text, data_dict)
            self.apply_button.setEnabled(self.match_combo.count() > 1)
        else:
            self.apply_button.setEnabled(False)

    def perform_manual_search(self):
        selected_items = self.missing_metadata_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Warning", "Please select a file to perform a manual search")
            return
        file_path = selected_items[0].data(Qt.UserRole)
        dialog = ManualSearchDialog(self.metadata_source_combo.currentText(), self)
        if dialog.exec():
            inputs = dialog.get_inputs()
            if not any(inputs.values()):
                QMessageBox.warning(self, "Warning", "Please enter at least one search term")
                return
            source = self.metadata_source_combo.currentText()
            api_key = self.google_api_key_text.text()
            if source == "Google Books" and not api_key:
                QMessageBox.warning(self, "Warning", "Please enter a valid Google Books API key")
                return
            try:
                matches = []
                if source == "Open Library":
                    matches = search_open_library_manual(inputs['title'], inputs['author'], inputs['series'])
                else:
                    matches = search_google_books_manual(inputs['title'], inputs['author'], api_key)
                if not matches:
                    QMessageBox.information(self, "No Results", "No matches found. Check your search terms or network connection.")
                    return
                self.metadata_matches[file_path] = matches
                self.update_match_combo()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Search failed: {str(e)}. Please check your internet connection or API key.")
                print(f"Manual search error: {e}")

    def apply_match(self):
        selected_items = self.missing_metadata_list.selectedItems()
        if not selected_items:
            return
        file_path = selected_items[0].data(Qt.UserRole)
        data_dict = self.match_combo.currentData()
        if data_dict:
            source = data_dict['source']
            if source == 'Open Library':
                olid = data_dict['olid']
                book_metadata = get_open_library_metadata(olid)
            elif source == 'Google Books':
                book_metadata = data_dict['metadata']
            else:
                book_metadata = None
            if book_metadata:
                if update_metadata(file_path, book_metadata, self.set_title_checkbox.isChecked()):
                    row = self.missing_metadata_list.currentRow()
                    self.missing_metadata_list.takeItem(row)
                    del self.metadata_matches[file_path]
                    self.status_bar.showMessage(f"Updated metadata for {os.path.basename(file_path)}")
                    if self.missing_metadata_list.count() > 0:
                        self.missing_metadata_list.setCurrentRow(min(row, self.missing_metadata_list.count() - 1))
                else:
                    self.status_bar.showMessage(f"Failed to update metadata for {os.path.basename(file_path)}")
            else:
                self.status_bar.showMessage(f"Failed to fetch book metadata for {os.path.basename(file_path)}")
        else:
            self.status_bar.showMessage("Please select a match")

    def skip_file(self):
        row = self.missing_metadata_list.currentRow()
        if row != -1:
            file_path = self.missing_metadata_list.item(row).data(Qt.UserRole)
            self.missing_metadata_list.takeItem(row)
            del self.metadata_matches[file_path]
            self.status_bar.showMessage(f"Skipped {os.path.basename(file_path)}")
            if self.missing_metadata_list.count() > 0:
                self.missing_metadata_list.setCurrentRow(min(row, self.missing_metadata_list.count() - 1))

    def next_file(self):
        current_row = self.missing_metadata_list.currentRow()
        if current_row != -1 and current_row < self.missing_metadata_list.count() - 1:
            self.missing_metadata_list.setCurrentRow(current_row + 1)

    def previous_file(self):
        current_row = self.missing_metadata_list.currentRow()
        if current_row > 0:
            self.missing_metadata_list.setCurrentRow(current_row - 1)

    def match_all(self):
        for row in range(self.missing_metadata_list.count() - 1, -1, -1):
            item = self.missing_metadata_list.item(row)
            file_path = item.data(Qt.UserRole)
            matches = self.metadata_matches.get(file_path, [])
            if matches:
                display_text, data_dict = matches[0]
                if data_dict['source'] == 'Open Library':
                    olid = data_dict['olid']
                    book_metadata = get_open_library_metadata(olid)
                elif data_dict['source'] == 'Google Books':
                    book_metadata = data_dict['metadata']
                else:
                    continue
                if book_metadata:
                    if update_metadata(file_path, book_metadata, self.set_title_checkbox.isChecked()):
                        self.missing_metadata_list.takeItem(row)
                        del self.metadata_matches[file_path]
                        self.status_bar.showMessage(f"Updated metadata for {os.path.basename(file_path)}")
                    else:
                        self.status_bar.showMessage(f"Failed to update metadata for {os.path.basename(file_path)}")
        if self.missing_metadata_list.count() == 0:
            self.status_bar.showMessage("All files matched or skipped")
        else:
            self.status_bar.showMessage(f"{self.missing_metadata_list.count()} files could not be matched, please review")

    def preview_changes(self):
        input_dir = self.input_dir_text.text()
        if not input_dir:
            QMessageBox.warning(self, "Warning", "Please select input directory")
            return
        output_dir = self.output_dir_text.text() if not self.same_as_input_checkbox.isChecked() else input_dir
        selected_extensions = [ext for ext, cb in self.file_types.items() if cb.isChecked()]
        if not selected_extensions:
            QMessageBox.warning(self, "Warning", "Please select at least one file type")
            return
        pattern = self.pattern_text.text()
        if not pattern:
            QMessageBox.warning(self, "Warning", "Please enter a path pattern")
            return

        files = []
        for root, _, filenames in os.walk(input_dir):
            for filename in filenames:
                if os.path.splitext(filename)[1].lower() in selected_extensions:
                    files.append(os.path.join(root, filename))

        self.preview_table.setRowCount(0)
        for file_path in files:
            metadata = extract_metadata(file_path)
            try:
                new_path = generate_new_path(file_path, pattern, output_dir, metadata)
                row = self.preview_table.rowCount()
                self.preview_table.insertRow(row)
                self.preview_table.setItem(row, 0, QTableWidgetItem(file_path))
                self.preview_table.setItem(row, 1, QTableWidgetItem(new_path))
            except ValueError as e:
                self.status_bar.showMessage(str(e))
                return
        self.preview_table.resizeColumnsToContents()
        self.status_bar.showMessage("Preview generated")

    def execute_changes(self):
        self.preview_changes()
        if self.preview_table.rowCount() == 0:
            QMessageBox.warning(self, "Warning", "No files to process")
            return
        reply = QMessageBox.question(self, "Confirm", "Are you sure you want to rename and organize the files as shown?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.No:
            return
        for row in range(self.preview_table.rowCount()):
            old_path = self.preview_table.item(row, 0).text()
            new_path = self.preview_table.item(row, 1).text()
            base, ext = os.path.splitext(new_path)
            counter = 1
            original_new_path = new_path
            while os.path.exists(new_path):
                new_path = f"{base} ({counter}){ext}"
                counter += 1
            if new_path != original_new_path:
                self.preview_table.setItem(row, 1, QTableWidgetItem(new_path))
            try:
                os.makedirs(os.path.dirname(new_path), exist_ok=True)
                shutil.move(old_path, new_path)
                self.status_bar.showMessage(f"Moved {os.path.basename(old_path)} to {os.path.basename(new_path)}")
            except Exception as e:
                self.status_bar.showMessage(f"Error moving {os.path.basename(old_path)}: {e}")
        self.status_bar.showMessage("Operation completed")

    def show_help(self):
        help_text = """
        Usage Instructions:
        1. Select the input directory containing your audiobook files.
        2. Select the output directory or check 'Use same as input directory'.
        3. Select file types to include (e.g., MP3, M4A, M4B).
        4. Choose a metadata source (Open Library or Google Books) and provide an API key for Google Books.
        5. The 'Files with Missing Metadata' list shows files needing metadata.
        6. Select a file, choose a match from the dropdown, or click 'Manual Search' to enter title/author/series.
        7. Click 'Apply' to update metadata, 'Skip' to ignore, or 'Match All' to auto-match all files.
        8. Use 'Next'/'Previous' to navigate files.
        9. Check 'Set title to book title' to update titles to book titles.
        10. Enter a path pattern (e.g., {artist}/{album}/{title}/{title}.{ext}).
        11. Click 'Preview' to review renaming/organizing changes.
        12. Click 'Rename and Organize' to apply changes.
        Troubleshooting:
        - Check console logs for API responses if no matches appear.
        - Ensure files are writable to avoid save errors.
        - Use clear filenames like 'Author - Title.mp3' for better results.
        - For Google Books, ensure a valid API key is provided.
        - If no results appear, verify internet connection and API status.
        Note: Back up files before modifying, as changes are permanent.
        """
        QMessageBox.information(self, "Help", help_text)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AudiobookOrganizer()
    window.show()
    sys.exit(app.exec())