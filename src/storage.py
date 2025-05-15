import os
import io
from pathlib import Path

class StorageHandler:
    def __init__(self):
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        
    def get_file_path(self, filename):
        """Get the full path for a file in the src directory."""
        return os.path.join(self.base_dir, filename)
    
    def file_exists(self, filename):
        """Check if a file exists in the src directory."""
        file_path = self.get_file_path(filename)
        return os.path.exists(file_path)
    
    def save_file(self, file_obj, filename):
        """Save a file to the src directory."""
        file_path = self.get_file_path(filename)
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        file_obj.save(file_path)
        return file_path
    
    def read_file(self, filename):
        """Read a file from the src directory."""
        file_path = self.get_file_path(filename)
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {filename}")
        return open(file_path, 'rb')
    
    def get_file_stream(self, filename):
        """Get a file as a BytesIO stream."""
        with self.read_file(filename) as f:
            stream = io.BytesIO(f.read())
            stream.seek(0)
            return stream 