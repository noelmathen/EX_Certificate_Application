#src/config.py
#!/usr/bin/env python3
"""
Application Configuration Module
Contains all configuration settings and constants
"""

import os
import sys
from pathlib import Path

class AppConfig:
    """Application configuration class"""
    
    # Application metadata
    VERSION = "1.0.0"
    REPO    = "noelmathen/EX_Certificate_Application"
    APP_NAME = "proEXy"
    APP_DESCRIPTION = "Professional EX Certificate Processor"
    
    # GitHub repository
    GITHUB_URL = "https://github.com/noelmathen/EX_Certificate_Application"
    
    # File paths
    ASSETS_DIR = "assets"
    LOGS_DIR = "logs"
    SRC_DIR = "src"
    
    # Output files
    PROSERV_OUTPUT = "Proserv_Certificates.xlsx"
    OMAN_OUTPUT = "Oman_Certificates.xlsx"
    
    # Directory names for classification
    PROSERV_DIR = "Proserv Certificates"
    OMAN_DIR = "Oman Certificates"
    
    # Excel configuration
    PROSERV_SHEET_NAME = "Proserv Certificates"
    OMAN_SHEET_NAME = "Oman Certificates"
    
    # UI Configuration
    WINDOW_WIDTH = 900
    WINDOW_HEIGHT = 700
    MIN_WIDTH = 800
    MIN_HEIGHT = 600
    
    # Processing configuration
    Y_TOLERANCE = 5  # For Oman processor
    EXPECTED_PROSERV_COLS = 18
    
    @classmethod
    def get_asset_path(cls, filename):
        """Get full path to asset file"""
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            base_path = sys._MEIPASS
        else:
            # Running as script
            base_path = os.path.dirname(os.path.abspath(__file__))
            base_path = os.path.dirname(base_path)  # Go up one level from src/
            
        return os.path.join(base_path, cls.ASSETS_DIR, filename)
    
    @classmethod
    def ensure_directories(cls):
        """Ensure required directories exist"""
        directories = [cls.LOGS_DIR, cls.ASSETS_DIR]
        for directory in directories:
            Path(directory).mkdir(exist_ok=True)
