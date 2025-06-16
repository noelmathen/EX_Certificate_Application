#src/certificate_processor.py
#!/usr/bin/env python3
"""
Certificate Processor Module
Handles the core certificate processing logic
"""

import os
import sys
import glob
import shutil
import logging
import subprocess
from pathlib import Path
import fitz  # PyMuPDF

class CertificateProcessor:
    """Main certificate processing class"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.stats = {
            'proserv': 0,
            'oman': 0,
            'unclassified': 0,
            'errors': 0
        }
        
        # Classification keys
        self.KEY_PROSERV_A = "ELECTRICAL EQUIPMENT"
        self.KEY_PROSERV_B = "HAZARDOUS AREAS"
        self.KEY_OMAN_A = "VISUAL & CLOSE INSPECTION"
        self.KEY_OMAN_B = "REPORT FOR"
        
        # Directory names
        self.PROSERV_DIR = "Proserv Certificates"
        self.OMAN_DIR = "Oman Certificates"
        
    def process_certificates(self, input_folder, progress_callback=None, stats_callback=None):
        """
        Main processing function
        
        Args:
            input_folder (str): Path to folder containing mixed certificates
            progress_callback (callable): Function to call for progress updates
            stats_callback (callable): Function to call for stats updates
        """
        self.logger.info(f"Starting certificate processing from: {input_folder}")
        # ─── switch cwd to the input folder so everything (dirs + Excel) is written there
        try:
            os.chdir(input_folder)
        except Exception as e:
            self.logger.warning(f"Could not change directory to '{input_folder}': {e}")
            
        # Reset statistics
        self.stats = {key: 0 for key in self.stats}
        
        try:
            # Step 1: Validate input
            if not os.path.isdir(input_folder):
                raise ValueError(f"Input folder '{input_folder}' not found")
                
            # Step 2: Get list of PDFs
            pdf_files = sorted(glob.glob(os.path.join(input_folder, "*.pdf")))
            if not pdf_files:
                raise ValueError(f"No PDF files found in '{input_folder}'")
                
            total_files = len(pdf_files)
            self.logger.info(f"Found {total_files} PDF files to process")
            
            # Step 3: Create/clear target directories
            self._setup_directories()
            
            if progress_callback:
                progress_callback(0, total_files, "Classifying certificates...")
            if stats_callback:
                stats_callback(self.stats)
                
            # Step 4: Classify and copy files
            unknown_files = []
            for i, pdf_file in enumerate(pdf_files):
                if progress_callback:
                    if not progress_callback(i + 1, total_files, 
                                           f"Processing: {os.path.basename(pdf_file)}"):
                        self.logger.info("Processing stopped by user")
                        return
                        
                try:
                    classification = self._classify_certificate(pdf_file)
                    filename = os.path.basename(pdf_file)
                    
                    if classification == "proserv":
                        dest_path = os.path.join(self.PROSERV_DIR, filename)
                        shutil.copy2(pdf_file, dest_path)
                        self.stats['proserv'] += 1
                        self.logger.info(f"📄 {filename:<40} → Proserv")
                        
                    elif classification == "oman":
                        dest_path = os.path.join(self.OMAN_DIR, filename)
                        shutil.copy2(pdf_file, dest_path)
                        self.stats['oman'] += 1
                        self.logger.info(f"📄 {filename:<40} → Oman")
                        
                    else:
                        unknown_files.append(filename)
                        self.stats['unclassified'] += 1
                        self.logger.warning(f"⚠️  {filename:<40} → Unclassified")
                        
                except Exception as e:
                    self.stats['errors'] += 1
                    self.logger.error(f"❌ Error processing {os.path.basename(pdf_file)}: {e}")
                    
                # Update stats
                if stats_callback:
                    stats_callback(self.stats)
                    
            # Report unclassified files
            if unknown_files:
                self.logger.warning(f"⚠️ Unclassified files ({len(unknown_files)}): {', '.join(unknown_files)}")
                
            # Step 5: Run specialized processors
            if progress_callback:
                progress_callback(total_files, total_files, "Running Proserv processor...")
                
            self._run_proserv_processor()
            
            if progress_callback:
                progress_callback(total_files, total_files, "Running Oman processor...")
                
            self._run_oman_processor()
            
            if progress_callback:
                progress_callback(total_files, total_files, "Processing completed!")
                
            self.logger.info("🎉 Certificate processing completed successfully!")
            self._log_final_statistics()
            
        except Exception as e:
            self.logger.error(f"Processing failed: {e}")
            raise
            
    def _classify_certificate(self, pdf_path):
        """
        Classify a PDF certificate as 'proserv', 'oman', or None
        
        Args:
            pdf_path (str): Path to PDF file
            
        Returns:
            str or None: Classification result
        """
        try:
            with fitz.open(pdf_path) as doc:
                if len(doc) == 0:
                    return None
                text = doc[0].get_text().upper()
                
        except Exception as e:
            self.logger.error(f"Cannot open {os.path.basename(pdf_path)}: {e}")
            return None
            
        # Check for Proserv indicators
        if self.KEY_PROSERV_A in text and self.KEY_PROSERV_B in text:
            return "proserv"
            
        # Check for Oman indicators
        if self.KEY_OMAN_A in text and self.KEY_OMAN_B in text:
            return "oman"
            
        return None
        
    def _setup_directories(self):
        """Create/clear target directories"""
        for directory in [self.PROSERV_DIR, self.OMAN_DIR]:
            if os.path.exists(directory):
                shutil.rmtree(directory)
            os.makedirs(directory, exist_ok=True)
            self.logger.info(f"Created directory: {directory}")
            
    def _run_proserv_processor(self):
        """Run the Proserv certificate processor"""
        try:
            self.logger.info("▶️  Running Proserv processor...")
            
            # Import and run proserv processor
            from src.proserv_processor import ProservProcessor
            processor = ProservProcessor()
            processor.process()
            
            self.logger.info("✅ Proserv processor completed")
            
        except Exception as e:
            self.logger.error(f"❌ Proserv processor failed: {e}")
            raise
            
    def _run_oman_processor(self):
        """Run the Oman certificate processor"""
        try:
            self.logger.info("▶️  Running Oman processor...")
            
            # Import and run oman processor
            from src.oman_processor import OmanProcessor
            processor = OmanProcessor()
            processor.process()
            
            self.logger.info("✅ Oman processor completed")
            
        except Exception as e:
            self.logger.error(f"❌ Oman processor failed: {e}")
            raise
            
    def _log_final_statistics(self):
        """Log final processing statistics"""
        self.logger.info("📊 Final Statistics:")
        self.logger.info(f"   • Proserv certificates: {self.stats['proserv']}")
        self.logger.info(f"   • Oman certificates: {self.stats['oman']}")
        self.logger.info(f"   • Unclassified files: {self.stats['unclassified']}")
        self.logger.info(f"   • Processing errors: {self.stats['errors']}")