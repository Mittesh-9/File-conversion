from flask import Flask, request, jsonify, render_template_string, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import os
import tempfile
import mimetypes
from pathlib import Path
import uuid
import threading
import time
from datetime import datetime, timedelta
import json

# File conversion libraries
from PIL import Image, ImageEnhance, ImageFilter
import pandas as pd
from docx import Document
from fpdf import FPDF
import zipfile
import tarfile
import py7zr
import moviepy.editor as mp
from pydub import AudioSegment
import markdown
import csv
import xml.etree.ElementTree as ET
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import openpyxl
from openpyxl.styles import Font, PatternFill
import sqlite3
import requests
import base64
from io import BytesIO

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this'
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max file size

# Global variables for tracking conversions
conversion_status = {}
conversion_results = {}

class FileConverter:
    def __init__(self):
        self.temp_dir = tempfile.gettempdir()
        self.supported_formats = {
            'image': {
                'input': ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp'],
                'output': ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp', '.pdf']
            },
            'document': {
                'input': ['.txt', '.docx', '.md', '.html', '.csv', '.xlsx'],
                'output': ['.pdf', '.txt', '.docx', '.html', '.md', '.csv', '.xlsx']
            },
            'audio': {
                'input': ['.mp3', '.wav', '.flac', '.ogg', '.m4a', '.aac'],
                'output': ['.mp3', '.wav', '.flac', '.ogg', '.m4a']
            },
            'video': {
                'input': ['.mp4', '.avi', '.mov', '.wmv', '.flv', '.mkv'],
                'output': ['.mp4', '.avi', '.mov', '.webm', '.gif']
            },
            'archive': {
                'input': ['.zip', '.rar', '.7z', '.tar', '.gz'],
                'output': ['.zip', '.7z', '.tar']
            }
        }
    
    def get_file_category(self, filename):
        ext = Path(filename).suffix.lower()
        for category, formats in self.supported_formats.items():
            if ext in formats['input']:
                return category
        return None
    
    def convert_image(self, input_path, output_path, output_format, quality=95, resize=None, enhance=None):
        with Image.open(input_path) as img:
            # Apply enhancements if specified
            if enhance:
                if enhance.get('brightness'):
                    enhancer = ImageEnhance.Brightness(img)
                    img = enhancer.enhance(enhance['brightness'])
                if enhance.get('contrast'):
                    enhancer = ImageEnhance.Contrast(img)
                    img = enhancer.enhance(enhance['contrast'])
                if enhance.get('sharpness'):
                    enhancer = ImageEnhance.Sharpness(img)
                    img = enhancer.enhance(enhance['sharpness'])
                if enhance.get('blur'):
                    img = img.filter(ImageFilter.GaussianBlur(radius=enhance['blur']))
            
            # Resize if specified
            if resize:
                img = img.resize(resize, Image.Resampling.LANCZOS)
            
            # Convert to RGB if saving as JPEG
            if output_format.lower() in ['jpg', 'jpeg'] and img.mode in ('RGBA', 'LA', 'P'):
                img = img.convert('RGB')
            
            # Special handling for PDF conversion
            if output_format.lower() == 'pdf':
                if img.mode in ('RGBA', 'LA', 'P'):
                    img = img.convert('RGB')
                img.save(output_path, 'PDF', quality=quality)
            else:
                img.save(output_path, quality=quality, optimize=True)
    
    def convert_document(self, input_path, output_path, output_format):
        input_ext = Path(input_path).suffix.lower()
        
        if input_ext == '.txt' and output_format == 'pdf':
            self._txt_to_pdf(input_path, output_path)
        elif input_ext == '.csv' and output_format == 'xlsx':
            df = pd.read_csv(input_path)
            df.to_excel(output_path, index=False, engine='openpyxl')
        elif input_ext == '.xlsx' and output_format == 'csv':
            df = pd.read_excel(input_path, engine='openpyxl')
            df.to_csv(output_path, index=False)
        elif input_ext == '.md' and output_format == 'html':
            with open(input_path, 'r', encoding='utf-8') as f:
                md_content = f.read()
            html_content = markdown.markdown(md_content, extensions=['tables', 'fenced_code'])
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(f"<!DOCTYPE html><html><head><meta charset='utf-8'></head><body>{html_content}</body></html>")
        elif input_ext == '.docx' and output_format == 'txt':
            doc = Document(input_path)
            with open(output_path, 'w', encoding='utf-8') as f:
                for paragraph in doc.paragraphs:
                    f.write(paragraph.text + '\n')
    
    def _txt_to_pdf(self, input_path, output_path):
        c = canvas.Canvas(output_path, pagesize=letter)
        width, height = letter
        
        with open(input_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        y_position = height - 72  # Start from top with margin
        line_height = 14
        
        for line in lines:
            if y_position < 72:  # If near bottom of page, start new page
                c.showPage()
                y_position = height - 72
            
            c.drawString(72, y_position, line.strip())
            y_position -= line_height
        
        c.save()
    
    def convert_audio(self, input_path, output_path, output_format, bitrate="192k"):
        audio = AudioSegment.from_file(input_path)
        
        export_params = {"format": output_format}
        if output_format == "mp3":
            export_params["bitrate"] = bitrate
        
        audio.export(output_path, **export_params)
    
    def convert_video(self, input_path, output_path, output_format, quality=None):
        video = mp.VideoFileClip(input_path)
        
        if output_format == "gif":
            video.write_gif(output_path, fps=15)
        else:
            write_params = {}
            if quality:
                write_params["bitrate"] = quality
            
            video.write_videofile(output_path, **write_params)
        
        video.close()
    
    def convert_archive(self, input_path, output_path, output_format):
        input_ext = Path(input_path).suffix.lower()
        
        # Extract files to temp directory first
        temp_extract_dir = os.path.join(self.temp_dir, str(uuid.uuid4()))
        os.makedirs(temp_extract_dir, exist_ok=True)
        
        try:
            # Extract based on input format
            if input_ext == '.zip':
                with zipfile.ZipFile(input_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_extract_dir)
            elif input_ext == '.7z':
                with py7zr.SevenZipFile(input_path, mode='r') as archive:
                    archive.extractall(temp_extract_dir)
            elif input_ext in ['.tar', '.gz']:
                with tarfile.open(input_path, 'r:*') as tar_ref:
                    tar_ref.extractall(temp_extract_dir)
            
            # Create new archive in desired format
            if output_format == 'zip':
                with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
                    for root, dirs, files in os.walk(temp_extract_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, temp_extract_dir)
                            zip_ref.write(file_path, arcname)
            elif output_format == '7z':
                with py7zr.SevenZipFile(output_path, 'w') as archive:
                    for root, dirs, files in os.walk(temp_extract_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, temp_extract_dir)
                            archive.write(file_path, arcname)
        
        finally:
            # Clean up temp directory
            import shutil
            shutil.rmtree(temp_extract_dir, ignore_errors=True)

converter = FileConverter()

def convert_file_async(conversion_id, input_path, output_path, file_category, output_format, options=None):
    """Asynchronous file conversion function"""
    try:
        conversion_status[conversion_id] = {
            'status': 'processing',
            'progress': 0,
            'message': 'Starting conversion...',
            'start_time': datetime.now()
        }
        
        # Update progress
        conversion_status[conversion_id]['progress'] = 25
        conversion_status[conversion_id]['message'] = 'Processing file...'
        
        # Perform conversion based on file category
        if file_category == 'image':
            enhance_options = None
            resize_options = None
            quality = 95
            
            if options:
                if 'enhance' in options:
                    enhance_options = options['enhance']
                if 'resize' in options:
                    resize_options = tuple(options['resize'])
                if 'quality' in options:
                    quality = options['quality']
            
            converter.convert_image(input_path, output_path, output_format, 
                                  quality=quality, resize=resize_options, enhance=enhance_options)
        
        elif file_category == 'document':
            converter.convert_document(input_path, output_path, output_format)
        
        elif file_category == 'audio':
            bitrate = options.get('bitrate', '192k') if options else '192k'
            converter.convert_audio(input_path, output_path, output_format, bitrate=bitrate)
        
        elif file_category == 'video':
            quality = options.get('quality') if options else None
            converter.convert_video(input_path, output_path, output_format, quality=quality)
        
        elif file_category == 'archive':
            converter.convert_archive(input_path, output_path, output_format)
        
        # Update progress
        conversion_status[conversion_id]['progress'] = 90
        conversion_status[conversion_id]['message'] = 'Finalizing...'
        
        # Check if output file was created successfully
        if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            conversion_status[conversion_id] = {
                'status': 'completed',
                'progress': 100,
                'message': 'Conversion completed successfully!',
                'end_time': datetime.now()
            }
            conversion_results[conversion_id] = {
                'output_path': output_path,
                'original_name': options.get('original_name', 'converted_file') if options else 'converted_file',
                'output_format': output_format
            }
        else:
            raise Exception("Output file was not created or is empty")
    
    except Exception as e:
        conversion_status[conversion_id] = {
            'status': 'error',
            'progress': 0,
            'message': f'Error: {str(e)}',
            'error_time': datetime.now()
        }

@app.route('/')
def index():
    return render_template_string("""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Advanced File Converter</title>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }

        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }

        .header h1 {
            font-size: 2.5rem;
            margin-bottom: 10px;
        }

        .header p {
            opacity: 0.9;
            font-size: 1.1rem;
        }

        .main-content {
            padding: 40px;
        }

        .upload-section {
            border: 3px dashed #ddd;
            border-radius: 15px;
            padding: 60px 20px;
            text-align: center;
            transition: all 0.3s ease;
            margin-bottom: 30px;
            cursor: pointer;
        }

        .upload-section.dragover {
            border-color: #667eea;
            background-color: #f8f9ff;
        }

        .upload-section:hover {
            border-color: #667eea;
            background-color: #f8f9ff;
        }

        .upload-icon {
            font-size: 4rem;
            color: #ddd;
            margin-bottom: 20px;
        }

        .upload-section.dragover .upload-icon,
        .upload-section:hover .upload-icon {
            color: #667eea;
        }

        .upload-text {
            font-size: 1.3rem;
            color: #666;
            margin-bottom: 15px;
        }

        .file-input {
            display: none;
        }

        .btn {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            padding: 12px 30px;
            border-radius: 25px;
            cursor: pointer;
            font-size: 1rem;
            transition: all 0.3s ease;
            display: inline-block;
            text-decoration: none;
        }

        .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(102, 126, 234, 0.3);
        }

        .conversion-options {
            display: none;
            background: #f8f9ff;
            border-radius: 15px;
            padding: 30px;
            margin-bottom: 30px;
        }

        .option-group {
            margin-bottom: 25px;
        }

        .option-group label {
            display: block;
            margin-bottom: 8px;
            font-weight: 600;
            color: #333;
        }

        .option-group select,
        .option-group input {
            width: 100%;
            padding: 10px;
            border: 2px solid #ddd;
            border-radius: 8px;
            font-size: 1rem;
            transition: border-color 0.3s ease;
        }

        .option-group select:focus,
        .option-group input:focus {
            outline: none;
            border-color: #667eea;
        }

        .advanced-options {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
        }

        .progress-container {
            display: none;
            background: #f8f9ff;
            border-radius: 15px;
            padding: 30px;
            margin-bottom: 30px;
        }

        .progress-bar {
            background: #e0e0e0;
            border-radius: 10px;
            overflow: hidden;
            height: 20px;
            margin-bottom: 15px;
        }

        .progress-fill {
            background: linear-gradient(90deg, #667eea, #764ba2);
            height: 100%;
            width: 0%;
            transition: width 0.3s ease;
        }

        .progress-text {
            text-align: center;
            color: #666;
            font-weight: 600;
        }

        .file-info {
            display: none;
            background: #e8f4fd;
            border-radius: 10px;
            padding: 20px;
            margin-bottom: 20px;
        }

        .file-info h3 {
            color: #2196F3;
            margin-bottom: 10px;
        }

        .result-section {
            display: none;
            background: #e8f5e8;
            border-radius: 15px;
            padding: 30px;
            text-align: center;
        }

        .result-section.error {
            background: #ffeaea;
        }

        .result-section h3 {
            color: #4caf50;
            margin-bottom: 15px;
        }

        .result-section.error h3 {
            color: #f44336;
        }

        .supported-formats {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-top: 40px;
        }

        .format-category {
            background: #f8f9ff;
            border-radius: 10px;
            padding: 20px;
            text-align: center;
        }

        .format-category h3 {
            color: #667eea;
            margin-bottom: 15px;
        }

        .format-list {
            font-size: 0.9rem;
            color: #666;
            line-height: 1.5;
        }

        .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #667eea;
            border-radius: 50%;
            width: 50px;
            height: 50px;
            animation: spin 1s linear infinite;
            margin: 20px auto;
        }

        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }

        .drag-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(102, 126, 234, 0.1);
            display: none;
            justify-content: center;
            align-items: center;
            z-index: 1000;
        }

        .drag-message {
            background: white;
            padding: 40px;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.2);
            text-align: center;
            font-size: 1.5rem;
            color: #667eea;
        }

        @media (max-width: 768px) {
            .main-content {
                padding: 20px;
            }
            
            .advanced-options {
                grid-template-columns: 1fr;
            }
        }
    </style>
</head>
<body>
    <div class="drag-overlay" id="dragOverlay">
        <div class="drag-message">
            <i class="fas fa-cloud-upload-alt" style="font-size: 3rem; margin-bottom: 20px;"></i>
            <div>Drop your file here to convert</div>
        </div>
    </div>

    <div class="container">
        <div class="header">
            <h1><i class="fas fa-exchange-alt"></i> Advanced File Converter</h1>
            <p>Convert files between different formats with advanced options and real-time progress tracking</p>
        </div>

        <div class="main-content">
            <div class="upload-section" id="uploadSection">
                <div class="upload-icon">
                    <i class="fas fa-cloud-upload-alt"></i>
                </div>
                <div class="upload-text">
                    Drag & Drop your file here or click to browse
                </div>
                <button class="btn" onclick="document.getElementById('fileInput').click()">
                    <i class="fas fa-folder-open"></i> Choose File
                </button>
                <input type="file" id="fileInput" class="file-input" multiple>
            </div>

            <div class="file-info" id="fileInfo">
                <h3><i class="fas fa-info-circle"></i> File Information</h3>
                <div id="fileDetails"></div>
            </div>

            <div class="conversion-options" id="conversionOptions">
                <h3><i class="fas fa-cogs"></i> Conversion Options</h3>
                <div class="option-group">
                    <label for="outputFormat">Output Format:</label>
                    <select id="outputFormat">
                        <option value="">Select format...</option>
                    </select>
                </div>
                
                <div class="advanced-options" id="advancedOptions">
                    <!-- Dynamic options will be inserted here -->
                </div>

                <button class="btn" onclick="startConversion()" style="margin-top: 20px;">
                    <i class="fas fa-play"></i> Start Conversion
                </button>
            </div>

            <div class="progress-container" id="progressContainer">
                <h3><i class="fas fa-spinner fa-spin"></i> Converting...</h3>
                <div class="progress-bar">
                    <div class="progress-fill" id="progressFill"></div>
                </div>
                <div class="progress-text" id="progressText">Preparing...</div>
            </div>

            <div class="result-section" id="resultSection">
                <h3 id="resultTitle"></h3>
                <div id="resultMessage"></div>
                <button class="btn" id="downloadBtn" style="display: none; margin-top: 15px;">
                    <i class="fas fa-download"></i> Download Converted File
                </button>
                <button class="btn" onclick="resetConverter()" style="margin-top: 15px; margin-left: 10px;">
                    <i class="fas fa-redo"></i> Convert Another File
                </button>
            </div>

            <div class="supported-formats">
                <div class="format-category">
                    <h3><i class="fas fa-image"></i> Images</h3>
                    <div class="format-list">
                        JPG, PNG, BMP, GIF, TIFF, WebP → PDF, JPG, PNG, etc.
                    </div>
                </div>
                <div class="format-category">
                    <h3><i class="fas fa-file-alt"></i> Documents</h3>
                    <div class="format-list">
                        TXT, DOCX, MD, HTML, CSV, XLSX → PDF, TXT, DOCX, etc.
                    </div>
                </div>
                <div class="format-category">
                    <h3><i class="fas fa-music"></i> Audio</h3>
                    <div class="format-list">
                        MP3, WAV, FLAC, OGG, M4A → MP3, WAV, FLAC, etc.
                    </div>
                </div>
                <div class="format-category">
                    <h3><i class="fas fa-video"></i> Video</h3>
                    <div class="format-list">
                        MP4, AVI, MOV, WMV, FLV → MP4, AVI, WebM, GIF
                    </div>
                </div>
                <div class="format-category">
                    <h3><i class="fas fa-archive"></i> Archives</h3>
                    <div class="format-list">
                        ZIP, RAR, 7Z, TAR, GZ → ZIP, 7Z, TAR
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
        let selectedFile = null;
        let conversionId = null;
        let statusCheckInterval = null;

        // File input handling
        document.getElementById('fileInput').addEventListener('change', handleFileSelect);
        
        // Drag and drop handling
        const uploadSection = document.getElementById('uploadSection');
        const dragOverlay = document.getElementById('dragOverlay');
        
        document.addEventListener('dragenter', (e) => {
            e.preventDefault();
            dragOverlay.style.display = 'flex';
        });
        
        document.addEventListener('dragover', (e) => {
            e.preventDefault();
        });
        
        document.addEventListener('dragleave', (e) => {
            if (e.clientX === 0 && e.clientY === 0) {
                dragOverlay.style.display = 'none';
            }
        });
        
        document.addEventListener('drop', (e) => {
            e.preventDefault();
            dragOverlay.style.display = 'none';
            
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                handleFileSelect({target: {files: files}});
            }
        });

        uploadSection.addEventListener('click', () => {
            document.getElementById('fileInput').click();
        });

        function handleFileSelect(event) {
            const file = event.target.files[0];
            if (!file) return;

            selectedFile = file;
            
            // Show file info
            const fileInfo = document.getElementById('fileInfo');
            const fileDetails = document.getElementById('fileDetails');
            
            fileDetails.innerHTML = `
                <p><strong>Name:</strong> ${file.name}</p>
                <p><strong>Size:</strong> ${formatFileSize(file.size)}</p>
                <p><strong>Type:</strong> ${file.type || 'Unknown'}</p>
                <p><strong>Last Modified:</strong> ${new Date(file.lastModified).toLocaleString()}</p>
            `;
            
            fileInfo.style.display = 'block';
            
            // Get supported output formats
            getOutputFormats(file.name);
        }

        function formatFileSize(bytes) {
            if (bytes === 0) return '0 Bytes';
            const k = 1024;
            const sizes = ['Bytes', 'KB', 'MB', 'GB'];
            const i = Math.floor(Math.log(bytes) / Math.log(k));
            return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
        }

        function getOutputFormats(filename) {
            fetch('/get_output_formats', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({filename: filename})
            })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    populateOutputFormats(data.formats, data.category);
                } else {
                    alert('Unsupported file format');
                }
            });
        }

        function populateOutputFormats(formats, category) {
            const select = document.getElementById('outputFormat');
            select.innerHTML = '<option value="">Select format...</option>';
            
            formats.forEach(format => {
                const option = document.createElement('option');
                option.value = format.replace('.', '');
                option.textContent = format.toUpperCase();
                select.appendChild(option);
            });
            
            select.addEventListener('change', () => {
                if (select.value) {
                    showAdvancedOptions(category, select.value);
                }
            });
            
            document.getElementById('conversionOptions').style.display = 'block';
        }

        function showAdvancedOptions(category, outputFormat) {
            const advancedOptions = document.getElementById('advancedOptions');
            advancedOptions.innerHTML = '';
            
            if (category === 'image') {
                advancedOptions.innerHTML = `
                    <div class="option-group">
                        <label for="quality">Quality (1-100):</label>
                        <input type="range" id="quality" min="1" max="100" value="95">
                        <span id="qualityValue">95</span>
                    </div>
                    <div class="option-group">
                        <label for="resize">Resize (Width x Height):</label>
                        <input type="text" id="resize" placeholder="e.g., 800x600">
                    </div>
                    <div class="option-group">
                        <label for="brightness">Brightness:</label>
                        <input type="range" id="brightness" min="0.5" max="2" step="0.1" value="1">
                        <span id="brightnessValue">1.0</span>
                    </div>
                    <div class="option-group">
                        <label for="contrast">Contrast:</label>
                        <input type="range" id="contrast" min="0.5" max="2" step="0.1" value="1">
                        <span id="contrastValue">1.0</span>
                    </div>
                `;
                
                // Add event listeners for range inputs
                ['quality', 'brightness', 'contrast'].forEach(id => {
                    const input = document.getElementById(id);
                    const valueSpan = document.getElementById(id + 'Value');
                    input.addEventListener('input', () => {
                        valueSpan.textContent = input.value;
                    });
                });
                
            } else if (category === 'audio') {
                advancedOptions.innerHTML = `
                    <div class="option-group">
                        <label for="bitrate">Bitrate:</label>
                        <select id="bitrate">
                            <option value="128k">128 kbps</option>
                            <option value="192k" selected>192 kbps</option>
                            <option value="256k">256 kbps</option>
                            <option value="320k">320 kbps</option>
                        </select>
                    </div>
                `;
            } else if (category === 'video') {
                advancedOptions.innerHTML = `
                    <div class="option-group">
                        <label for="videoQuality">Quality:</label>
                        <select id="videoQuality">
                            <option value="">Default</option>
                            <option value="1000k">Low (1000k)</option>
                            <option value="2000k">Medium (2000k)</option>
                            <option value="5000k">High (5000k)</option>
                        </select>
                    </div>
                `;
            }
        }

        function startConversion() {
            if (!selectedFile) {
                alert('Please select a file first');
                return;
            }
            
            const outputFormat = document.getElementById('outputFormat').value;
            if (!outputFormat) {
                alert('Please select an output format');
                return;
            }
            
            // Collect conversion options
            const options = {};
            
            // Image options
            const quality = document.getElementById('quality');
            if (quality) options.quality = parseInt(quality.value);
            
            const resize = document.getElementById('resize');
            if (resize && resize.value) {
                const dimensions = resize.value.split('x');
                if (dimensions.length === 2) {
                    options.resize = [parseInt(dimensions[0]), parseInt(dimensions[1])];
                }
            }
            
            const brightness = document.getElementById('brightness');
            const contrast = document.getElementById('contrast');
            if (brightness || contrast) {
                options.enhance = {};
                if (brightness) options.enhance.brightness = parseFloat(brightness.value);
                if (contrast) options.enhance.contrast = parseFloat(contrast.value);
            }
            
            // Audio options
            const bitrate = document.getElementById('bitrate');
            if (bitrate) options.bitrate = bitrate.value;
            
            // Video options
            const videoQuality = document.getElementById('videoQuality');
            if (videoQuality && videoQuality.value) options.quality = videoQuality.value;
            
            // Hide options and show progress
            document.getElementById('conversionOptions').style.display = 'none';
            document.getElementById('progressContainer').style.display = 'block';
            
            // Upload file and start conversion
            uploadAndConvert(selectedFile, outputFormat, options);
        }

        function uploadAndConvert(file, outputFormat, options) {
            const formData = new FormData();
            formData.append('file', file);
            formData.append('output_format', outputFormat);
            formData.append('options', JSON.stringify(options));
            
            fetch('/convert', {
                method: 'POST',
                body: formData
            })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    conversionId = data.conversion_id;
                    checkConversionStatus();
                } else {
                    showError(data.error || 'Conversion failed');
                }
            })
            .catch(error => {
                showError('Network error: ' + error.message);
            });
        }

        function checkConversionStatus() {
            if (!conversionId) return;
            
            statusCheckInterval = setInterval(() => {
                fetch(`/status/${conversionId}`)
                .then(response => response.json())
                .then(data => {
                    updateProgress(data.progress, data.message);
                    
                    if (data.status === 'completed') {
                        clearInterval(statusCheckInterval);
                        showSuccess(data.message);
                    } else if (data.status === 'error') {
                        clearInterval(statusCheckInterval);
                        showError(data.message);
                    }
                })
                .catch(error => {
                    clearInterval(statusCheckInterval);
                    showError('Status check failed: ' + error.message);
                });
            }, 1000);
        }

        function updateProgress(progress, message) {
            document.getElementById('progressFill').style.width = progress + '%';
            document.getElementById('progressText').textContent = message;
        }

        function showSuccess(message) {
            document.getElementById('progressContainer').style.display = 'none';
            
            const resultSection = document.getElementById('resultSection');
            const resultTitle = document.getElementById('resultTitle');
            const resultMessage = document.getElementById('resultMessage');
            const downloadBtn = document.getElementById('downloadBtn');
            
            resultSection.className = 'result-section';
            resultTitle.innerHTML = '<i class="fas fa-check-circle"></i> Conversion Successful!';
            resultMessage.textContent = message;
            
            downloadBtn.style.display = 'inline-block';
            downloadBtn.onclick = () => downloadFile();
            
            resultSection.style.display = 'block';
        }

        function showError(message) {
            document.getElementById('progressContainer').style.display = 'none';
            
            const resultSection = document.getElementById('resultSection');
            const resultTitle = document.getElementById('resultTitle');
            const resultMessage = document.getElementById('resultMessage');
            
            resultSection.className = 'result-section error';
            resultTitle.innerHTML = '<i class="fas fa-exclamation-circle"></i> Conversion Failed';
            resultMessage.textContent = message;
            
            document.getElementById('downloadBtn').style.display = 'none';
            
            resultSection.style.display = 'block';
        }

        function downloadFile() {
            if (conversionId) {
                window.open(`/download/${conversionId}`, '_blank');
            }
        }

        function resetConverter() {
            selectedFile = null;
            conversionId = null;
            
            if (statusCheckInterval) {
                clearInterval(statusCheckInterval);
                statusCheckInterval = null;
            }
            
            document.getElementById('fileInput').value = '';
            document.getElementById('fileInfo').style.display = 'none';
            document.getElementById('conversionOptions').style.display = 'none';
            document.getElementById('progressContainer').style.display = 'none';
            document.getElementById('resultSection').style.display = 'none';
            
            // Reset progress
            document.getElementById('progressFill').style.width = '0%';
            document.getElementById('progressText').textContent = 'Preparing...';
        }
    </script>
</body>
</html>
""")

@app.route('/get_output_formats', methods=['POST'])
def get_output_formats():
    try:
        data = request.get_json()
        filename = data.get('filename', '')
        
        category = converter.get_file_category(filename)
        if not category:
            return jsonify({'success': False, 'error': 'Unsupported file format'})
        
        formats = converter.supported_formats[category]['output']
        return jsonify({
            'success': True, 
            'formats': formats, 
            'category': category
        })
    
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/convert', methods=['POST'])
def convert_file():
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'No file uploaded'})
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': 'No file selected'})
        
        output_format = request.form.get('output_format', '')
        options_str = request.form.get('options', '{}')
        
        try:
            options = json.loads(options_str)
        except:
            options = {}
        
        # Add original filename to options
        options['original_name'] = file.filename
        
        # Generate unique conversion ID
        conversion_id = str(uuid.uuid4())
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        input_path = os.path.join(converter.temp_dir, f"{conversion_id}_{filename}")
        file.save(input_path)
        
        # Determine output file path
        file_category = converter.get_file_category(filename)
        if not file_category:
            os.remove(input_path)
            return jsonify({'success': False, 'error': 'Unsupported file format'})
        
        base_name = Path(filename).stem
        output_filename = f"{base_name}.{output_format}"
        output_path = os.path.join(converter.temp_dir, f"{conversion_id}_{output_filename}")
        
        # Start conversion in background thread
        thread = threading.Thread(
            target=convert_file_async,
            args=(conversion_id, input_path, output_path, file_category, output_format, options)
        )
        thread.daemon = True
        thread.start()
        
        return jsonify({'success': True, 'conversion_id': conversion_id})
    
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/status/<conversion_id>')
def get_conversion_status(conversion_id):
    status = conversion_status.get(conversion_id, {
        'status': 'not_found',
        'progress': 0,
        'message': 'Conversion not found'
    })
    return jsonify(status)

@app.route('/download/<conversion_id>')
def download_file(conversion_id):
    try:
        if conversion_id not in conversion_results:
            return "File not found", 404
        
        result = conversion_results[conversion_id]
        output_path = result['output_path']
        
        if not os.path.exists(output_path):
            return "File not found", 404
        
        # Determine the download filename
        original_name = result.get('original_name', 'converted_file')
        base_name = Path(original_name).stem
        output_format = result.get('output_format', 'bin')
        download_name = f"{base_name}.{output_format}"
        
        return send_file(
            output_path,
            as_attachment=True,
            download_name=download_name
        )
    
    except Exception as e:
        return f"Error downloading file: {str(e)}", 500

@app.route('/cleanup')
def cleanup_files():
    """Clean up old conversion files"""
    try:
        current_time = datetime.now()
        cleanup_count = 0
        
        # Clean up files older than 1 hour
        for conversion_id in list(conversion_status.keys()):
            status = conversion_status[conversion_id]
            
            # Check if conversion is old
            start_time = status.get('start_time')
            end_time = status.get('end_time') or status.get('error_time')
            
            if start_time and (current_time - start_time > timedelta(hours=1)):
                # Clean up files
                if conversion_id in conversion_results:
                    output_path = conversion_results[conversion_id]['output_path']
                    if os.path.exists(output_path):
                        os.remove(output_path)
                    
                    # Also clean up input file
                    input_pattern = os.path.join(converter.temp_dir, f"{conversion_id}_*")
                    import glob
                    for file_path in glob.glob(input_pattern):
                        if os.path.exists(file_path):
                            os.remove(file_path)
                    
                    del conversion_results[conversion_id]
                
                del conversion_status[conversion_id]
                cleanup_count += 1
        
        return jsonify({
            'success': True,
            'cleaned_up': cleanup_count,
            'message': f'Cleaned up {cleanup_count} old conversions'
        })
    
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

# Auto-cleanup thread
def auto_cleanup():
    while True:
        time.sleep(3600)  # Run every hour
        try:
            with app.app_context():
                cleanup_files()
        except:
            pass

cleanup_thread = threading.Thread(target=auto_cleanup)
cleanup_thread.daemon = True
cleanup_thread.start()

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)