#!/usr/bin/env python3
"""
Final Whisper - AI-powered transcription with smart subtitle formatting
By Final Film
"""

import sys
import os
import subprocess

# Helper function to run subprocess without visible console window
def _run_hidden(cmd, **kwargs):
    """Run subprocess without showing console window on Windows"""
    startupinfo = None
    creationflags = 0
    if sys.platform == 'win32':
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        creationflags = subprocess.CREATE_NO_WINDOW
    return subprocess.run(cmd, startupinfo=startupinfo, creationflags=creationflags, **kwargs)

# Store found paths for debugging
_found_site_packages = []

# When running as frozen EXE, add system site-packages to sys.path FIRST
# This allows the EXE to import whisper/torch from system installation
if getattr(sys, 'frozen', False):
    def _add_system_packages():
        """Add system Python site-packages and stdlib to path for frozen EXE"""
        global _found_site_packages
        
        added_paths = []
        
        # Method 1: Ask Python directly where everything is
        for python_cmd in ['py', 'python', 'python3']:
            try:
                # Get multiple paths in one call for efficiency
                result = _run_hidden(
                    [python_cmd, '-c', '''
import site
import sys
import os

# Print site-packages
for p in site.getsitepackages():
    print(p)
    
# Print user site-packages
print(site.getusersitepackages())

# Print stdlib path (for modules like timeit, etc)
print(os.path.dirname(os.__file__))
'''],
                    capture_output=True, text=True, timeout=10
                )
                if result.returncode == 0 and result.stdout.strip():
                    for path in result.stdout.strip().split('\n'):
                        path = path.strip()
                        if path and os.path.exists(path) and path not in sys.path:
                            sys.path.insert(0, path)
                            added_paths.append(path)
                
                if added_paths:
                    break  # Found paths, stop trying other python commands
            except Exception:
                continue
        
        # Method 2: Try common locations as fallback
        possible_paths = []
        
        for ver in ['312', '311', '310', '39']:
            possible_paths.extend([
                # Site-packages
                os.path.expanduser(f'~\\AppData\\Local\\Programs\\Python\\Python{ver}\\Lib\\site-packages'),
                os.path.expanduser(f'~\\AppData\\Roaming\\Python\\Python{ver}\\site-packages'),
                f'C:\\Python{ver}\\Lib\\site-packages',
                f'C:\\Program Files\\Python{ver}\\Lib\\site-packages',
                # Standard library (for timeit, etc.)
                os.path.expanduser(f'~\\AppData\\Local\\Programs\\Python\\Python{ver}\\Lib'),
                f'C:\\Python{ver}\\Lib',
                f'C:\\Program Files\\Python{ver}\\Lib',
            ])
        
        for path in possible_paths:
            if os.path.exists(path) and path not in sys.path:
                sys.path.insert(0, path)
                added_paths.append(path)
        
        _found_site_packages = added_paths
    
    _add_system_packages()


# Version is read from version.txt (for auto-increment via GitHub Actions)
def _get_version():
    try:
        # Check for version.txt - handle both script and bundled EXE
        if getattr(sys, 'frozen', False):
            # Running as compiled EXE - check PyInstaller's temp folder first
            if hasattr(sys, '_MEIPASS'):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(sys.executable)
        else:
            # Running as script
            base_path = os.path.dirname(os.path.abspath(__file__))

        version_file = os.path.join(base_path, 'version.txt')
        if os.path.exists(version_file):
            with open(version_file, 'r') as f:
                return f.read().strip()
    except:
        pass
    return "1.07"  # Fallback version

VERSION = _get_version()

GITHUB_REPO = "Gerlif/final-whisper"
UPDATE_CHECK_URL = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/version.txt"
RELEASES_URL = f"https://github.com/{GITHUB_REPO}/releases/latest"

import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import subprocess
import threading
from pathlib import Path
import re
import ctypes


def get_python_executable():
    """Get the Python executable path - handles both script and frozen EXE"""
    if getattr(sys, 'frozen', False):
        # Running as frozen EXE - need to find system Python
        import shutil

        # Try common Python commands
        for cmd in ['python', 'python3', 'py']:
            python_path = shutil.which(cmd)
            if python_path:
                # Verify it's actually Python
                try:
                    result = subprocess.run(
                        [python_path, '--version'],
                        capture_output=True,
                        text=True,
                        timeout=5
                    )
                    if result.returncode == 0 and 'Python' in result.stdout:
                        return python_path
                except:
                    continue

        # If we can't find Python, return None and we'll error later
        return None
    else:
        # Running as script - use current Python
        return sys.executable


def is_admin():
    """Check if running with administrator privileges"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def run_as_admin(install_gpu=False):
    """Restart the application with administrator privileges"""
    try:
        if getattr(sys, 'frozen', False):
            # Running as EXE
            script = sys.executable
            params = " --install-gpu" if install_gpu else ""
        else:
            # Running as script
            script = sys.executable
            params = f'"{os.path.abspath(__file__)}"'
            if install_gpu:
                params += " --install-gpu"

        print(f"Requesting elevation: {script} {params}")

        # Request elevation - returns HINSTANCE handle
        # Return value > 32 = success, <= 32 = error code
        result = ctypes.windll.shell32.ShellExecuteW(
            None,       # hwnd
            "runas",    # lpOperation
            script,     # lpFile
            params,     # lpParameters
            None,       # lpDirectory
            1           # nShowCmd (SW_SHOWNORMAL)
        )

        # Check if elevation was successful
        if result > 32:
            print(f"Elevation successful, handle: {result}")
            return True
        else:
            error_codes = {
                0: "Out of memory or resources",
                2: "File not found",
                3: "Path not found",
                5: "Access denied",
                8: "Out of memory",
                26: "Sharing violation",
                27: "File association incomplete or invalid",
                28: "DDE timeout",
                29: "DDE transaction failed",
                30: "DDE busy",
                31: "No file association",
                32: "DDE failed"
            }
            error_msg = error_codes.get(result, f"Unknown error code: {result}")
            print(f"Elevation failed: {error_msg}")
            return False

    except Exception as e:
        print(f"Exception during elevation: {e}")
        import traceback
        traceback.print_exc()
        return False


class CollapsibleFrame(ttk.Frame):
    """A frame that can be collapsed/expanded by clicking the header."""
    
    def __init__(self, parent, text="", collapsed=False, **kwargs):
        super().__init__(parent, **kwargs)
        
        self._collapsed = collapsed
        self._text = text
        self._status = ""  # Optional status indicator
        
        # Header frame with toggle button
        self.header = ttk.Frame(self)
        self.header.pack(fill=tk.X)
        
        # Toggle button (arrow + text)
        self.toggle_btn = ttk.Button(
            self.header, 
            text=self._get_header_text(),
            command=self.toggle,
            style="Collapsible.TButton"
        )
        self.toggle_btn.pack(fill=tk.X)
        
        # Content frame (the actual content goes here)
        self.content = ttk.Frame(self, padding="12")
        if not collapsed:
            self.content.pack(fill=tk.BOTH, expand=True)
    
    def _get_header_text(self):
        arrow = "▼" if not self._collapsed else "▶"
        status = f"  [{self._status}]" if self._status else ""
        return f"  {arrow}  {self._text}{status}"
    
    def set_status(self, status):
        """Set a status indicator that shows in the header"""
        self._status = status
        self.toggle_btn.config(text=self._get_header_text())
    
    def set_title(self, text):
        """Update the title text"""
        self._text = text
        self.toggle_btn.config(text=self._get_header_text())
    
    def toggle(self):
        """Toggle collapsed state"""
        self._collapsed = not self._collapsed
        self.toggle_btn.config(text=self._get_header_text())
        
        if self._collapsed:
            self.content.pack_forget()
        else:
            self.content.pack(fill=tk.BOTH, expand=True)
    
    def collapse(self):
        """Collapse the frame"""
        if not self._collapsed:
            self.toggle()
    
    def expand(self):
        """Expand the frame"""
        if self._collapsed:
            self.toggle()


class GradientProgressBar(tk.Canvas):
    """A custom progress bar with gradient colors and animation."""
    
    def __init__(self, parent, height=27, **kwargs):
        super().__init__(parent, height=height, 
                        highlightthickness=0, bg='#1e1e1e', **kwargs)
        self._value = 0
        self._maximum = 100
        self._height = height
        self._radius = 4  # Less rounded corners
        
        # Purple gradient colors - mirrored for seamless loop
        base = ['#8B5CF6', '#A855F7', '#D946EF', '#EC4899']
        # Create mirrored palette: A B C D D C B A for perfect seamless loop
        self._colors = base + base[::-1]
        
        # Animation state
        self._animating = False
        self._color_offset = 0
        self._animation_speed = 40  # ms between frames (faster)
        
        # Bind to resize events to update width
        self.bind('<Configure>', self._on_resize)
        
        # Draw initial state
        self._draw()
    
    def _on_resize(self, event):
        """Handle resize events."""
        self._draw()
    
    def start_animation(self):
        """Start the gradient cycling animation."""
        if not self._animating:
            self._animating = True
            self._animate()
    
    def stop_animation(self):
        """Stop the gradient cycling animation."""
        self._animating = False
        self._color_offset = 0
        self._draw()
    
    def _animate(self):
        """Animate the gradient by cycling colors."""
        if not self._animating:
            return
        
        # Decrement offset for rightward movement (faster)
        self._color_offset = (self._color_offset - 0.025) % 1.0
        self._draw()
        
        # Schedule next frame
        self.after(self._animation_speed, self._animate)
    
    def _interpolate_color(self, color1, color2, factor):
        """Interpolate between two hex colors."""
        r1, g1, b1 = int(color1[1:3], 16), int(color1[3:5], 16), int(color1[5:7], 16)
        r2, g2, b2 = int(color2[1:3], 16), int(color2[3:5], 16), int(color2[5:7], 16)
        r = int(r1 + (r2 - r1) * factor)
        g = int(g1 + (g2 - g1) * factor)
        b = int(b1 + (b2 - b1) * factor)
        return f'#{r:02x}{g:02x}{b:02x}'
    
    def _get_gradient_color(self, x, width):
        """Get the gradient color at position x with animation offset."""
        if width <= 0:
            return self._colors[0]
        
        num_colors = len(self._colors)
        
        # Calculate position in the color cycle (0.0 to 1.0) with animation offset
        pos = (x / width + self._color_offset) % 1.0
        
        # Map to color array - use num_colors for full wrap
        scaled_pos = pos * num_colors
        segment = int(scaled_pos) % num_colors
        factor = scaled_pos - int(scaled_pos)
        
        next_segment = (segment + 1) % num_colors
        
        return self._interpolate_color(self._colors[segment], self._colors[next_segment], factor)
    
    def _draw(self):
        """Draw the progress bar."""
        self.delete("all")
        
        # Get current width from canvas
        width = self.winfo_width()
        if width <= 1:
            width = 400  # Default width before first render
        
        height = self._height
        radius = self._radius
        
        # Calculate fill width
        fill_width = int((self._value / self._maximum) * width) if self._maximum > 0 else 0
        
        # Draw background (rounded rectangle)
        self._round_rectangle(0, 0, width, height, radius, fill='#3c3c3c', outline='')
        
        # Draw gradient fill
        if fill_width > radius * 2:
            # Draw gradient using vertical lines
            for x in range(fill_width):
                color = self._get_gradient_color(x, fill_width)
                
                # Simple clipping for rounded corners
                if x < radius:
                    # Left rounded part - approximate with smaller lines
                    import math
                    y_inset = radius - int(math.sqrt(radius**2 - (radius - x)**2))
                    self.create_line(x, y_inset, x, height - y_inset, fill=color, width=1)
                elif x > fill_width - radius:
                    # Right rounded part
                    import math
                    dist = fill_width - x
                    y_inset = radius - int(math.sqrt(max(0, radius**2 - (radius - dist)**2)))
                    self.create_line(x, y_inset, x, height - y_inset, fill=color, width=1)
                else:
                    # Middle part - full height
                    self.create_line(x, 0, x, height, fill=color, width=1)
        elif fill_width > 0:
            # Very small progress - just draw a small rounded rect
            self._round_rectangle(0, 0, fill_width, height, min(radius, fill_width//2), 
                                 fill=self._colors[0], outline='')
    
    def _round_rectangle(self, x1, y1, x2, y2, radius, **kwargs):
        """Draw a rounded rectangle."""
        points = [
            x1 + radius, y1,
            x2 - radius, y1,
            x2, y1,
            x2, y1 + radius,
            x2, y2 - radius,
            x2, y2,
            x2 - radius, y2,
            x1 + radius, y2,
            x1, y2,
            x1, y2 - radius,
            x1, y1 + radius,
            x1, y1,
        ]
        return self.create_polygon(points, smooth=True, **kwargs)
    
    def set_value(self, value):
        """Set the progress value (0-100)."""
        new_value = max(0, min(self._maximum, value))
        # Only update if value changed by at least 1%
        if abs(new_value - self._value) >= 1 or new_value == 0 or new_value >= self._maximum:
            self._value = new_value
            # Only redraw if not animating (animation will redraw anyway)
            if not self._animating:
                self._draw()
    
    def get_value(self):
        """Get the current value."""
        return self._value
    
    def configure(self, **kwargs):
        """Configure the progress bar."""
        if 'value' in kwargs:
            self.set_value(kwargs.pop('value'))
        super().configure(**kwargs)
    
    def config(self, **kwargs):
        """Alias for configure."""
        self.configure(**kwargs)
    
    def __setitem__(self, key, value):
        """Allow setting value like progress['value'] = 50."""
        if key == 'value':
            self.set_value(value)
        else:
            super().__setitem__(key, value)
    
    def __getitem__(self, key):
        """Allow getting value like progress['value']."""
        if key == 'value':
            return self._value
        return super().__getitem__(key)


def split_balanced_lines(text, max_chars):
    """
    Split text into balanced lines with max character limit.
    Top line slightly longer if needed. Never splits words.
    Handles very long text by splitting into multiple lines if needed.
    """
    words = text.split()
    
    if len(text) <= max_chars:
        return [text]
    
    # Try to split into two lines first
    best_split = 0
    best_score = float('inf')
    
    for i in range(1, len(words)):
        line1 = ' '.join(words[:i])
        line2 = ' '.join(words[i:])
        
        # Both lines must be within limit
        if len(line1) > max_chars or len(line2) > max_chars:
            continue
        
        # Score: prefer balanced lines, with slight preference for top line longer
        diff = abs(len(line1) - len(line2))
        # Add small penalty if bottom line is longer (we prefer top line longer)
        if len(line2) > len(line1):
            score = diff + 0.5
        else:
            score = diff
        
        if score < best_score:
            best_score = score
            best_split = i
    
    # If we found a valid 2-line split, use it
    if best_split > 0:
        line1 = ' '.join(words[:best_split])
        line2 = ' '.join(words[best_split:])
        return [line1, line2]
    
    # Text is too long for 2 lines - split into chunks that fit
    lines = []
    current_line = []
    current_length = 0
    
    for word in words:
        word_length = len(word)
        # +1 for space if not first word in line
        space = 1 if current_line else 0
        
        if current_length + space + word_length <= max_chars:
            current_line.append(word)
            current_length += space + word_length
        else:
            # Save current line and start new one
            if current_line:
                lines.append(' '.join(current_line))
            current_line = [word]
            current_length = word_length
    
    # Don't forget the last line
    if current_line:
        lines.append(' '.join(current_line))
    
    return lines if lines else [text]


# Words that shouldn't end a subtitle (conjunctions, articles, prepositions, auxiliaries)
ORPHAN_WORDS = {
    # Danish - conjunctions, prepositions, articles, pronouns, auxiliaries
    'og', 'eller', 'men', 'for', 'så', 'at', 'som', 'der', 'den', 'det', 
    'de', 'en', 'et', 'i', 'på', 'til', 'med', 'af', 'om', 'fra', 'ved',
    'har', 'er', 'var', 'vil', 'kan', 'skal', 'må', 'hvis', 'når', 'hvor',
    'denne', 'dette', 'disse', 'min', 'din', 'sin', 'vores', 'jeres', 'deres',
    'jeg', 'du', 'han', 'hun', 'vi', 'dem', 'sig', 'os', 'mig', 'dig',
    'ikke', 'også', 'bare', 'kun', 'jo', 'da', 'nu', 'her', 'der',
    'bliver', 'blive', 'blev', 'været', 'være', 'havde', 'have',
    'ville', 'kunne', 'skulle', 'måtte', 'burde',
    'meget', 'mere', 'mest', 'nogle', 'noget', 'nogen', 'ingen', 'alle',
    'efter', 'før', 'under', 'over', 'mellem', 'igennem', 'gennem',
    'fordi', 'siden', 'mens', 'indtil', 'uden', 'inden',
    # English - conjunctions, prepositions, articles, pronouns, auxiliaries
    'the', 'a', 'an', 'and', 'or', 'but', 'if', 'to', 'of', 'in', 'on',
    'is', 'are', 'was', 'were', 'be', 'been', 'have', 'has', 'had', 'for',
    'with', 'that', 'this', 'these', 'those', 'my', 'your', 'his', 'her',
    'i', 'you', 'he', 'she', 'we', 'they', 'it', 'me', 'him', 'us', 'them',
    'not', 'also', 'just', 'only', 'so', 'then', 'now', 'here', 'there',
    'will', 'would', 'could', 'should', 'can', 'may', 'might', 'must',
    'very', 'more', 'most', 'some', 'any', 'no', 'all', 'each', 'every',
    'by', 'at', 'from', 'into', 'through', 'during', 'before', 'after',
    'above', 'below', 'between', 'under', 'over', 'about', 'against'
}

# Minimum characters before we allow a sentence-break split
MIN_CHARS_BEFORE_SPLIT = 50

# Minimum characters for a standalone subtitle (otherwise merge with neighbor)
MIN_SUBTITLE_CHARS = 20


def format_timestamp(seconds):
    """Convert seconds to SRT timestamp format (HH:MM:SS,mmm)"""
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    millis = int((seconds % 1) * 1000)
    return f"{hours:02d}:{minutes:02d}:{secs:02d},{millis:03d}"


def is_sentence_end(word_text):
    """Check if a word ends a sentence"""
    return word_text.rstrip().endswith(('.', '?', '!'))


def is_orphan_word(word_text):
    """Check if a word shouldn't end a subtitle"""
    clean = word_text.lower().strip().rstrip('.,!?:;')
    return clean in ORPHAN_WORDS


def generate_smart_srt(result, output_file, max_chars_per_line=40, max_lines=2):
    """
    Generate SRT file from Whisper result using word-level timestamps.
    Splits subtitles at sentence boundaries and avoids orphaned words.
    
    Args:
        result: Whisper transcription result with word_timestamps=True
        output_file: Path to output SRT file
        max_chars_per_line: Maximum characters per line
        max_lines: Maximum lines per subtitle (typically 2)
    """
    max_subtitle_chars = max_chars_per_line * max_lines
    subtitles = []
    
    for segment in result.get('segments', []):
        words = segment.get('words', [])
        
        if not words:
            # Fallback if no word-level timestamps
            text = segment['text'].strip()
            if text:  # Only add non-empty subtitles
                subtitles.append({
                    'start': segment['start'],
                    'end': segment['end'],
                    'text': text
                })
            continue
        
        # Group words into subtitle chunks
        current_words = []
        current_chars = 0
        
        for i, word_info in enumerate(words):
            word = word_info.get('word', '').strip()
            if not word:
                continue
            
            word_len = len(word) + (1 if current_words else 0)  # +1 for space
            
            # Check if we should start a new subtitle
            should_split = False
            
            # Rule 1: Character limit exceeded - must split
            if current_chars + word_len > max_subtitle_chars and current_words:
                should_split = True
            
            # Rule 2: Previous word ended a sentence - only split if we have enough content
            elif current_words and current_chars >= MIN_CHARS_BEFORE_SPLIT:
                prev_word = current_words[-1]['word']
                if is_sentence_end(prev_word):
                    should_split = True
            
            if should_split:
                # Check for orphan word at end - if so, move it to next subtitle
                while current_words and is_orphan_word(current_words[-1]['word']) and len(current_words) > 1:
                    # Move orphan words to be added to next subtitle
                    orphan = current_words.pop()
                    # Prepend to current processing (will be added after split)
                    word_info = {'word': orphan['word'] + ' ' + word_info.get('word', ''), 
                                 'start': orphan['start'], 
                                 'end': word_info.get('end', orphan['end'])}
                    word = word_info['word'].strip()
                    word_len = len(word)
                
                if current_words:
                    # Save current subtitle - normalize whitespace
                    text = ' '.join(w['word'].strip() for w in current_words)
                    text = ' '.join(text.split())  # Normalize any double spaces
                    if text.strip():
                        subtitles.append({
                            'start': current_words[0]['start'],
                            'end': current_words[-1]['end'],
                            'text': text
                        })
                    current_words = []
                    current_chars = 0
            
            # Add word to current subtitle (strip whitespace)
            word_info_clean = word_info.copy()
            word_info_clean['word'] = word_info.get('word', '').strip()
            current_words.append(word_info_clean)
            current_chars += len(word_info_clean['word']) + (1 if len(current_words) > 1 else 0)
        
        # Don't forget the last subtitle in the segment
        if current_words:
            text = ' '.join(w['word'].strip() for w in current_words)
            text = ' '.join(text.split())  # Normalize any double spaces
            if text.strip():
                subtitles.append({
                    'start': current_words[0]['start'],
                    'end': current_words[-1]['end'],
                    'text': text
                })
    
    # Merge very short subtitles with neighbors
    subtitles = merge_short_subtitles(subtitles, MIN_SUBTITLE_CHARS, max_subtitle_chars)
    
    # Remove empty subtitles
    subtitles = [s for s in subtitles if s['text'].strip()]
    
    # Enforce minimum subtitle duration (1.5 seconds)
    subtitles = enforce_minimum_duration(subtitles, min_duration=1.5)
    
    # Write SRT file
    with open(output_file, 'w', encoding='utf-8') as f:
        for i, sub in enumerate(subtitles, 1):
            # Skip empty subtitles
            if not sub['text'].strip():
                continue
                
            # Format lines within subtitle
            lines = split_balanced_lines(sub['text'], max_chars_per_line)
            
            # Fix orphaned words within lines
            if len(lines) >= 2:
                lines = fix_line_orphans(lines, max_chars_per_line)
            
            # Limit to max lines
            if len(lines) > max_lines:
                lines = lines[:max_lines]
            
            # Write subtitle entry
            start_ts = format_timestamp(sub['start'])
            end_ts = format_timestamp(sub['end'])
            
            f.write(f"{i}\n")
            f.write(f"{start_ts} --> {end_ts}\n")
            f.write('\n'.join(lines))
            f.write('\n\n')


def merge_short_subtitles(subtitles, min_chars, max_chars):
    """
    Merge subtitles that are too short with their neighbors.
    Uses word count as primary metric - single words should always be merged.
    Respects complete sentences (3+ words ending with punctuation).
    """
    if len(subtitles) <= 1:
        return subtitles
    
    def word_count(text):
        return len(text.split())
    
    def is_complete_sentence(text):
        """Check if text is a complete sentence (3+ words, ends with punctuation)"""
        text = text.strip()
        words = word_count(text)
        ends_with_punct = text.endswith(('.', '?', '!'))
        return words >= 3 and ends_with_punct
    
    def should_merge(text):
        """Determine if a subtitle should be merged"""
        words = word_count(text)
        # Always merge 1-2 word subtitles, unless it's a complete sentence
        if words <= 2:
            return not is_complete_sentence(text)
        return False
    
    merged = []
    i = 0
    
    while i < len(subtitles):
        current = subtitles[i]
        
        # Check if current subtitle should be merged
        if should_merge(current['text']):
            merged_successfully = False
            
            # Try to merge with PREVIOUS subtitle first (usually works better)
            if merged:
                prev = merged[-1]
                combined_text = prev['text'] + ' ' + current['text']
                
                if len(combined_text) <= max_chars:
                    # Merge with previous
                    merged[-1] = {
                        'start': prev['start'],
                        'end': current['end'],
                        'text': combined_text
                    }
                    i += 1
                    merged_successfully = True
            
            # If couldn't merge with previous, try next
            if not merged_successfully and i + 1 < len(subtitles):
                next_sub = subtitles[i + 1]
                combined_text = current['text'] + ' ' + next_sub['text']
                
                if len(combined_text) <= max_chars:
                    # Merge with next
                    merged.append({
                        'start': current['start'],
                        'end': next_sub['end'],
                        'text': combined_text
                    })
                    i += 2  # Skip next since we merged it
                    merged_successfully = True
            
            if merged_successfully:
                continue
        
        # No merge needed or possible
        merged.append(current)
        i += 1
    
    # Second pass: catch any remaining 1-2 word subtitles
    final = []
    for sub in merged:
        if should_merge(sub['text']) and final:
            # Try to merge with previous
            prev = final[-1]
            combined_text = prev['text'] + ' ' + sub['text']
            if len(combined_text) <= max_chars:
                final[-1] = {
                    'start': prev['start'],
                    'end': sub['end'],
                    'text': combined_text
                }
                continue
        final.append(sub)
    
    return final


def enforce_minimum_duration(subtitles, min_duration=1.5, extend_duration=1.5):
    """
    Extend subtitle end times to ensure minimum display duration AND
    add extra breathing room after each subtitle when space allows.
    
    Args:
        subtitles: List of subtitle dicts with 'start', 'end', 'text'
        min_duration: Minimum duration in seconds (default 1.5s)
        extend_duration: Extra time to add after subtitle ends (default 1.5s)
    
    Returns:
        List of subtitles with adjusted end times
    """
    if not subtitles:
        return subtitles
    
    adjusted = []
    for i, sub in enumerate(subtitles):
        current_end = sub['end']
        
        # Calculate the ideal new end time:
        # Either min_duration from start, or extend_duration after current end, whichever is later
        min_end = sub['start'] + min_duration
        extended_end = current_end + extend_duration
        ideal_end = max(min_end, extended_end)
        
        # Check if there's a next subtitle
        if i + 1 < len(subtitles):
            next_start = subtitles[i + 1]['start']
            # Don't overlap with next subtitle - leave small gap (0.05s)
            max_end = next_start - 0.05
            new_end = min(ideal_end, max_end)
        else:
            # Last subtitle - can extend freely
            new_end = ideal_end
        
        # Only extend if we actually gain duration
        if new_end > current_end:
            adjusted.append({
                'start': sub['start'],
                'end': new_end,
                'text': sub['text']
            })
        else:
            adjusted.append(sub)
    
    return adjusted


def fix_line_orphans(lines, max_chars):
    """
    Fix orphaned words at line ends within a subtitle.
    """
    if len(lines) < 2:
        return lines
    
    fixed = list(lines)
    
    for i in range(len(fixed) - 1):
        words = fixed[i].split()
        if not words:
            continue
        
        last_word = words[-1]
        if is_orphan_word(last_word):
            new_current = ' '.join(words[:-1])
            new_next = last_word + ' ' + fixed[i + 1]
            
            if new_current and len(new_next) <= max_chars:
                fixed[i] = new_current
                fixed[i + 1] = new_next
    
    return fixed


class WhisperGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Final Whisper v{VERSION}")
        self.root.geometry("1300x900")
        
        # Set window icon
        self.set_window_icon()
        
        # Apply dark theme
        self.setup_dark_theme()
        
        # Variables
        self.video_path = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(Path.home() / "whisper_output"))
        self.language = tk.StringVar(value="da: Danish")
        self.model = tk.StringVar(value="large")
        self.max_line_count = tk.IntVar(value=2)
        self.use_word_timestamps = tk.BooleanVar(value=True)
        self.max_chars_per_line = tk.IntVar(value=40)
        self.context_prompt = tk.StringVar(value="")
        self.use_gpu = tk.BooleanVar(value=True)
        self.device_info = tk.StringVar(value="Checking...")
        self.processing = False
        self.stop_requested = False
        self.transcription_start_time = None
        self.timer_running = False
        self.last_progress_value = None
        self.last_audio_position = None
        self.last_audio_duration = None
        self.transcription_process = None  # For subprocess transcription
        self.selected_files = None  # For batch processing multiple files
        
        # Anti-hallucination settings
        self.condition_on_previous_text = tk.BooleanVar(value=True)  # Default True (Whisper default)
        self.no_speech_threshold = tk.DoubleVar(value=0.6)  # Default 0.6 (Whisper default)
        self.hallucination_silence_threshold = tk.DoubleVar(value=0.0)  # Default 0 (disabled)
        
        # AI Proofreading settings
        self.use_ai_proofreading = tk.BooleanVar(value=False)
        self.anthropic_api_key = tk.StringVar(value="")
        
        # Load saved API key if exists
        self.load_api_key()
        
        # Load saved settings (model, language, checkboxes, etc.)
        self.load_settings()
        
        self.create_widgets()
        self.check_whisper_installation()
        # Note: check_gpu_availability() is called by check_whisper_installation() 
        # after confirming Whisper is installed (or after Whisper is installed)

        # Check for updates after a delay (let initialization complete first)
        self.root.after(3000, self.check_for_updates)

        # Log startup diagnostics
        self.log_startup_diagnostics()

        # Check if we were launched with --install-gpu flag (after admin elevation)
        if '--install-gpu' in sys.argv:
            if is_admin():
                self.log("🔧 Auto-starting GPU installation (elevated privileges)")
                # Give the UI time to load, then start installation
                self.root.after(1000, self.install_cuda_pytorch_direct)
            else:
                self.log("⚠️ WARNING: --install-gpu flag detected but NOT running as admin!")
                self.log("This suggests the elevation failed. Installation will likely fail.")
                messagebox.showwarning("Not Administrator",
                    "The app was supposed to restart with admin rights, but elevation failed.\n\n"
                    "GPU installation requires administrator privileges.\n\n"
                    "Please right-click Final Whisper.exe and select 'Run as administrator'.")
    
    def log_startup_diagnostics(self):
        """Log diagnostic information at startup"""
        self.log(f"Final Whisper v{VERSION}")
        
        # Store debug info for error reporting but don't show by default
        self._debug_info = {
            'admin': is_admin(),
            'python': sys.version.split()[0],
            'executable': sys.executable,
            'site_packages': _found_site_packages if getattr(sys, 'frozen', False) else []
        }
        
        if '--install-gpu' in sys.argv:
            self.log(f"Running as Administrator")
        self.log("")  # Blank line for readability

    def set_window_icon(self):
        """Set the window icon (both title bar and taskbar)"""
        try:
            import os
            import sys
            
            # Possible icon locations
            icon_paths = [
                'icon.ico',  # Same directory
                os.path.join(os.path.dirname(__file__), 'icon.ico'),  # Script directory
                os.path.join(os.path.dirname(sys.executable), 'icon.ico'),  # Exe directory
            ]
            
            # If running as PyInstaller bundle
            if hasattr(sys, '_MEIPASS'):
                icon_paths.insert(0, os.path.join(sys._MEIPASS, 'icon.ico'))
            
            icon_path = None
            for path in icon_paths:
                if os.path.exists(path):
                    icon_path = path
                    break
            
            if icon_path:
                # Set title bar icon (this works reliably)
                self.root.iconbitmap(default=icon_path)
                
                # For taskbar, try to use iconphoto with PIL
                try:
                    from PIL import Image, ImageTk
                    # Load ICO and get the largest size
                    ico = Image.open(icon_path)
                    # Convert to PhotoImage
                    icon_photo = ImageTk.PhotoImage(ico)
                    self.root.iconphoto(True, icon_photo)
                    self._icon_photo = icon_photo  # Keep reference to prevent garbage collection
                except ImportError:
                    # PIL not available, try with PNG if it exists
                    try:
                        png_path = icon_path.replace('.ico', '.png')
                        if os.path.exists(png_path):
                            icon_image = tk.PhotoImage(file=png_path)
                            self.root.iconphoto(True, icon_image)
                            self._icon_image = icon_image
                    except Exception:
                        pass
                except Exception:
                    pass
                    
        except Exception:
            pass  # Icon is optional, continue without it
    
    def setup_dark_theme(self):
        """Setup a modern dark theme"""
        # Purple accent color
        accent_color = '#8B5CF6'
        accent_hover = '#A855F7'
        
        # Try to use sv_ttk (Sun Valley theme) if available
        try:
            import sv_ttk
            sv_ttk.set_theme("dark")
            # Customize sv_ttk with purple accent
            style = ttk.Style()
            style.configure('Collapsible.TButton',
                padding=(8, 6),
                font=('Segoe UI', 9, 'bold'),
                anchor='w'
            )
            # Override accent colors for buttons and checkboxes
            style.configure('Accent.TButton', background=accent_color)
            style.map('Accent.TButton', 
                background=[('active', accent_hover), ('pressed', '#7C3AED')])
            # Try to set accent color for checkboxes
            style.configure('TCheckbutton', indicatorcolor=accent_color)
            style.map('TCheckbutton',
                indicatorcolor=[('selected', accent_color), ('selected active', accent_hover)])
            return
        except ImportError:
            # Try to install sv_ttk automatically
            try:
                import subprocess
                import sys
                subprocess.check_call([sys.executable, "-m", "pip", "install", "sv-ttk", "--quiet"],
                                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                import sv_ttk
                sv_ttk.set_theme("dark")
                # Customize sv_ttk with purple accent
                style = ttk.Style()
                style.configure('Collapsible.TButton',
                    padding=(8, 6),
                    font=('Segoe UI', 9, 'bold'),
                    anchor='w'
                )
                # Override accent colors
                style.configure('Accent.TButton', background=accent_color)
                style.map('Accent.TButton', 
                    background=[('active', accent_hover), ('pressed', '#7C3AED')])
                style.configure('TCheckbutton', indicatorcolor=accent_color)
                style.map('TCheckbutton',
                    indicatorcolor=[('selected', accent_color), ('selected active', accent_hover)])
                return
            except:
                pass  # Fall back to custom theme
        
        # Fallback: Custom dark theme
        self.root.configure(bg='#1e1e1e')
        
        style = ttk.Style()
        
        # Try to use clam as base (more customizable)
        try:
            style.theme_use('clam')
        except:
            pass
        
        # Define colors
        colors = {
            'bg': '#1e1e1e',
            'fg': '#ffffff',
            'select_bg': '#8B5CF6',
            'select_fg': '#ffffff',
            'frame_bg': '#252526',
            'entry_bg': '#3c3c3c',
            'button_bg': '#8B5CF6',
            'button_hover': '#A855F7',
            'border': '#3c3c3c',
            'label_frame_bg': '#252526',
            'disabled_fg': '#6d6d6d',
            'accent': '#8B5CF6',
            'accent_hover': '#A855F7',
            'accent_pressed': '#7C3AED',
            'success': '#4ec9b0',
            'warning': '#dcdcaa',
            'error': '#f14c4c'
        }
        
        # Configure styles
        style.configure('.',
            background=colors['bg'],
            foreground=colors['fg'],
            fieldbackground=colors['entry_bg'],
            borderwidth=0,
            focusthickness=0
        )
        
        style.configure('TFrame', background=colors['bg'])
        style.configure('TLabel', background=colors['bg'], foreground=colors['fg'], font=('Segoe UI', 9))
        style.configure('TLabelframe', background=colors['frame_bg'], foreground=colors['fg'])
        style.configure('TLabelframe.Label', background=colors['frame_bg'], foreground=colors['fg'], font=('Segoe UI', 9, 'bold'))
        
        style.configure('TButton',
            background=colors['button_bg'],
            foreground=colors['fg'],
            padding=(12, 6),
            font=('Segoe UI', 9)
        )
        style.map('TButton',
            background=[('active', colors['button_hover']), ('pressed', colors['accent'])],
            foreground=[('disabled', colors['disabled_fg'])]
        )
        
        style.configure('Accent.TButton',
            background=colors['accent'],
            foreground='#ffffff',
            padding=(16, 8),
            font=('Segoe UI', 10, 'bold')
        )
        style.map('Accent.TButton',
            background=[('active', colors['accent_hover']), ('pressed', colors['accent_pressed'])]
        )
        
        style.configure('TEntry',
            fieldbackground=colors['entry_bg'],
            foreground=colors['fg'],
            insertcolor=colors['fg'],
            padding=6
        )
        
        style.configure('TCombobox',
            fieldbackground=colors['entry_bg'],
            background=colors['entry_bg'],
            foreground=colors['fg'],
            arrowcolor=colors['fg'],
            padding=4
        )
        style.map('TCombobox',
            fieldbackground=[('readonly', colors['entry_bg'])],
            selectbackground=[('readonly', colors['select_bg'])]
        )
        
        style.configure('TCheckbutton',
            background=colors['bg'],
            foreground=colors['fg'],
            font=('Segoe UI', 9)
        )
        style.map('TCheckbutton',
            background=[('active', colors['bg'])]
        )
        
        style.configure('TSpinbox',
            fieldbackground=colors['entry_bg'],
            background=colors['entry_bg'],
            foreground=colors['fg'],
            arrowcolor=colors['fg'],
            padding=4
        )
        
        style.configure('Horizontal.TProgressbar',
            background=colors['accent'],
            troughcolor=colors['entry_bg'],
            borderwidth=0,
            thickness=16
        )
        
        # Collapsible section button style
        style.configure('Collapsible.TButton',
            background=colors['frame_bg'],
            foreground=colors['fg'],
            padding=(8, 6),
            font=('Segoe UI', 9, 'bold'),
            anchor='w'
        )
        style.map('Collapsible.TButton',
            background=[('active', colors['entry_bg']), ('pressed', colors['entry_bg'])]
        )
        
        # Store colors for use in other widgets
        self.colors = colors
        
    def create_widgets(self):
        # Main container
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Header with logo and title
        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # Load and display logo (left side, smaller, vertically centered)
        self.logo_image = None
        try:
            import os
            import sys
            
            # Possible logo locations
            logo_paths = [
                'logo.png',
                os.path.join(os.path.dirname(__file__), 'logo.png'),
                os.path.join(os.path.dirname(sys.executable), 'logo.png'),
            ]
            if hasattr(sys, '_MEIPASS'):
                logo_paths.insert(0, os.path.join(sys._MEIPASS, 'logo.png'))
            
            logo_loaded = False
            for logo_path in logo_paths:
                if os.path.exists(logo_path):
                    try:
                        from PIL import Image, ImageTk
                        img = Image.open(logo_path)
                        # Resize to ~65% of original (was 180x65, now ~117x42)
                        new_width = int(img.width * 0.65)
                        new_height = int(img.height * 0.65)
                        img = img.resize((new_width, new_height), Image.LANCZOS)
                        self.logo_image = ImageTk.PhotoImage(img)
                        logo_label = ttk.Label(header_frame, image=self.logo_image)
                        logo_label.pack(side=tk.LEFT, padx=(0, 15), pady=(5, 5))
                        logo_loaded = True
                        break
                    except ImportError:
                        # PIL not available, try with tk.PhotoImage (PNG only, no resize)
                        try:
                            self.logo_image = tk.PhotoImage(file=logo_path)
                            # Subsample to make it smaller (approximate 65% reduction)
                            self.logo_image = self.logo_image.subsample(2, 2)
                            logo_label = ttk.Label(header_frame, image=self.logo_image)
                            logo_label.pack(side=tk.LEFT, padx=(0, 15), pady=(5, 5))
                            logo_loaded = True
                            break
                        except Exception:
                            pass
                    except Exception:
                        pass
        except Exception:
            pass  # Logo is optional
        
        # Title (right of logo)
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(side=tk.LEFT, fill=tk.Y)

        # Version and update button in same row
        version_row = ttk.Frame(title_frame)
        version_row.pack(anchor=tk.W)

        title_label = ttk.Label(version_row, text=f"Final Whisper v{VERSION}", font=("Segoe UI", 18, "bold"))
        title_label.pack(side=tk.LEFT)

        # Update button (initially hidden)
        self.update_button = ttk.Button(version_row, text="🔄 Update Available - Click to Install",
                                       command=self.download_and_install_update, style="Accent.TButton")
        self.update_button.pack(side=tk.LEFT, padx=10)
        self.update_button.pack_forget()  # Hide initially
        self.new_version = None  # Store the new version when available

        subtitle_label = ttk.Label(title_frame, text="AI-powered transcription with smart formatting",
                                   font=("Segoe UI", 9))
        subtitle_label.pack(anchor=tk.W)
        
        # Content area with two columns
        content_frame = ttk.Frame(main_frame)
        content_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Left column for settings
        self.left_panel = ttk.Frame(content_frame)
        self.left_panel.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 15))
        left_frame = self.left_panel  # Alias for backward compatibility
        
        # Right column for log
        right_frame = ttk.Frame(content_frame)
        right_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # File selection (row 0)
        self.files_section = CollapsibleFrame(left_frame, text="Files")
        self.files_section.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        file_frame = self.files_section.content
        
        ttk.Label(file_frame, text="Input:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(file_frame, textvariable=self.video_path, width=38).grid(row=0, column=1, padx=8)
        
        # Button frame for file/folder selection
        btn_frame = ttk.Frame(file_frame)
        btn_frame.grid(row=0, column=2)
        ttk.Button(btn_frame, text="File(s)", command=self.browse_video).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="Folder", command=self.browse_input_folder).pack(side=tk.LEFT, padx=(4, 0))
        
        ttk.Label(file_frame, text="Output Folder:").grid(row=1, column=0, sticky=tk.W, pady=(8,0))
        ttk.Entry(file_frame, textvariable=self.output_dir, width=38).grid(row=1, column=1, padx=8, pady=(8,0))
        ttk.Button(file_frame, text="Browse", command=self.browse_output).grid(row=1, column=2, sticky=tk.W, pady=(8,0))
        
        # Transcription settings (row 1)
        self.transcription_section = CollapsibleFrame(left_frame, text="Transcription")
        self.transcription_section.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        settings_frame = self.transcription_section.content
        settings_frame.columnconfigure(1, weight=1)  # Make column 1 expand
        
        ttk.Label(settings_frame, text="Language:").grid(row=0, column=0, sticky=tk.W)
        self.lang_combo = ttk.Combobox(settings_frame, textvariable=self.language,
                                 values=self.get_available_languages(),
                                 width=20, state="readonly")
        self.lang_combo.grid(row=0, column=1, sticky=tk.W, padx=8)
        
        ttk.Label(settings_frame, text="Context/Prompt:").grid(row=1, column=0, sticky=(tk.W, tk.N), pady=(8,0))
        
        # Context frame to hold both single and multi-file modes
        self.context_frame = ttk.Frame(settings_frame)
        self.context_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=8, pady=(8,0))
        self.context_frame.columnconfigure(0, weight=1)
        
        # Create context entry with dark theme colors - full width and taller
        self.context_entry = tk.Text(self.context_frame, height=5, wrap=tk.WORD,
                                     bg='#3c3c3c', fg='#ffffff', insertbackground='#ffffff',
                                     relief='flat', font=('Segoe UI', 9), padx=6, pady=6)
        self.context_entry.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # Bind text widget to variable
        def update_context(*args):
            self.context_prompt.set(self.context_entry.get("1.0", tk.END).strip())
            # Also update per-file prompt if in batch mode
            if hasattr(self, '_batch_prompts') and self._batch_prompts and hasattr(self, '_current_prompt_index'):
                self._batch_prompts[self._current_prompt_index] = self.context_entry.get("1.0", tk.END).strip()
        self.context_entry.bind("<KeyRelease>", update_context)
        
        # Batch context controls (hidden by default)
        self.batch_context_frame = ttk.Frame(self.context_frame)
        # Will be shown when multiple files selected
        
        # Navigation frame with arrows and file indicator
        self.batch_nav_frame = ttk.Frame(self.batch_context_frame)
        self.batch_nav_frame.pack(fill=tk.X, pady=(6, 0))
        
        self.prev_file_btn = ttk.Button(self.batch_nav_frame, text="◀", width=3, command=self._prev_batch_file)
        self.prev_file_btn.pack(side=tk.LEFT)
        
        self.batch_file_indicator = ttk.Label(self.batch_nav_frame, text="File 1/3: filename.mp4", font=("Segoe UI", 9))
        self.batch_file_indicator.pack(side=tk.LEFT, padx=8, expand=True)
        
        self.next_file_btn = ttk.Button(self.batch_nav_frame, text="▶", width=3, command=self._next_batch_file)
        self.next_file_btn.pack(side=tk.LEFT)
        
        # Same prompt for all checkbox
        self.use_same_prompt = tk.BooleanVar(value=True)
        self.same_prompt_check = ttk.Checkbutton(
            self.batch_context_frame, 
            text="Use same prompt for all files",
            variable=self.use_same_prompt,
            command=self._toggle_same_prompt
        )
        self.same_prompt_check.pack(anchor=tk.W, pady=(4, 0))
        
        # Initialize batch prompt storage
        self._batch_prompts = []
        self._current_prompt_index = 0
        
        hint_label = ttk.Label(settings_frame, text="Names, terms, companies to help recognition...", 
                 font=("Segoe UI", 8))
        hint_label.grid(row=2, column=1, sticky=tk.W, padx=8, pady=(2, 0))
        
        # Model section (row 2) - will be collapsed dynamically if GPU+CUDA detected
        self.model_section = CollapsibleFrame(left_frame, text="Model")
        self.model_section.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        self.model_section.collapse()  # Start collapsed
        model_frame = self.model_section.content
        
        ttk.Label(model_frame, text="Model:").grid(row=0, column=0, sticky=tk.W)
        self.model_combo = ttk.Combobox(model_frame, textvariable=self.model, 
                                   values=self.get_available_models(),
                                   state="readonly", width=15)
        self.model_combo.grid(row=0, column=1, sticky=tk.W, padx=(8, 12))
        
        # Update model section status when model changes
        self.model_combo.bind('<<ComboboxSelected>>', lambda e: self.update_model_status())
        
        ttk.Checkbutton(model_frame, text="Use GPU (if available)", 
                       variable=self.use_gpu,
                       command=self.update_model_status).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(8,0))
        
        self.device_label = ttk.Label(model_frame, textvariable=self.device_info, font=("Segoe UI", 8))
        self.device_label.grid(row=1, column=2, sticky=tk.W, pady=(8,0))
        
        ttk.Button(model_frame, text="Setup GPU", 
                  command=self.manual_gpu_setup).grid(row=1, column=3, padx=5, pady=(8,0))
        
        # Anti-hallucination options (collapsed by default)
        ttk.Separator(model_frame, orient='horizontal').grid(row=3, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(10, 5))
        
        ttk.Label(model_frame, text="Anti-hallucination:", font=("Segoe UI", 9, "bold")).grid(row=4, column=0, sticky=tk.W)
        
        # Condition on previous text checkbox
        ttk.Checkbutton(model_frame, text="Condition on previous text", 
                       variable=self.condition_on_previous_text).grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=(4,0))
        ttk.Label(model_frame, text="(uncheck to reduce hallucination cascades)", 
                 font=("Segoe UI", 8)).grid(row=5, column=2, columnspan=2, sticky=tk.W, pady=(4,0))
        
        # No speech threshold
        ttk.Label(model_frame, text="No speech threshold:").grid(row=6, column=0, sticky=tk.W, pady=(4,0))
        no_speech_spin = ttk.Spinbox(model_frame, from_=0.0, to=1.0, increment=0.1, 
                                     textvariable=self.no_speech_threshold, width=6)
        no_speech_spin.grid(row=6, column=1, sticky=tk.W, padx=8, pady=(4,0))
        ttk.Label(model_frame, text="(0.6 default, lower = more sensitive)", 
                 font=("Segoe UI", 8)).grid(row=6, column=2, columnspan=2, sticky=tk.W, pady=(4,0))
        
        # Hallucination silence threshold  
        ttk.Label(model_frame, text="Silence threshold:").grid(row=7, column=0, sticky=tk.W, pady=(4,0))
        silence_spin = ttk.Spinbox(model_frame, from_=0.0, to=5.0, increment=0.5, 
                                   textvariable=self.hallucination_silence_threshold, width=6)
        silence_spin.grid(row=7, column=1, sticky=tk.W, padx=8, pady=(4,0))
        ttk.Label(model_frame, text="(seconds, 0=off, try 2.0 if hallucinating)", 
                 font=("Segoe UI", 8)).grid(row=7, column=2, columnspan=2, sticky=tk.W, pady=(4,0))
        
        # Initial model status update
        self.update_model_status()
        
        # Subtitle formatting settings (row 3)
        self.subtitle_section = CollapsibleFrame(left_frame, text="Subtitle Formatting", collapsed=True)
        self.subtitle_section.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        format_frame = self.subtitle_section.content
        
        ttk.Label(format_frame, text="Max chars/line:").grid(row=0, column=0, sticky=tk.W)
        ttk.Spinbox(format_frame, from_=30, to=50, textvariable=self.max_chars_per_line, 
                   width=8).grid(row=0, column=1, sticky=tk.W, padx=8)
        
        ttk.Label(format_frame, text="Max lines:").grid(row=0, column=2, sticky=tk.W, padx=(16, 0))
        ttk.Spinbox(format_frame, from_=1, to=3, textvariable=self.max_line_count, 
                   width=8).grid(row=0, column=3, sticky=tk.W, padx=8)
        
        # Smart formatting checkbox
        ttk.Checkbutton(format_frame, text="Smart formatting (word-level timestamps)", 
                       variable=self.use_word_timestamps,
                       command=self.update_features_display).grid(row=1, column=0, columnspan=4, 
                                                                   sticky=tk.W, pady=(10, 4))
        
        # Features label (will be updated dynamically)
        self.features_label = ttk.Label(format_frame, text="", font=("Segoe UI", 8))
        self.features_label.grid(row=2, column=0, columnspan=4, sticky=tk.W, padx=(20, 0))
        self.update_features_display()  # Set initial text
        
        # AI Proofreading section (row 4)
        self.ai_section = CollapsibleFrame(left_frame, text="AI Proofreading", collapsed=True)
        self.ai_section.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        proofread_frame = self.ai_section.content
        
        ttk.Checkbutton(proofread_frame, text="Enable AI proofreading (Claude Sonnet)", 
                       variable=self.use_ai_proofreading,
                       command=self.toggle_api_key_field).grid(row=0, column=0, columnspan=3, 
                                                                sticky=tk.W)
        
        ttk.Label(proofread_frame, text="Fixes transcription errors, grammar, and punctuation", 
                 font=("Segoe UI", 8)).grid(row=1, column=0, columnspan=3, sticky=tk.W, padx=(20, 0))
        
        ttk.Label(proofread_frame, text="API Key:").grid(row=2, column=0, sticky=tk.W, pady=(8, 0))
        self.api_key_entry = ttk.Entry(proofread_frame, textvariable=self.anthropic_api_key, 
                                        width=32, show="*")
        self.api_key_entry.grid(row=2, column=1, sticky=tk.W, padx=8, pady=(8, 0))
        
        self.save_key_btn = ttk.Button(proofread_frame, text="Save", 
                                        command=self.save_api_key, width=6)
        self.save_key_btn.grid(row=2, column=2, pady=(8, 0))
        
        # Initially disable API key field if proofreading is off
        self.toggle_api_key_field()
        
        # Update proofreading status indicator
        self.update_proofreading_status()
        
        # Process buttons
        button_frame = ttk.Frame(left_frame)
        button_frame.grid(row=5, column=0, pady=(5, 10), sticky=(tk.W, tk.E))
        
        self.process_btn = ttk.Button(button_frame, text="▶  Start Transcription", 
                                     command=self.start_transcription, style="Accent.TButton")
        self.process_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 8))
        
        self.stop_btn = ttk.Button(button_frame, text="■  Stop", 
                                   command=self.stop_transcription, state='disabled')
        self.stop_btn.pack(side=tk.LEFT, expand=True, fill=tk.X)
        
        # Batch file label (shown during processing) - Text widget for mixed formatting
        # sv_ttk dark theme background is #1c1c1c
        self.batch_file_label = tk.Text(left_frame, height=1, width=60, 
                                        bg='#1c1c1c', fg='#d4d4d4', 
                                        relief='flat', font=('Segoe UI', 9),
                                        highlightthickness=0, borderwidth=0,
                                        pady=0, padx=0)
        self.batch_file_label.grid(row=6, column=0, sticky=tk.W, pady=(8, 6))
        self.batch_file_label.tag_configure('normal', foreground='#d4d4d4')
        self.batch_file_label.tag_configure('bold', foreground='#ffffff', font=('Segoe UI', 9, 'bold'))
        self.batch_file_label.configure(state='disabled')
        
        # Progress
        self.progress = GradientProgressBar(left_frame)
        self.progress.grid(row=7, column=0, sticky=(tk.W, tk.E), pady=(0, 4))
        
        # Progress label as Text widget for colored segments
        # sv_ttk dark theme background is #1c1c1c
        self.progress_label = tk.Text(left_frame, height=1, width=70, 
                                      bg='#1c1c1c', fg='#d4d4d4', 
                                      relief='flat', font=('Segoe UI', 10),
                                      highlightthickness=0, pady=2, padx=0,
                                      borderwidth=0)
        self.progress_label.grid(row=8, column=0, sticky=tk.W, pady=(4, 0))
        self.progress_label.configure(state='disabled')
        
        # Configure tags for progress label colors
        self.progress_label.tag_configure('normal', foreground='#d4d4d4')
        self.progress_label.tag_configure('separator', foreground='#6a4c93')  # Dark purple separators
        self.progress_label.tag_configure('dim', foreground='#888888')
        
        # Output log (right side)
        log_frame = ttk.LabelFrame(right_frame, text="  Output Log  ", padding="12")
        log_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Create log with dark theme
        self.log_text = scrolledtext.ScrolledText(log_frame, height=32, width=70, state='disabled',
                                                   bg='#1e1e1e', fg='#d4d4d4', insertbackground='#ffffff',
                                                   relief='flat', font=('Consolas', 9), padx=8, pady=8)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure color tags for log
        self.log_text.tag_configure('success', foreground='#4ec9b0')      # Teal/green for success ✅
        self.log_text.tag_configure('error', foreground='#f14c4c')        # Red for errors ❌
        self.log_text.tag_configure('warning', foreground='#cca700')      # Yellow/orange for warnings ⚠️
        self.log_text.tag_configure('info', foreground='#3794ff')         # Blue for info 📥 💡
        self.log_text.tag_configure('timestamp', foreground='#6a4c93')    # Dark purple for timestamps
        self.log_text.tag_configure('transcript', foreground='#9d7cd8')   # Lighter purple for transcript text
        self.log_text.tag_configure('header', foreground='#569cd6')       # Blue for headers/separators
        self.log_text.tag_configure('dim', foreground='#6a6a6a')          # Dim gray for less important info
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(0, weight=0)  # Header
        main_frame.rowconfigure(1, weight=1)  # Content
        content_frame.columnconfigure(0, weight=0)
        content_frame.columnconfigure(1, weight=1)
        content_frame.rowconfigure(0, weight=1)
        left_frame.columnconfigure(0, weight=1)
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
    def log(self, message):
        """Add message to log with color support (thread-safe)"""
        def do_log():
            self.log_text.config(state='normal')
            
            # Determine color tag based on message content
            import re
            
            # Check for timestamp pattern [MM:SS.mmm --> MM:SS.mmm] text
            timestamp_match = re.match(r'(\[[\d:.]+ --> [\d:.]+\])(.*)', message)
            if timestamp_match:
                # Insert timestamp in darker purple
                self.log_text.insert(tk.END, timestamp_match.group(1), 'timestamp')
                # Insert transcript text in lighter purple
                self.log_text.insert(tk.END, timestamp_match.group(2) + "\n", 'transcript')
            elif message.startswith('✅') or message.startswith('☑'):
                self.log_text.insert(tk.END, message + "\n", 'success')
            elif message.startswith('❌'):
                self.log_text.insert(tk.END, message + "\n", 'error')
            elif message.startswith('⚠️') or message.startswith('⚠'):
                self.log_text.insert(tk.END, message + "\n", 'warning')
            elif message.startswith('📥') or message.startswith('💡') or message.startswith('🔍') or message.startswith('🎤') or message.startswith('📁') or message.startswith('🎬') or message.startswith('📄'):
                self.log_text.insert(tk.END, message + "\n", 'info')
            elif message.strip().startswith('===') or message.strip().startswith('---') or message.startswith('═') or (len(message.strip()) > 10 and len(message.strip().replace('=', '').replace('-', '')) == 0):
                self.log_text.insert(tk.END, message + "\n", 'header')
            elif message.startswith('  Batch') or message.startswith('Processing in batches'):
                self.log_text.insert(tk.END, message + "\n", 'dim')
            else:
                self.log_text.insert(tk.END, message + "\n")
            
            self.log_text.see(tk.END)
            self.log_text.config(state='disabled')
        
        try:
            # If we're in the main thread, do it directly
            # Otherwise, schedule it
            self.root.after(0, do_log)
        except:
            # Fallback if something goes wrong
            pass
    
    def set_progress_text(self, text):
        """Set progress label text with colored separators"""
        self.progress_label.configure(state='normal')
        self.progress_label.delete('1.0', tk.END)
        
        # Split by | separator and add with colors
        parts = text.split(' | ')
        for i, part in enumerate(parts):
            if i > 0:
                self.progress_label.insert(tk.END, '  |  ', 'separator')
            self.progress_label.insert(tk.END, part, 'normal')
        
        self.progress_label.configure(state='disabled')
    
    def get_progress_text(self):
        """Get current progress label text"""
        return self.progress_label.get('1.0', tk.END).strip()
    
    def set_batch_file_text(self, prefix="", filename=""):
        """Set batch file label text with bold filename"""
        self.batch_file_label.configure(state='normal')
        self.batch_file_label.delete('1.0', tk.END)
        
        if prefix or filename:
            self.batch_file_label.insert(tk.END, prefix, 'normal')
            self.batch_file_label.insert(tk.END, filename, 'bold')
        
        self.batch_file_label.configure(state='disabled')
    
    def clear_batch_file_text(self):
        """Clear the batch file label"""
        self.batch_file_label.configure(state='normal')
        self.batch_file_label.delete('1.0', tk.END)
        self.batch_file_label.configure(state='disabled')
    
    def update_features_display(self):
        """Update the features label based on smart formatting checkbox"""
        if self.use_word_timestamps.get():
            text = "✓ Sentence splitting  ✓ Balanced lines  ✓ No orphaned words"
        else:
            text = "✗ Basic formatting only (no smart splitting)"
        self.features_label.config(text=text)
    
    def toggle_api_key_field(self):
        """Enable/disable API key field based on proofreading checkbox"""
        if self.use_ai_proofreading.get():
            self.api_key_entry.config(state='normal')
            self.save_key_btn.config(state='normal')
        else:
            self.api_key_entry.config(state='disabled')
            self.save_key_btn.config(state='disabled')
        # Update the section header status
        self.update_proofreading_status()
    
    def update_proofreading_status(self):
        """Update the AI Proofreading section header with on/off status"""
        if hasattr(self, 'ai_section'):
            if self.use_ai_proofreading.get():
                self.ai_section.set_status("✓ ON")
            else:
                self.ai_section.set_status("OFF")
    
    def update_model_status(self):
        """Update the Model section header with model name and GPU status"""
        if hasattr(self, 'model_section'):
            model_name = self.model.get()
            gpu_status = "GPU" if self.use_gpu.get() else "CPU"
            self.model_section.set_status(f"{model_name} • {gpu_status}")
    
    def play_completion_chime(self):
        """Play a pleasant chime sound when transcription is complete."""
        def play():
            try:
                import sys
                if sys.platform == 'win32':
                    # Windows: Use pleasant system notification sound
                    import winsound
                    # Play the Windows "Mail Notification" or asterisk sound
                    winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS | winsound.SND_ASYNC)
                else:
                    # macOS/Linux: Try system bell or afplay on mac
                    try:
                        if sys.platform == 'darwin':
                            import subprocess
                            # Use macOS system sound
                            subprocess.run(['afplay', '/System/Library/Sounds/Glass.aiff'], 
                                         capture_output=True)
                        else:
                            # Linux: Try paplay or just print bell
                            print('\a')  # Terminal bell
                    except:
                        print('\a')  # Fallback to terminal bell
            except:
                pass  # Silently fail if sound doesn't work
        
        # Play in background thread to not block UI
        threading.Thread(target=play, daemon=True).start()
    
    def check_for_updates(self):
        """Check GitHub for newer version in background."""
        def check():
            try:
                remote_version = None
                
                # For frozen EXE, use subprocess to avoid SSL issues
                if getattr(sys, 'frozen', False):
                    try:
                        result = _run_hidden(
                            ['py', '-c', f'''
import urllib.request
req = urllib.request.Request("{UPDATE_CHECK_URL}", headers={{"User-Agent": "Final-Whisper"}})
with urllib.request.urlopen(req, timeout=10) as r:
    print(r.read().decode("utf-8").strip())
'''],
                            capture_output=True, text=True, timeout=15
                        )
                        if result.returncode == 0 and result.stdout.strip():
                            remote_version = result.stdout.strip()
                    except Exception as e:
                        pass  # Silently fail
                else:
                    import urllib.request
                    req = urllib.request.Request(
                        UPDATE_CHECK_URL,
                        headers={'User-Agent': 'Final-Whisper-Update-Checker'}
                    )
                    with urllib.request.urlopen(req, timeout=10) as response:
                        remote_version = response.read().decode('utf-8').strip()
                
                # Compare versions
                if remote_version:
                    if self._is_newer_version(remote_version, VERSION):
                        # Show update prompt in main thread
                        self.root.after(0, lambda: self._show_update_dialog(remote_version))
            except Exception as e:
                pass  # Silently fail - don't bother user if check fails
        
        threading.Thread(target=check, daemon=True).start()
    
    def _is_newer_version(self, remote, local):
        """Compare version strings. Returns True if remote is newer."""
        try:
            remote_parts = [int(x) for x in remote.split('.')]
            local_parts = [int(x) for x in local.split('.')]
            return remote_parts > local_parts
        except:
            return False
    
    def _show_update_dialog(self, new_version):
        """Show update button when a new version is available."""
        self.new_version = new_version
        self.update_button.pack(side=tk.LEFT, padx=10)
        self.log(f"✨ Update available: v{new_version} (current: v{VERSION})")

    def download_and_install_update(self):
        """Download and install the latest version."""
        if not self.new_version:
            return

        try:
            import tempfile

            # Disable the update button during download
            self.update_button.config(state='disabled', text='⏳ Downloading update...')
            self.log(f"📥 Downloading Final Whisper v{self.new_version}...")

            def download_and_replace():
                try:
                    # Download the latest EXE
                    download_url = f"https://github.com/{GITHUB_REPO}/releases/latest/download/FinalWhisper.exe"

                    temp_dir = tempfile.gettempdir()
                    new_exe_path = os.path.join(temp_dir, "FinalWhisper_New.exe")

                    # For frozen EXE, use subprocess to download (avoids SSL issues)
                    if getattr(sys, 'frozen', False):
                        result = _run_hidden(
                            ['py', '-c', f'''
import urllib.request
urllib.request.urlretrieve("{download_url}", r"{new_exe_path}")
print("OK")
'''],
                            capture_output=True, text=True, timeout=300  # 5 min timeout
                        )
                        if result.returncode != 0 or "OK" not in result.stdout:
                            raise Exception(f"Download failed: {result.stderr}")
                    else:
                        import urllib.request
                        urllib.request.urlretrieve(download_url, new_exe_path)

                    self.root.after(0, lambda: self.log(f"✅ Downloaded successfully!"))
                    self.root.after(0, lambda: self.log(f"🔄 Installing update..."))

                    # Get current EXE path
                    if getattr(sys, 'frozen', False):
                        current_exe = sys.executable
                    else:
                        # For testing in dev mode
                        self.root.after(0, lambda: messagebox.showinfo(
                            "Update Downloaded",
                            f"New version downloaded to:\n{new_exe_path}\n\n"
                            "(Auto-install only works in compiled EXE)"
                        ))
                        return

                    # Create a batch script to replace the EXE after we exit
                    batch_script = f"""@echo off
title Updating Final Whisper...
echo Updating Final Whisper...

REM Wait for app to close
timeout /t 2 /nobreak >nul
taskkill /F /IM "FinalWhisper.exe" >nul 2>&1
timeout /t 1 /nobreak >nul

REM Replace executable
del /F /Q "{current_exe}" 2>nul
move /Y "{new_exe_path}" "{current_exe}"
if errorlevel 1 (
    echo ERROR: Update failed. Please try again.
    pause
    exit /b 1
)

echo.
echo Update complete! You can now open Final Whisper.
del "%~f0"
"""

                    batch_path = os.path.join(temp_dir, "update_final_whisper.bat")
                    with open(batch_path, 'w') as f:
                        f.write(batch_script)

                    # Show message and wait for user to click OK before exiting
                    def show_and_exit():
                        messagebox.showinfo(
                            "Update Ready",
                            f"Final Whisper v{self.new_version} has been downloaded!\n\n"
                            "The application will now close and update.\n\n"
                            "Please reopen Final Whisper after a few seconds."
                        )
                        # Launch update script and exit AFTER user clicks OK
                        subprocess.Popen(['cmd', '/c', batch_path],
                                       creationflags=subprocess.CREATE_NO_WINDOW)
                        self.root.quit()
                    
                    self.root.after(0, show_and_exit)

                except Exception as e:
                    self.root.after(0, lambda: self.log(f"❌ Update failed: {e}"))
                    self.root.after(0, lambda: messagebox.showerror(
                        "Update Failed",
                        f"Could not download update:\n{e}\n\n"
                        f"Please download manually from:\n{RELEASES_URL}"
                    ))
                    self.root.after(0, lambda: self.update_button.config(
                        state='normal', text='🔄 Update Available - Click to Install'
                    ))

            # Run download in background thread
            threading.Thread(target=download_and_replace, daemon=True).start()

        except Exception as e:
            self.log(f"❌ Update error: {e}")
            messagebox.showerror("Update Error", f"Failed to start update:\n{e}")
            self.update_button.config(state='normal', text='🔄 Update Available - Click to Install')

    def restart_application(self):
        """Restart the application"""
        try:
            import time

            # Create a batch script to restart after this instance closes
            if getattr(sys, 'frozen', False):
                exe_path = sys.executable
            else:
                exe_path = sys.executable
                script_path = __file__

            # On Windows, use a batch file to restart
            if os.name == 'nt':
                import tempfile
                batch_script = f"""@echo off
echo Restarting Final Whisper...
timeout /t 2 /nobreak >nul
start "" "{exe_path}" {'' if getattr(sys, 'frozen', False) else f'"{script_path}"'}
del "%~f0"
"""
                batch_path = os.path.join(tempfile.gettempdir(), "restart_final_whisper.bat")
                with open(batch_path, 'w') as f:
                    f.write(batch_script)

                # Launch the restart script
                subprocess.Popen(['cmd', '/c', batch_path],
                               creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0)
            else:
                # On other platforms, just start a new instance
                if getattr(sys, 'frozen', False):
                    subprocess.Popen([exe_path])
                else:
                    subprocess.Popen([exe_path, script_path])

            # Close current instance
            self.log("Closing application for restart...")
            self.root.after(500, self.root.destroy)  # Use destroy instead of quit

        except Exception as e:
            self.log(f"❌ Failed to restart: {e}")
            import traceback
            self.log(traceback.format_exc())
            messagebox.showerror("Restart Failed",
                f"Please manually close and restart the application.\n\nError: {e}")

    def get_config_path(self):
        """Get path to config file"""
        config_dir = Path.home() / ".final_whisper"
        config_dir.mkdir(exist_ok=True)
        return config_dir / "config.json"
    
    def load_config(self):
        """Load config from file"""
        try:
            config_path = self.get_config_path()
            if config_path.exists():
                import json
                with open(config_path, 'r') as f:
                    return json.load(f)
        except Exception:
            pass
        return {}

    def save_config(self, config):
        """Save config to file"""
        try:
            import json
            config_path = self.get_config_path()
            with open(config_path, 'w') as f:
                json.dump(config, f)
        except Exception as e:
            self.log(f"⚠️ Failed to save config: {e}")

    def load_settings(self):
        """Load user settings from config"""
        try:
            config = self.load_config()
            
            # Model settings
            if 'model' in config:
                self.model.set(config['model'])
            if 'language' in config:
                self.language.set(config['language'])
            if 'use_gpu' in config:
                self.use_gpu.set(config['use_gpu'])
            
            # Subtitle formatting
            if 'max_chars_per_line' in config:
                self.max_chars_per_line.set(config['max_chars_per_line'])
            if 'max_line_count' in config:
                self.max_line_count.set(config['max_line_count'])
            if 'use_word_timestamps' in config:
                self.use_word_timestamps.set(config['use_word_timestamps'])
            
            # Anti-hallucination settings
            if 'condition_on_previous_text' in config:
                self.condition_on_previous_text.set(config['condition_on_previous_text'])
            if 'no_speech_threshold' in config:
                self.no_speech_threshold.set(config['no_speech_threshold'])
            if 'hallucination_silence_threshold' in config:
                self.hallucination_silence_threshold.set(config['hallucination_silence_threshold'])
            
            # AI Proofreading
            if 'use_ai_proofreading' in config:
                self.use_ai_proofreading.set(config['use_ai_proofreading'])
            
            # Output directory
            if 'output_dir' in config and config['output_dir']:
                self.output_dir.set(config['output_dir'])
                
        except Exception as e:
            pass  # Silently fail - will use defaults
    
    def save_settings(self):
        """Save user settings to config"""
        try:
            config = self.load_config()
            
            # Model settings
            config['model'] = self.model.get()
            config['language'] = self.language.get()
            config['use_gpu'] = self.use_gpu.get()
            
            # Subtitle formatting
            config['max_chars_per_line'] = self.max_chars_per_line.get()
            config['max_line_count'] = self.max_line_count.get()
            config['use_word_timestamps'] = self.use_word_timestamps.get()
            
            # Anti-hallucination settings
            config['condition_on_previous_text'] = self.condition_on_previous_text.get()
            config['no_speech_threshold'] = self.no_speech_threshold.get()
            config['hallucination_silence_threshold'] = self.hallucination_silence_threshold.get()
            
            # AI Proofreading
            config['use_ai_proofreading'] = self.use_ai_proofreading.get()
            
            # Output directory (only save if not default)
            output_dir = self.output_dir.get()
            default_output = str(Path.home() / "whisper_output")
            if output_dir != default_output:
                config['output_dir'] = output_dir
            
            self.save_config(config)
        except Exception as e:
            pass  # Silently fail

    def load_api_key(self):
        """Load API key from config file"""
        try:
            config = self.load_config()
            if 'anthropic_api_key' in config and config['anthropic_api_key']:
                self.anthropic_api_key.set(config['anthropic_api_key'])
                # Auto-enable proofreading if we have a saved API key
                self.use_ai_proofreading.set(True)
        except Exception:
            pass
    
    def save_api_key(self):
        """Save API key to config file"""
        try:
            config = self.load_config()
            config['anthropic_api_key'] = self.anthropic_api_key.get()
            self.save_config(config)
            self.log("✅ API key saved")
        except Exception as e:
            self.log(f"❌ Failed to save API key: {e}")
    
    def _proofread_via_subprocess(self, srt_file, language="da", context=""):
        """Run AI proofreading via subprocess (for frozen EXE with SSL issues)"""
        api_key = self.anthropic_api_key.get().strip()
        if not api_key:
            self.log("⚠️ No API key provided, skipping proofreading")
            return None
        
        self.log("\n🔍 Starting AI proofreading...")
        
        # Create a proofreading script
        proofread_script = '''
import sys, json, urllib.request, urllib.error, re, time

api_key = sys.argv[1]
srt_file = sys.argv[2]
language = sys.argv[3]
context = sys.argv[4] if len(sys.argv) > 4 else ""

# Language names
lang_names = {'da': 'Danish', 'en': 'English', 'de': 'German', 'fr': 'French',
              'es': 'Spanish', 'it': 'Italian', 'no': 'Norwegian', 'sv': 'Swedish'}
lang_name = lang_names.get(language, language.title())

# Read SRT file
with open(srt_file, 'r', encoding='utf-8') as f:
    srt_content = f.read()

subtitle_blocks = re.split(r'\\n\\n+', srt_content.strip())
total = len(subtitle_blocks)
print(f"Found {total} subtitles", flush=True)

BATCH_SIZE = 60
corrected_blocks = []

def proofread_batch(batch_content, lang_name, api_key, context):
    system_prompt = f"""You are a professional subtitle proofreader for {lang_name} content.
Fix grammar, spelling, and punctuation errors while preserving the SRT format exactly.
Keep timestamps unchanged. Only fix the text content.
Return ONLY the corrected SRT content, nothing else."""
    
    user_prompt = f"Proofread these {lang_name} subtitles"
    if context:
        user_prompt += f" (context: {context})"
    user_prompt += f":\\n\\n{batch_content}"
    
    data = json.dumps({
        "model": "claude-sonnet-4-20250514",
        "max_tokens": 8192,
        "messages": [{"role": "user", "content": user_prompt}],
        "system": system_prompt
    }).encode('utf-8')
    
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=data,
        headers={
            "Content-Type": "application/json",
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01"
        }
    )
    
    with urllib.request.urlopen(req, timeout=120) as response:
        result = json.loads(response.read().decode('utf-8'))
        return result['content'][0]['text']

if total > BATCH_SIZE:
    print(f"Processing in batches of {BATCH_SIZE}...", flush=True)
    total_batches = (total + BATCH_SIZE - 1) // BATCH_SIZE
    
    for batch_start in range(0, total, BATCH_SIZE):
        batch_end = min(batch_start + BATCH_SIZE, total)
        batch_num = (batch_start // BATCH_SIZE) + 1
        print(f"BATCH:{batch_num}/{total_batches}", flush=True)
        
        batch_blocks = subtitle_blocks[batch_start:batch_end]
        batch_content = '\\n\\n'.join(batch_blocks)
        
        # Retry logic for rate limits and overload errors
        max_retries = 3
        for attempt in range(max_retries):
            try:
                result = proofread_batch(batch_content, lang_name, api_key, context)
                corrected_batch_blocks = re.split(r'\\n\\n+', result.strip())
                corrected_blocks.extend(corrected_batch_blocks)
                print(f"OK:{batch_num}", flush=True)
                break
            except urllib.error.HTTPError as e:
                if e.code in (429, 529, 503, 502) and attempt < max_retries - 1:
                    wait_time = 30 * (attempt + 1)  # 30s, 60s, 90s
                    print(f"RETRY:{batch_num}:Error {e.code}, waiting {wait_time}s... (attempt {attempt+1}/{max_retries})", flush=True)
                    time.sleep(wait_time)
                else:
                    print(f"FAIL:{batch_num}:{e}", flush=True)
                    corrected_blocks.extend(batch_blocks)
                    break
            except Exception as e:
                print(f"FAIL:{batch_num}:{e}", flush=True)
                corrected_blocks.extend(batch_blocks)
                break
        
        # Add delay between batches to avoid rate limits (except after last batch)
        if batch_end < total:
            time.sleep(2)
    
    corrected_content = '\\n\\n'.join(corrected_blocks) + '\\n'
else:
    try:
        corrected_content = proofread_batch(srt_content, lang_name, api_key, context)
        print("OK:1", flush=True)
    except Exception as e:
        print(f"FAIL:1:{e}", flush=True)
        corrected_content = srt_content

# Write corrected content
with open(srt_file, 'w', encoding='utf-8') as f:
    f.write(corrected_content)

print("DONE", flush=True)
'''
        
        cmd = ['py', '-c', proofread_script, api_key, srt_file, language, context]
        
        # Hide console window
        startupinfo = None
        creationflags = 0
        if os.name == 'nt':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            creationflags = subprocess.CREATE_NO_WINDOW
        
        try:
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding='utf-8',
                errors='replace',
                startupinfo=startupinfo,
                creationflags=creationflags
            )
            
            for line in process.stdout:
                line = line.strip()
                if not line:
                    continue
                
                if line.startswith('Found'):
                    self.log(line)
                elif line.startswith('Processing'):
                    self.log(line)
                elif line.startswith('BATCH:'):
                    parts = line[6:].split('/')
                    batch_num, total_batches = int(parts[0]), int(parts[1])
                    self._proofreading_batch_info = f"Batch {batch_num}/{total_batches}"
                    self._proofreading_total_batches = total_batches
                    self.log(f"  Batch {batch_num}/{total_batches}...")
                elif line.startswith('RETRY:'):
                    # Retrying due to error (429, 529, 503, etc.)
                    parts = line[6:].split(':', 1)
                    batch_num = parts[0]
                    message = parts[1] if len(parts) > 1 else "Retrying..."
                    self._proofreading_batch_info = f"Batch {batch_num} - {message}"
                    self.log(f"  ⏳ Batch {batch_num}: {message}")
                elif line.startswith('OK:'):
                    # Batch succeeded - update progress bar
                    batch_num = int(line[3:])
                    total_batches = getattr(self, '_proofreading_total_batches', 1)
                    # Progress goes from 85% to 100% during proofreading
                    progress_per_batch = 15 / total_batches
                    new_progress = 85 + (batch_num * progress_per_batch)
                    self.root.after(0, lambda p=new_progress: self.progress.set_value(min(p, 100)))
                elif line.startswith('FAIL:'):
                    # Batch failed - still update progress
                    parts = line[5:].split(':', 1)
                    batch_num = int(parts[0])
                    error = parts[1] if len(parts) > 1 else "Unknown error"
                    total_batches = getattr(self, '_proofreading_total_batches', 1)
                    progress_per_batch = 15 / total_batches
                    new_progress = 85 + (batch_num * progress_per_batch)
                    self.root.after(0, lambda p=new_progress: self.progress.set_value(min(p, 100)))
                    self.log(f"  ⚠️ Batch {batch_num} failed: {error}")
                elif line == 'DONE':
                    self.root.after(0, lambda: self.progress.set_value(100))
                    self.log("✅ Proofreading complete!")
            
            process.wait()
            self._proofreading_batch_info = None
            return srt_file
            
        except Exception as e:
            self.log(f"❌ Proofreading error: {e}")
            return None
    
    def proofread_srt_with_ai(self, srt_file, language="da", context=""):
        """Use Claude API to proofread and fix the SRT file"""
        import json
        import urllib.request
        import urllib.error
        import re
        
        api_key = self.anthropic_api_key.get().strip()
        if not api_key:
            self.log("⚠️ No API key provided, skipping proofreading")
            return None
        
        self.log("\n🔍 Starting AI proofreading...")
        if context:
            self.log(f"Using context: {context[:80]}{'...' if len(context) > 80 else ''}")
        
        # Read the SRT file
        with open(srt_file, 'r', encoding='utf-8') as f:
            srt_content = f.read()
        
        # Language names for the prompt
        lang_names = {
            'da': 'Danish', 'en': 'English', 'de': 'German', 'fr': 'French',
            'es': 'Spanish', 'it': 'Italian', 'no': 'Norwegian', 'sv': 'Swedish',
            'nl': 'Dutch', 'pl': 'Polish', 'pt': 'Portuguese', 'fi': 'Finnish',
            'ja': 'Japanese', 'ko': 'Korean', 'zh': 'Chinese', 'ru': 'Russian'
        }
        lang_name = lang_names.get(language, language.title())
        
        # Store context for batch processing
        self._proofread_context = context
        
        # Parse subtitles into blocks
        subtitle_blocks = re.split(r'\n\n+', srt_content.strip())
        total_subtitles = len(subtitle_blocks)
        
        self.log(f"Found {total_subtitles} subtitles")
        
        # Process in batches if needed (60 subtitles per batch)
        BATCH_SIZE = 60
        
        if total_subtitles > BATCH_SIZE:
            self.log(f"Processing in batches of {BATCH_SIZE}...")
            corrected_blocks = []
            total_cost = 0
            total_batches = (total_subtitles + BATCH_SIZE - 1) // BATCH_SIZE
            
            for batch_start in range(0, total_subtitles, BATCH_SIZE):
                batch_end = min(batch_start + BATCH_SIZE, total_subtitles)
                batch_num = (batch_start // BATCH_SIZE) + 1
                
                # Update progress label with batch info
                self._proofreading_batch_info = f"Batch {batch_num}/{total_batches}"
                
                batch_blocks = subtitle_blocks[batch_start:batch_end]
                batch_content = '\n\n'.join(batch_blocks)
                
                self.log(f"  Batch {batch_num}/{total_batches} (subtitles {batch_start+1}-{batch_end})...")
                
                result, cost = self._proofread_batch(batch_content, lang_name, api_key)
                if result is None:
                    self.log(f"  ⚠️ Batch {batch_num} failed, keeping original")
                    corrected_blocks.extend(batch_blocks)
                else:
                    # Parse the corrected batch back into blocks
                    corrected_batch_blocks = re.split(r'\n\n+', result.strip())
                    corrected_blocks.extend(corrected_batch_blocks)
                    total_cost += cost
            
            # Clear batch info
            self._proofreading_batch_info = None
            
            corrected_content = '\n\n'.join(corrected_blocks) + '\n'
            
            self.log(f"✅ Proofreading complete!")
            self.log(f"   Estimated total cost: ${total_cost:.4f}")
        else:
            # Small file - process all at once
            self.log(f"Sending to Claude API...")
            corrected_content, cost = self._proofread_batch(srt_content, lang_name, api_key)
            
            if corrected_content is None:
                return None
            
            self.log(f"✅ Proofreading complete!")
            self.log(f"   Estimated cost: ${cost:.4f}")
        
        # Overwrite the original SRT with the corrected version
        with open(srt_file, 'w', encoding='utf-8') as f:
            f.write(corrected_content)
        
        # Count changes (rough estimate)
        original_lines = srt_content.strip().split('\n')
        corrected_lines = corrected_content.strip().split('\n')
        changes = sum(1 for a, b in zip(original_lines, corrected_lines) if a != b)
        
        self.log(f"   ~{changes} lines modified")
        self.log(f"   Saved to: {Path(srt_file).name}")
        
        return srt_file  # Return the same file path (now contains proofread content)
    
    def _proofread_batch(self, srt_content, lang_name, api_key):
        """Proofread a batch of subtitles. Returns (corrected_content, cost) or (None, 0) on error."""
        import json
        import urllib.request
        import urllib.error
        
        # Get context if available
        context = getattr(self, '_proofread_context', '')
        context_section = ""
        if context:
            context_section = f"""
CONTEXT (use this to correctly spell names, terms, and understand the content):
{context}
"""
        
        # Add language-specific instructions
        lang_specific = ""
        if lang_name == "Danish":
            lang_specific = """
DANISH COMMA RULES (grammatisk komma/nyt komma):
- Add comma BEFORE subordinate clauses (ledsætninger/bisætninger)
- Comma before: at, der, som, hvis, når, fordi, hvor, hvad, hvilken, om, etc. when they start a clause
- Examples: "Jeg tror, at...", "Den bog, som jeg læste, ...", "Hvis du kommer, så..."
- Comma before "og" when it connects two main clauses with different subjects
"""
        elif lang_name == "German":
            lang_specific = """
GERMAN COMMA RULES:
- Always comma before subordinate clauses (dass, weil, wenn, obwohl, etc.)
- Comma before relative clauses (der, die, das, welcher, etc.)
- Comma before infinitive clauses with "zu" when they have objects/adverbs
- Examples: "Ich glaube, dass...", "Das Buch, das ich lese, ...", "Wenn du kommst, dann..."
"""
        elif lang_name == "Norwegian" or lang_name == "Swedish":
            lang_specific = f"""
{lang_name.upper()} COMMA RULES:
- Similar to Danish: comma before subordinate clauses
- Comma before: at/att, som, hvis/om, når/när, fordi/för att, etc.
- Comma before "og/och" when connecting main clauses with different subjects
"""
        elif lang_name == "English":
            lang_specific = """
ENGLISH COMMA RULES:
- Comma AFTER introductory clauses: "If you come, I will..."
- NO comma before "that" in restrictive clauses: "The book that I read"
- Comma before "which" in non-restrictive clauses: "The book, which I read, ..."
- Oxford comma before last item in lists is optional but be consistent
"""
        elif lang_name == "French":
            lang_specific = """
FRENCH COMMA RULES:
- Generally fewer commas than English or Danish
- No comma before "que" (that): "Je pense que..."
- Comma to separate clauses with different subjects
- Comma after long introductory phrases
"""
        elif lang_name == "Spanish":
            lang_specific = """
SPANISH COMMA RULES:
- No comma before "que" (that): "Creo que..."
- Comma before "pero", "sino", "aunque" (but, although)
- No comma between subject and verb, even with long subjects
- Comma to separate items in a list (no Oxford comma)
"""

        prompt = f"""Proofread these {lang_name} subtitles. Fix transcription errors, grammar, spelling, and punctuation.
{context_section}{lang_specific}
RULES:
- Keep EXACT same number of subtitles and timestamps
- Keep same line structure (2 lines stays 2 lines)  
- Lines must stay under 42 characters
- Only fix errors, don't rephrase
- Return ONLY the corrected SRT, no explanations

{srt_content}"""

        headers = {
            'Content-Type': 'application/json',
            'x-api-key': api_key,
            'anthropic-version': '2023-06-01'
        }
        
        data = {
            'model': 'claude-sonnet-4-20250514',
            'max_tokens': 8000,
            'messages': [
                {'role': 'user', 'content': prompt}
            ]
        }
        
        try:
            req = urllib.request.Request(
                'https://api.anthropic.com/v1/messages',
                data=json.dumps(data).encode('utf-8'),
                headers=headers,
                method='POST'
            )
            
            with urllib.request.urlopen(req, timeout=120) as response:
                result = json.loads(response.read().decode('utf-8'))
            
            if 'content' in result and len(result['content']) > 0:
                corrected_content = result['content'][0]['text']
                
                # Calculate cost
                input_tokens = len(srt_content) // 4
                output_tokens = len(corrected_content) // 4
                cost = (input_tokens * 0.003 + output_tokens * 0.015) / 1000
                
                return corrected_content, cost
            else:
                return None, 0
                
        except urllib.error.HTTPError as e:
            error_body = e.read().decode('utf-8')
            self.log(f"❌ API error: {e.code}")
            try:
                error_json = json.loads(error_body)
                if 'error' in error_json:
                    self.log(f"   {error_json['error'].get('message', error_body)}")
            except:
                self.log(f"   {error_body[:200]}")
            return None, 0
        except urllib.error.URLError as e:
            if 'timed out' in str(e).lower():
                self.log(f"❌ Request timed out")
            else:
                self.log(f"❌ Network error: {str(e)}")
            return None, 0
        except Exception as e:
            self.log(f"❌ Error: {str(e)}")
            return None, 0
    
    def get_available_models(self):
        """Get list of available Whisper models dynamically"""
        # Preferred order: turbo and large at top, then specific versions
        preferred_order = ['turbo', 'large', 'large-v3', 'large-v2', 'large-v1', 
                          'medium', 'small', 'base', 'tiny']
        
        # For frozen EXE, use subprocess to query whisper
        if getattr(sys, 'frozen', False):
            try:
                # Hide console window
                startupinfo = None
                creationflags = 0
                if os.name == 'nt':
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    startupinfo.wShowWindow = subprocess.SW_HIDE
                    creationflags = subprocess.CREATE_NO_WINDOW
                
                result = subprocess.run(
                    ['py', '-c', 'import whisper; print(",".join(whisper._MODELS.keys()))'],
                    capture_output=True, text=True, timeout=10,
                    startupinfo=startupinfo, creationflags=creationflags
                )
                if result.returncode == 0 and result.stdout.strip():
                    models = result.stdout.strip().split(',')
                    sorted_models = []
                    for m in preferred_order:
                        if m in models:
                            sorted_models.append(m)
                            models.remove(m)
                    # Add any remaining models
                    sorted_models.extend(sorted(models, reverse=True))
                    return sorted_models
            except:
                pass
        else:
            # Try to get models from whisper module directly
            try:
                import whisper
                if hasattr(whisper, '_MODELS'):
                    models = list(whisper._MODELS.keys())
                    sorted_models = []
                    for m in preferred_order:
                        if m in models:
                            sorted_models.append(m)
                            models.remove(m)
                    sorted_models.extend(sorted(models, reverse=True))
                    return sorted_models
            except:
                pass
        
        # Fallback to known models if whisper not installed yet
        return ["turbo", "large", "large-v3", "large-v2", "medium", "small", "base", "tiny"]
    
    def get_available_languages(self):
        """Get list of available Whisper languages with display names"""
        # For frozen EXE, use subprocess to query whisper
        if getattr(sys, 'frozen', False):
            try:
                # Hide console window
                startupinfo = None
                creationflags = 0
                if os.name == 'nt':
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    startupinfo.wShowWindow = subprocess.SW_HIDE
                    creationflags = subprocess.CREATE_NO_WINDOW
                
                result = subprocess.run(
                    ['py', '-c', '''
import whisper
langs = whisper.tokenizer.LANGUAGES
for code, name in sorted(langs.items(), key=lambda x: x[1]):
    print(f"{code}:{name}")
'''],
                    capture_output=True, text=True, timeout=10,
                    startupinfo=startupinfo, creationflags=creationflags
                )
                if result.returncode == 0 and result.stdout.strip():
                    lang_list = []
                    for line in result.stdout.strip().split('\n'):
                        if ':' in line:
                            code, name = line.split(':', 1)
                            lang_list.append(f"{code}: {name.title()}")
                    
                    # Move Danish and English to top
                    prioritized = []
                    for priority_code in ['da', 'en']:
                        for lang in lang_list[:]:
                            if lang.startswith(f"{priority_code}:"):
                                prioritized.append(lang)
                                lang_list.remove(lang)
                                break
                    return prioritized + lang_list
            except:
                pass
        else:
            # Try to get languages from whisper module directly
            try:
                import whisper
                if hasattr(whisper, 'tokenizer') and hasattr(whisper.tokenizer, 'LANGUAGES'):
                    languages = whisper.tokenizer.LANGUAGES
                    lang_list = [f"{code}: {name.title()}" for code, name in sorted(languages.items(), key=lambda x: x[1])]
                    prioritized = []
                    for priority_code in ['da', 'en']:
                        for lang in lang_list[:]:
                            if lang.startswith(f"{priority_code}:"):
                                prioritized.append(lang)
                                lang_list.remove(lang)
                                break
                    return prioritized + lang_list
            except:
                pass
        
        # Fallback to common languages if whisper not installed yet
        return [
            "da: Danish", "en: English", "de: German", "fr: French", 
            "es: Spanish", "it: Italian", "nl: Dutch", "no: Norwegian",
            "sv: Swedish", "fi: Finnish", "pl: Polish", "pt: Portuguese",
            "ru: Russian", "ja: Japanese", "ko: Korean", "zh: Chinese",
            "ar: Arabic", "hi: Hindi", "tr: Turkish", "uk: Ukrainian"
        ]
    
    def get_language_code(self):
        """Extract language code from the dropdown value (e.g., 'da: Danish' -> 'da')"""
        lang_value = self.language.get()
        if ':' in lang_value:
            return lang_value.split(':')[0].strip()
        return lang_value
    
    def check_whisper_installation(self):
        """Check if Whisper is installed and install if missing"""
        def check():
            # For frozen EXE, use subprocess to avoid Python version conflicts
            if getattr(sys, 'frozen', False):
                try:
                    result = _run_hidden(
                        ['py', '-c', 'import whisper; print("OK")'],
                        capture_output=True, text=True, timeout=30
                    )
                    if result.returncode == 0 and "OK" in result.stdout:
                        self.log("✅ Whisper is installed")
                        # Now check GPU availability
                        self.root.after(0, self.check_gpu_availability)
                        return
                    else:
                        self.log("⚠️ Whisper not found")
                except subprocess.TimeoutExpired:
                    self.log("⚠️ Whisper check timed out")
                except Exception as e:
                    self.log(f"⚠️ Whisper check error: {e}")
                
                # Offer to install Whisper
                self.root.after(0, self._show_whisper_install_button)
                return
            
            # Running as script - can import directly
            try:
                import whisper
                self.log("✅ Whisper is installed")
                # Now check GPU availability
                self.root.after(0, self.check_gpu_availability)
            except ImportError as e:
                self.log("⚠️ Whisper not found")
                self.log("⚠️ Whisper not found - installing automatically...")
                self.log("\n" + "="*60)
                self.log("Installing OpenAI Whisper...")
                self.log("="*60 + "\n")

                try:
                    install_cmd = [sys.executable, "-m", "pip", "install", "--user", "openai-whisper"]

                    process = subprocess.Popen(
                        install_cmd,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.STDOUT,
                        text=True,
                        bufsize=1
                    )

                    # Stream output
                    for line in process.stdout:
                        self.log(line.rstrip())

                    process.wait()

                    if process.returncode == 0:
                        self.log("\n✅ Whisper installed successfully!")
                        self.log("Ready to transcribe!\n")
                        # Now check GPU availability
                        self.root.after(0, self.check_gpu_availability)
                    else:
                        self.log("\n❌ Whisper installation failed!")
                        self.log("Please try manually: pip install --user openai-whisper\n")
                        messagebox.showerror("Installation Failed",
                            "Failed to install Whisper automatically.\n\n"
                            "Please run: pip install --user openai-whisper")

                except Exception as e:
                    self.log(f"\n❌ Error during installation: {str(e)}")
                    messagebox.showerror("Error", f"Installation error:\n{str(e)}")

        threading.Thread(target=check, daemon=True).start()
    
    def _show_whisper_install_button(self):
        """Show a prominent button to install Whisper and disable the app until installed"""
        # Collapse all sections
        if hasattr(self, 'files_section'):
            self.files_section.collapse()
        if hasattr(self, 'transcription_section'):
            self.transcription_section.collapse()
        if hasattr(self, 'model_section'):
            self.model_section.collapse()
        if hasattr(self, 'subtitle_section'):
            self.subtitle_section.collapse()
        if hasattr(self, 'ai_section'):
            self.ai_section.collapse()
        
        # Disable the start button
        self.process_btn.config(state='disabled')
        
        # Hide all sections temporarily
        if hasattr(self, 'files_section'):
            self.files_section.grid_remove()
        if hasattr(self, 'transcription_section'):
            self.transcription_section.grid_remove()
        if hasattr(self, 'model_section'):
            self.model_section.grid_remove()
        if hasattr(self, 'subtitle_section'):
            self.subtitle_section.grid_remove()
        if hasattr(self, 'ai_section'):
            self.ai_section.grid_remove()
        
        # Create a prominent install frame
        if hasattr(self, 'setup_frame'):
            self.setup_frame.destroy()
        
        self.setup_frame = ttk.Frame(self.left_panel)
        self.setup_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=20)
        
        # Title
        title_label = ttk.Label(
            self.setup_frame, 
            text="⚠️ Setup Required",
            font=('Segoe UI', 14, 'bold')
        )
        title_label.pack(pady=(0, 10))
        
        # Description
        desc_label = ttk.Label(
            self.setup_frame,
            text="OpenAI Whisper is required for transcription.\nClick the button below to install it automatically.",
            justify=tk.CENTER
        )
        desc_label.pack(pady=(0, 15))
        
        # Big install button
        self.whisper_install_btn = ttk.Button(
            self.setup_frame,
            text="📥  Install Whisper",
            command=self._install_whisper,
            style='Accent.TButton'
        )
        self.whisper_install_btn.pack(ipadx=20, ipady=10)
        
        # Status label
        self.setup_status_label = ttk.Label(
            self.setup_frame,
            text="",
            foreground='gray'
        )
        self.setup_status_label.pack(pady=(10, 0))
    
    def _install_whisper(self):
        """Install Whisper via subprocess"""
        if hasattr(self, 'whisper_install_btn'):
            self.whisper_install_btn.config(state='disabled', text='⏳ Installing...')
        if hasattr(self, 'setup_status_label'):
            self.setup_status_label.config(text="Installing Whisper... This may take a few minutes.")
        
        def install():
            self.log("\n" + "="*60)
            self.log("📦 Installing OpenAI Whisper...")
            self.log("="*60 + "\n")
            
            try:
                # Use --progress-bar off for cleaner output, we'll show our own status
                process = subprocess.Popen(
                    ['py', '-m', 'pip', 'install', '--user', '--progress-bar', 'off', 'openai-whisper'],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding='utf-8',
                    errors='replace',
                    bufsize=1,
                    creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                )
                
                current_package = ""
                for line in process.stdout:
                    line = line.rstrip()
                    if not line:
                        continue
                    
                    # Skip some noisy lines
                    if line.startswith('  ') and 'Uninstalling' not in line and 'Installing' not in line:
                        continue  # Skip indented detail lines
                    
                    self.log(line)
                    
                    # Update status with current action
                    if 'Collecting' in line:
                        pkg = line.replace('Collecting ', '').split()[0]
                        current_package = pkg
                        self.root.after(0, lambda p=pkg: self.setup_status_label.config(text=f"Collecting {p}...") if hasattr(self, 'setup_status_label') else None)
                    elif 'Downloading' in line:
                        # Extract size if available
                        import re
                        size_match = re.search(r'\(([^)]+)\)', line)
                        size_info = f" ({size_match.group(1)})" if size_match else ""
                        self.root.after(0, lambda s=size_info, p=current_package: self.setup_status_label.config(text=f"Downloading {p}{s}...") if hasattr(self, 'setup_status_label') else None)
                    elif 'Installing collected packages' in line:
                        self.root.after(0, lambda: self.setup_status_label.config(text="Installing packages...") if hasattr(self, 'setup_status_label') else None)
                    elif 'Successfully installed' in line:
                        self.root.after(0, lambda: self.setup_status_label.config(text="Installation complete!") if hasattr(self, 'setup_status_label') else None)
                
                process.wait()
                
                if process.returncode == 0:
                    self.log("\n✅ Whisper installed successfully!")
                    self.root.after(0, self._on_whisper_installed)
                else:
                    self.log(f"\n❌ Installation failed (code {process.returncode})")
                    self.root.after(0, lambda: self._on_install_failed("Whisper"))
            except Exception as e:
                self.log(f"\n❌ Error: {e}")
                self.root.after(0, lambda: self._on_install_failed("Whisper"))
        
        threading.Thread(target=install, daemon=True).start()
    
    def _on_install_failed(self, package_name):
        """Handle installation failure"""
        if hasattr(self, 'whisper_install_btn'):
            self.whisper_install_btn.config(state='normal', text=f'🔄 Retry Install {package_name}')
        if hasattr(self, 'setup_status_label'):
            self.setup_status_label.config(text=f"Installation failed. Check the log for details.")
    
    def _on_whisper_installed(self):
        """Called after Whisper is successfully installed"""
        # Remove setup frame
        if hasattr(self, 'setup_frame'):
            self.setup_frame.destroy()
            delattr(self, 'setup_frame')
        
        # Show all sections again
        if hasattr(self, 'files_section'):
            self.files_section.grid()
            self.files_section.expand()
        if hasattr(self, 'transcription_section'):
            self.transcription_section.grid()
            self.transcription_section.expand()
        if hasattr(self, 'model_section'):
            self.model_section.grid()
        if hasattr(self, 'subtitle_section'):
            self.subtitle_section.grid()
        if hasattr(self, 'ai_section'):
            self.ai_section.grid()
        
        # Re-enable start button
        self.process_btn.config(state='normal')
        
        self.log("\n✅ Whisper is ready!")
        self.log("")
        
        # Now check GPU/PyTorch
        self.check_gpu_availability()
    
    def check_gpu_availability(self):
        """Check if GPU is available for Whisper"""
        def check():
            gpu_ready = False
            
            # For frozen EXE, use subprocess to avoid Python version conflicts
            if getattr(sys, 'frozen', False):
                self.log("Checking GPU availability...")
                try:
                    result = _run_hidden(
                        ['py', '-c', '''
import torch
if torch.cuda.is_available():
    print("GPU:" + torch.cuda.get_device_name(0))
else:
    print("CPU_ONLY")
'''],
                        capture_output=True, text=True, timeout=60
                    )
                    
                    output = result.stdout.strip()
                    
                    if result.returncode == 0 and output.startswith("GPU:"):
                        gpu_name = output[4:]  # Remove "GPU:" prefix
                        self.device_info.set(f"✅ GPU: {gpu_name}")
                        self.use_gpu.set(True)
                        self.log(f"✅ GPU detected: {gpu_name}")
                        gpu_ready = True
                    elif "CPU_ONLY" in output:
                        self.device_info.set("⚠️ CPU-only PyTorch")
                        self.use_gpu.set(False)
                        self.log("⚠️ PyTorch installed but without CUDA support")
                        if self.check_nvidia_gpu():
                            self.root.after(0, self._show_gpu_install_button)
                    else:
                        # PyTorch not installed or error
                        self.device_info.set("⚠️ PyTorch not installed")
                        self.use_gpu.set(False)
                        self.log("⚠️ PyTorch not found")
                        # Show PyTorch install button
                        self.root.after(0, self._show_pytorch_install_button)
                except subprocess.TimeoutExpired:
                    self.device_info.set("⚠️ GPU check timed out")
                    self.use_gpu.set(False)
                    self.log("⚠️ GPU check timed out")
                except Exception as e:
                    self.device_info.set("⚠️ GPU check failed")
                    self.use_gpu.set(False)
                    self.log(f"⚠️ GPU check error: {e}")
            else:
                # Running as script - can import directly
                try:
                    self.log("Checking GPU availability...")
                    import torch
                    
                    if torch.cuda.is_available():
                        gpu_name = torch.cuda.get_device_name(0)
                        self.device_info.set(f"✅ GPU: {gpu_name}")
                        self.use_gpu.set(True)
                        self.log(f"✅ GPU detected: {gpu_name}")
                        gpu_ready = True
                    else:
                        self.device_info.set("⚠️ CPU-only PyTorch")
                        self.use_gpu.set(False)
                        self.log("⚠️ PyTorch installed but without CUDA support")
                        if self.check_nvidia_gpu():
                            self.offer_gpu_setup()
                except ImportError:
                    self.device_info.set("⚠️ PyTorch not installed")
                    self.use_gpu.set(False)
                    self.log("⚠️ PyTorch not found")
                    if self.check_nvidia_gpu():
                        self.offer_gpu_setup()
                except Exception as e:
                    self.device_info.set("⚠️ GPU check failed")
                    self.use_gpu.set(False)
                    self.log(f"⚠️ GPU check error: {e}")
            
            # Update model status after GPU check completes
            if hasattr(self, 'model_section'):
                self.root.after(150, self.update_model_status)
        
        threading.Thread(target=check, daemon=True).start()
    
    def _show_pytorch_install_button(self):
        """Show button to install PyTorch at the top of the UI"""
        # Check if NVIDIA GPU is present
        has_nvidia = self.check_nvidia_gpu()
        
        # Create install frame at the TOP (row 0)
        if hasattr(self, 'pytorch_frame'):
            self.pytorch_frame.destroy()
        
        self.pytorch_frame = ttk.Frame(self.left_panel)
        self.pytorch_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # Shift all sections down by changing their row
        if hasattr(self, 'files_section'):
            self.files_section.grid(row=1)
        if hasattr(self, 'transcription_section'):
            self.transcription_section.grid(row=2)
        if hasattr(self, 'model_section'):
            self.model_section.grid(row=3)
        if hasattr(self, 'subtitle_section'):
            self.subtitle_section.grid(row=4)
        if hasattr(self, 'ai_section'):
            self.ai_section.grid(row=5)
        
        if has_nvidia:
            self.log("\n🎮 NVIDIA GPU detected!")
            
            # Title
            title_label = ttk.Label(
                self.pytorch_frame,
                text="⚡ GPU Acceleration Available",
                font=('Segoe UI', 11, 'bold')
            )
            title_label.pack(pady=(0, 5))
            
            info_label = ttk.Label(
                self.pytorch_frame,
                text="Install PyTorch with CUDA for ~10x faster transcription",
                wraplength=400
            )
            info_label.pack(pady=(0, 10))
            
            self.pytorch_install_btn = ttk.Button(
                self.pytorch_frame,
                text="⚡ Install PyTorch with GPU Support",
                command=self._install_pytorch_cuda
            )
            self.pytorch_install_btn.pack(ipadx=15, ipady=8)
            
            # Skip option
            skip_btn = ttk.Button(
                self.pytorch_frame,
                text="Skip (use CPU instead - slower)",
                command=self._install_pytorch_cpu
            )
            skip_btn.pack(pady=(8, 0))
        else:
            self.log("\n💻 No NVIDIA GPU detected - will use CPU mode.")
            
            title_label = ttk.Label(
                self.pytorch_frame,
                text="📦 PyTorch Required",
                font=('Segoe UI', 11, 'bold')
            )
            title_label.pack(pady=(0, 5))
            
            info_label = ttk.Label(
                self.pytorch_frame,
                text="PyTorch is required for transcription",
                wraplength=400
            )
            info_label.pack(pady=(0, 10))
            
            self.pytorch_install_btn = ttk.Button(
                self.pytorch_frame,
                text="📥 Install PyTorch",
                command=self._install_pytorch_cpu
            )
            self.pytorch_install_btn.pack(ipadx=15, ipady=8)
    
    def _show_gpu_install_button(self):
        """Show button to upgrade PyTorch to CUDA version at top of UI"""
        # Create frame at the TOP (row 0)
        if hasattr(self, 'pytorch_frame'):
            self.pytorch_frame.destroy()
        
        self.pytorch_frame = ttk.Frame(self.left_panel)
        self.pytorch_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # Shift all sections down
        if hasattr(self, 'files_section'):
            self.files_section.grid(row=1)
        if hasattr(self, 'transcription_section'):
            self.transcription_section.grid(row=2)
        if hasattr(self, 'model_section'):
            self.model_section.grid(row=3)
        if hasattr(self, 'subtitle_section'):
            self.subtitle_section.grid(row=4)
        if hasattr(self, 'ai_section'):
            self.ai_section.grid(row=5)
        
        # Title
        title_label = ttk.Label(
            self.pytorch_frame,
            text="⚡ GPU Acceleration Available",
            font=('Segoe UI', 11, 'bold')
        )
        title_label.pack(pady=(0, 5))
        
        info_label = ttk.Label(
            self.pytorch_frame,
            text="Upgrade PyTorch for ~10x faster transcription (optional)",
            wraplength=400
        )
        info_label.pack(pady=(0, 10))
        
        self.pytorch_install_btn = ttk.Button(
            self.pytorch_frame,
            text="⚡ Upgrade to GPU-accelerated PyTorch",
            command=self._install_pytorch_cuda
        )
        self.pytorch_install_btn.pack(ipadx=15, ipady=8)
        
        # Add dismiss button
        dismiss_btn = ttk.Button(
            self.pytorch_frame,
            text="Keep CPU version",
            command=self._dismiss_pytorch_upgrade
        )
        dismiss_btn.pack(pady=(8, 0))
    
    def _dismiss_pytorch_upgrade(self):
        """Dismiss the PyTorch upgrade prompt"""
        if hasattr(self, 'pytorch_frame'):
            self.pytorch_frame.destroy()
            delattr(self, 'pytorch_frame')
        
        # Restore section positions to original rows
        if hasattr(self, 'files_section'):
            self.files_section.grid(row=0)
        if hasattr(self, 'transcription_section'):
            self.transcription_section.grid(row=1)
        if hasattr(self, 'model_section'):
            self.model_section.grid(row=2)
        if hasattr(self, 'subtitle_section'):
            self.subtitle_section.grid(row=3)
        if hasattr(self, 'ai_section'):
            self.ai_section.grid(row=4)
        
        self.log("ℹ️ Using CPU mode for transcription.")
    
    def _install_pytorch_cpu(self):
        """Install CPU-only PyTorch"""
        if hasattr(self, 'pytorch_install_btn'):
            self.pytorch_install_btn.config(state='disabled', text='⏳ Installing PyTorch...')
        
        # Add a status label to the pytorch frame if not exists
        if hasattr(self, 'pytorch_frame') and not hasattr(self, 'pytorch_status_label'):
            self.pytorch_status_label = ttk.Label(self.pytorch_frame, text="", foreground='gray')
            self.pytorch_status_label.pack(pady=(10, 0))
        
        def update_status(text):
            if hasattr(self, 'pytorch_status_label'):
                self.root.after(0, lambda: self.pytorch_status_label.config(text=text))
        
        def install():
            self.log("\n" + "="*60)
            self.log("📦 Installing PyTorch (CPU version)...")
            self.log("="*60 + "\n")
            
            update_status("Starting installation...")
            
            try:
                process = subprocess.Popen(
                    ['py', '-m', 'pip', 'install', '--user', '--progress-bar', 'off', 'torch', 'torchaudio'],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding='utf-8',
                    errors='replace',
                    bufsize=1,
                    creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                )
                
                current_package = ""
                for line in process.stdout:
                    line = line.rstrip()
                    if not line:
                        continue
                    if line.startswith('  ') and 'Uninstalling' not in line and 'Installing' not in line:
                        continue
                    
                    self.log(line)
                    
                    if 'Collecting' in line:
                        pkg = line.replace('Collecting ', '').split()[0]
                        current_package = pkg
                        update_status(f"Collecting {pkg}...")
                    elif 'Downloading' in line:
                        import re
                        size_match = re.search(r'\(([^)]+)\)', line)
                        size_info = f" ({size_match.group(1)})" if size_match else ""
                        update_status(f"Downloading {current_package}{size_info}...")
                    elif 'Installing collected packages' in line:
                        update_status("Installing packages...")
                    elif 'Successfully installed' in line:
                        update_status("Installation complete!")
                
                process.wait()
                
                if process.returncode == 0:
                    self.log("\n✅ PyTorch installed successfully!")
                    self.root.after(0, self._on_pytorch_installed)
                else:
                    self.log(f"\n❌ Installation failed (code {process.returncode})")
                    self.root.after(0, lambda: self.pytorch_install_btn.config(
                        state='normal', text='🔄 Retry Install'))
            except Exception as e:
                self.log(f"\n❌ Error: {e}")
                self.root.after(0, lambda: self.pytorch_install_btn.config(
                    state='normal', text='🔄 Retry Install'))
        
        threading.Thread(target=install, daemon=True).start()
    
    def _install_pytorch_cuda(self):
        """Install CUDA-enabled PyTorch"""
        if hasattr(self, 'pytorch_install_btn'):
            self.pytorch_install_btn.config(state='disabled', text='⏳ Installing PyTorch with CUDA...')
        
        # Add a status label to the pytorch frame if not exists
        if hasattr(self, 'pytorch_frame') and not hasattr(self, 'pytorch_status_label'):
            self.pytorch_status_label = ttk.Label(self.pytorch_frame, text="", foreground='gray')
            self.pytorch_status_label.pack(pady=(10, 0))
        
        def update_status(text):
            if hasattr(self, 'pytorch_status_label'):
                self.root.after(0, lambda: self.pytorch_status_label.config(text=text))
        
        def install():
            self.log("\n" + "="*60)
            self.log("📦 Installing PyTorch with CUDA support...")
            self.log("   This may take several minutes (PyTorch is ~2GB)")
            self.log("="*60 + "\n")
            
            update_status("Starting installation... (this may take a few minutes)")
            
            try:
                # Install PyTorch with CUDA 12.1 support
                process = subprocess.Popen(
                    ['py', '-m', 'pip', 'install', '--user', '--progress-bar', 'off', 'torch', 'torchaudio', 
                     '--index-url', 'https://download.pytorch.org/whl/cu121'],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding='utf-8',
                    errors='replace',
                    bufsize=1,
                    creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                )
                
                current_package = ""
                for line in process.stdout:
                    line = line.rstrip()
                    if not line:
                        continue
                    if line.startswith('  ') and 'Uninstalling' not in line and 'Installing' not in line:
                        continue
                    
                    self.log(line)
                    
                    if 'Collecting' in line:
                        pkg = line.replace('Collecting ', '').split()[0]
                        current_package = pkg
                        update_status(f"Collecting {pkg}...")
                    elif 'Downloading' in line:
                        import re
                        size_match = re.search(r'\(([^)]+)\)', line)
                        size_info = f" ({size_match.group(1)})" if size_match else ""
                        update_status(f"Downloading {current_package}{size_info}...")
                    elif 'Installing collected packages' in line:
                        update_status("Installing packages...")
                    elif 'Successfully installed' in line:
                        update_status("Installation complete!")
                
                process.wait()
                
                if process.returncode == 0:
                    self.log("\n✅ PyTorch with CUDA installed successfully!")
                    self.root.after(0, self._on_pytorch_installed)
                else:
                    self.log(f"\n❌ Installation failed (code {process.returncode})")
                    self.root.after(0, lambda: self.pytorch_install_btn.config(
                        state='normal', text='🔄 Retry Install'))
            except Exception as e:
                self.log(f"\n❌ Error: {e}")
                self.root.after(0, lambda: self.pytorch_install_btn.config(
                    state='normal', text='🔄 Retry Install'))
        
        threading.Thread(target=install, daemon=True).start()
    
    def _on_pytorch_installed(self):
        """Called after PyTorch is successfully installed"""
        # Remove install frame
        if hasattr(self, 'pytorch_frame'):
            self.pytorch_frame.destroy()
            delattr(self, 'pytorch_frame')
        
        # Restore section positions to original rows
        if hasattr(self, 'files_section'):
            self.files_section.grid(row=0)
        if hasattr(self, 'transcription_section'):
            self.transcription_section.grid(row=1)
        if hasattr(self, 'model_section'):
            self.model_section.grid(row=2)
        if hasattr(self, 'subtitle_section'):
            self.subtitle_section.grid(row=3)
        if hasattr(self, 'ai_section'):
            self.ai_section.grid(row=4)
        
        self.log("")
        
        # Re-check GPU availability to update status
        self.check_gpu_availability()
    
    def check_nvidia_gpu(self):
        """Check if NVIDIA GPU is present using nvidia-smi"""
        try:
            # Try common nvidia-smi locations
            nvidia_smi_paths = [
                'nvidia-smi',  # In PATH
                r'C:\Windows\System32\nvidia-smi.exe',
                r'C:\Program Files\NVIDIA Corporation\NVSMI\nvidia-smi.exe',
            ]

            for nvidia_smi in nvidia_smi_paths:
                try:
                    result = subprocess.run([nvidia_smi], capture_output=True, timeout=5,
                                          creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0)
                    if result.returncode == 0:
                        return True
                except:
                    continue
            return False
        except:
            return False
    
    def offer_gpu_setup(self):
        """Offer to install CUDA-enabled PyTorch"""
        # Check if user previously declined, already completed setup, or recently attempted
        config = self.load_config()
        if config.get('gpu_setup_declined', False):
            return  # User previously said no
        if config.get('gpu_setup_completed', False):
            return  # Setup already done

        # Check if we attempted setup recently (within last 5 minutes)
        import time
        last_attempt = config.get('gpu_setup_last_attempt', 0)
        if time.time() - last_attempt < 300:  # 5 minutes
            self.log("⏳ GPU setup was recently attempted - skipping popup")
            return  # Don't spam the user

        response = messagebox.askyesno(
            "GPU Setup Available",
            "An NVIDIA GPU was detected but PyTorch with CUDA support is not installed.\n\n"
            "Would you like to automatically install GPU-accelerated PyTorch?\n\n"
            "This will make transcription ~10x faster!\n\n"
            "Installation may take a few minutes."
        )

        if response:
            # Mark attempt time to prevent popup loop
            config['gpu_setup_last_attempt'] = time.time()
            self.save_config(config)
            self.install_cuda_pytorch()
        else:
            # User declined - remember this choice
            config['gpu_setup_declined'] = True
            self.save_config(config)
    
    def install_cuda_pytorch(self):
        """Install CUDA-enabled PyTorch"""
        # Check for admin privileges
        if not is_admin():
            response = messagebox.askyesno(
                "Administrator Access Required",
                "Installing PyTorch requires administrator privileges.\n\n"
                "Would you like to restart this application as administrator?\n\n"
                "(The app will close and reopen with elevated privileges)"
            )

            if response:
                self.log("🔐 Requesting administrator privileges...")
                if run_as_admin(install_gpu=True):
                    # Successfully requested elevation, close this instance
                    self.log("✓ UAC prompt should appear - please approve")
                    self.log("This window will close and reopen with admin rights...")
                    self.root.after(1000, self.root.quit)  # Give time to read the message
                    return
                else:
                    self.log("❌ Failed to request administrator elevation")
                    messagebox.showerror("Elevation Failed",
                        "Could not restart as administrator.\n\n"
                        "This might happen if:\n"
                        "• UAC is disabled\n"
                        "• The EXE path contains special characters\n"
                        "• Windows security settings block elevation\n\n"
                        "Try: Right-click Final Whisper.exe → 'Run as administrator'")
                    return
            else:
                # Show manual installation instructions
                self.show_manual_install_instructions()
                return

        # If we get here, we have admin rights - proceed with installation
        self.install_cuda_pytorch_direct()

    def install_cuda_pytorch_direct(self):
        """Actually perform the PyTorch installation (called after admin check)"""
        # Prevent multiple simultaneous installations
        if hasattr(self, '_installing_gpu') and self._installing_gpu:
            self.log("⚠️ GPU installation already in progress")
            return

        self._installing_gpu = True

        # Get the Python executable (important for frozen EXE)
        python_exe = get_python_executable()
        if not python_exe:
            self.log("❌ ERROR: Could not find Python executable!")
            self.log("Please ensure Python is installed and in your PATH.")
            messagebox.showerror("Python Not Found",
                "Could not find Python executable.\n\n"
                "Please ensure Python is installed and in your system PATH.")
            self._installing_gpu = False
            return

        self.log("\n" + "="*60)
        self.log("Installing GPU-accelerated PyTorch...")
        self.log("="*60 + "\n")
        self.log(f"Admin status: {'Yes' if is_admin() else 'No'}")
        self.log(f"Python executable: {python_exe}\n")

        def install():
            try:
                self.root.after(0, lambda: self.log("Installation thread started..."))
                # Detect CUDA version
                self.log("Detecting CUDA version...")
                cuda_version = None
                
                try:
                    result = subprocess.run(['nvidia-smi'], capture_output=True, text=True, timeout=5)
                    if result.returncode == 0:
                        # Parse CUDA version from nvidia-smi output
                        for line in result.stdout.split('\n'):
                            if 'CUDA Version' in line:
                                import re
                                match = re.search(r'CUDA Version:\s*(\d+\.\d+)', line)
                                if match:
                                    cuda_version = match.group(1)
                                    break
                except:
                    pass
                
                if cuda_version:
                    self.log(f"Found CUDA {cuda_version}")
                    
                    # Determine which PyTorch version to install
                    cuda_major = int(cuda_version.split('.')[0])
                    
                    if cuda_major >= 12:
                        torch_index = "https://download.pytorch.org/whl/cu121"
                        self.log("Installing PyTorch with CUDA 12.1 support...")
                    elif cuda_major >= 11:
                        torch_index = "https://download.pytorch.org/whl/cu118"
                        self.log("Installing PyTorch with CUDA 11.8 support...")
                    else:
                        self.log(f"⚠️ CUDA {cuda_version} is too old. Need CUDA 11.x or newer.")
                        messagebox.showerror("CUDA Too Old", 
                            f"Your CUDA version ({cuda_version}) is too old.\n\n"
                            "Please update your NVIDIA drivers.")
                        return
                else:
                    # Default to CUDA 11.8 if we can't detect
                    torch_index = "https://download.pytorch.org/whl/cu118"
                    self.log("Could not detect CUDA version, defaulting to CUDA 11.8...")
                
                # Uninstall existing PyTorch
                self.log("\nUninstalling CPU-only PyTorch...")
                uninstall_result = subprocess.run(
                    [python_exe, "-m", "pip", "uninstall", "-y", "torch", "torchvision", "torchaudio"],
                    capture_output=True, text=True
                )
                if uninstall_result.returncode == 0:
                    self.log("✓ Uninstall complete")
                else:
                    self.log(f"Note: Uninstall returned code {uninstall_result.returncode}")

                # Install CUDA PyTorch
                self.log(f"\nInstalling PyTorch with {torch_index.split('/')[-1]}...")
                self.log("This will take several minutes - downloading ~2GB...\n")

                install_cmd = [
                    python_exe, "-m", "pip", "install",
                    "torch", "torchvision", "torchaudio",
                    "--index-url", torch_index
                ]

                self.log(f"Running: {' '.join(install_cmd)}\n")

                process = subprocess.Popen(
                    install_cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1,
                    universal_newlines=True
                )

                # Stream output - collect and log periodically
                output_lines = []
                try:
                    while True:
                        line = process.stdout.readline()
                        if not line:
                            break
                        line = line.rstrip()
                        if line:
                            output_lines.append(line)
                            # Log every 10 lines or important messages
                            if len(output_lines) % 10 == 0 or 'Error' in line or 'Successfully' in line:
                                self.root.after(0, lambda lines=output_lines[:]: [self.log(l) for l in lines])
                                output_lines = []
                except Exception as e:
                    self.root.after(0, lambda: self.log(f"Error reading pip output: {e}"))

                process.wait()

                # Log any remaining lines
                if output_lines:
                    self.root.after(0, lambda lines=output_lines[:]: [self.log(l) for l in lines])

                self.root.after(0, lambda code=process.returncode: self.log(f"\nPip install finished with code: {code}"))
                
                if process.returncode == 0:
                    self.log("\n✅ GPU-accelerated PyTorch installed successfully!")
                    self.log("\nRestarting application...")

                    # Mark that GPU setup was completed successfully
                    config = self.load_config()
                    config['gpu_setup_completed'] = True
                    self.save_config(config)

                    # Show success message and restart
                    self.root.after(0, lambda: messagebox.showinfo("Success",
                        "GPU-accelerated PyTorch installed successfully!\n\n"
                        "The application will now restart to enable GPU acceleration."))

                    # Restart the application
                    self.root.after(100, self.restart_application)
                else:
                    self.log("\n❌ Installation failed!")
                    self.log(f"Exit code: {process.returncode}")

                    # Show detailed error with manual installation option
                    self.root.after(0, lambda: self.show_manual_install_instructions(failed=True))

            except Exception as e:
                self.root.after(0, lambda: self.log(f"\n❌ CRITICAL ERROR during installation: {str(e)}"))
                import traceback
                tb = traceback.format_exc()
                self.root.after(0, lambda: self.log(tb))
                self.root.after(0, lambda: self.show_manual_install_instructions(failed=True, error=str(e)))
            finally:
                # Always clear the installation flag
                self._installing_gpu = False
                self.root.after(0, lambda: self.log("Installation thread finished."))

        threading.Thread(target=install, daemon=True).start()

    def show_manual_install_instructions(self, failed=False, error=None):
        """Show manual installation instructions"""
        # Detect CUDA version for the command
        cuda_version = None
        try:
            result = subprocess.run(['nvidia-smi'], capture_output=True, text=True, timeout=5,
                                  creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0)
            if result.returncode == 0:
                import re
                for line in result.stdout.split('\n'):
                    if 'CUDA Version' in line:
                        match = re.search(r'CUDA Version:\s*(\d+\.\d+)', line)
                        if match:
                            cuda_version = match.group(1)
                            break
        except:
            pass

        # Determine PyTorch index URL
        if cuda_version:
            cuda_major = int(cuda_version.split('.')[0])
            if cuda_major >= 12:
                torch_index = "https://download.pytorch.org/whl/cu121"
                cuda_label = "CUDA 12.1"
            else:
                torch_index = "https://download.pytorch.org/whl/cu118"
                cuda_label = "CUDA 11.8"
        else:
            torch_index = "https://download.pytorch.org/whl/cu118"
            cuda_label = "CUDA 11.8"

        install_cmd = f'pip install torch torchvision torchaudio --index-url {torch_index}'

        if failed:
            title = "Installation Failed"
            message = f"Automatic installation failed.\n\n"
            if error:
                message += f"Error: {error}\n\n"
            message += "Please try installing manually:\n\n"
        else:
            title = "Manual Installation"
            message = "To install GPU-accelerated PyTorch manually:\n\n"

        message += f"1. Open Command Prompt as Administrator\n"
        message += f"2. Run this command:\n\n"
        message += f"{install_cmd}\n\n"
        message += f"3. Restart Final Whisper\n\n"
        message += f"Would you like to copy the command to clipboard?"

        response = messagebox.askyesno(title, message)

        if response:
            # Copy command to clipboard
            try:
                self.root.clipboard_clear()
                self.root.clipboard_append(install_cmd)
                self.root.update()
                messagebox.showinfo("Copied",
                    "Installation command copied to clipboard!\n\n"
                    "Paste it into Command Prompt (run as Administrator).")
                self.log(f"\n📋 Command copied to clipboard: {install_cmd}")
            except Exception as e:
                messagebox.showerror("Copy Failed", f"Could not copy to clipboard:\n{e}")

    def manual_gpu_setup(self):
        """Manually trigger GPU setup"""
        if self.check_nvidia_gpu():
            try:
                import torch
                if torch.cuda.is_available():
                    messagebox.showinfo("GPU Already Working",
                        f"GPU is already configured and working!\n\n"
                        f"Detected: {torch.cuda.get_device_name(0)}")
                    return
            except ImportError:
                pass

            # Clear the preferences since user is manually requesting setup
            config = self.load_config()
            if 'gpu_setup_declined' in config:
                del config['gpu_setup_declined']
            if 'gpu_setup_completed' in config:
                del config['gpu_setup_completed']
            self.save_config(config)

            self.install_cuda_pytorch()
        else:
            messagebox.showwarning("No NVIDIA GPU",
                "No NVIDIA GPU detected on this system.\n\n"
                "GPU acceleration requires an NVIDIA graphics card with CUDA support.")
        
    def browse_video(self):
        """Browse for video file(s) - supports multiple selection"""
        filenames = filedialog.askopenfilenames(
            title="Select Video/Audio File(s)",
            filetypes=[
                ("Video files", "*.mp4 *.mkv *.avi *.mov *.m4v *.webm *.mp3 *.wav *.m4a *.flac *.ogg"),
                ("All files", "*.*")
            ]
        )
        if filenames:
            if len(filenames) == 1:
                # Single file
                self.video_path.set(filenames[0])
                self.selected_files = None  # Clear batch list
                input_dir = str(Path(filenames[0]).parent)
                self.output_dir.set(input_dir)
                self._hide_batch_context_controls()
            else:
                # Multiple files
                self.selected_files = list(filenames)
                self.video_path.set(f"{len(filenames)} files selected")
                # Set output folder to same directory as first file
                input_dir = str(Path(filenames[0]).parent)
                self.output_dir.set(input_dir)
                self.log(f"📁 Selected {len(filenames)} files for batch processing")
                self._show_batch_context_controls()
    
    def browse_input_folder(self):
        """Browse for input folder (batch mode)"""
        dirname = filedialog.askdirectory(title="Select Folder with Video/Audio Files")
        if dirname:
            # Get video files in folder
            video_extensions = {'.mp4', '.mkv', '.avi', '.mov', '.m4v', '.webm', '.mp3', '.wav', '.m4a', '.flac', '.ogg'}
            files = sorted([f for f in Path(dirname).iterdir() if f.suffix.lower() in video_extensions])
            
            if files:
                self.selected_files = [str(f) for f in files]
                self.video_path.set(f"{len(files)} files in folder")
                self.output_dir.set(dirname)
                self.log(f"📁 Found {len(files)} video/audio files in folder")
                self._show_batch_context_controls()
            else:
                self.video_path.set(dirname)
                self.selected_files = None
                self.output_dir.set(dirname)
                self.log("⚠️ No video/audio files found in folder")
                self._hide_batch_context_controls()
            
    def browse_output(self):
        """Browse for output directory"""
        dirname = filedialog.askdirectory(title="Select Output Folder")
        if dirname:
            self.output_dir.set(dirname)
    
    def _show_batch_context_controls(self):
        """Show the batch context controls when multiple files selected"""
        if not self.selected_files:
            return
        
        # Initialize prompts for each file (copy current prompt to all)
        current_prompt = self.context_entry.get("1.0", tk.END).strip()
        self._batch_prompts = [current_prompt] * len(self.selected_files)
        self._current_prompt_index = 0
        
        # Show the batch controls
        self.batch_context_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(4, 0))
        
        # Set initial state based on checkbox (same prompt is default)
        self._toggle_same_prompt()
    
    def _hide_batch_context_controls(self):
        """Hide the batch context controls"""
        self.batch_context_frame.grid_forget()
        self._batch_prompts = []
        self._current_prompt_index = 0
    
    def _update_batch_file_indicator(self):
        """Update the file indicator label"""
        if not self.selected_files:
            return
        
        idx = self._current_prompt_index
        total = len(self.selected_files)
        filename = Path(self.selected_files[idx]).name
        
        # Truncate filename if too long
        if len(filename) > 35:
            filename = filename[:32] + "..."
        
        self.batch_file_indicator.config(text=f"File {idx + 1}/{total}: {filename}")
        
        # Update button states
        self.prev_file_btn.config(state='normal' if idx > 0 else 'disabled')
        self.next_file_btn.config(state='normal' if idx < total - 1 else 'disabled')
    
    def _prev_batch_file(self):
        """Navigate to previous file's prompt"""
        if self._current_prompt_index > 0:
            # Save current prompt
            self._batch_prompts[self._current_prompt_index] = self.context_entry.get("1.0", tk.END).strip()
            
            # Move to previous
            self._current_prompt_index -= 1
            
            # Load prompt for this file
            self.context_entry.delete("1.0", tk.END)
            self.context_entry.insert("1.0", self._batch_prompts[self._current_prompt_index])
            
            self._update_batch_file_indicator()
    
    def _next_batch_file(self):
        """Navigate to next file's prompt"""
        if self._current_prompt_index < len(self.selected_files) - 1:
            # Save current prompt
            self._batch_prompts[self._current_prompt_index] = self.context_entry.get("1.0", tk.END).strip()
            
            # Move to next
            self._current_prompt_index += 1
            
            # Load prompt for this file
            self.context_entry.delete("1.0", tk.END)
            self.context_entry.insert("1.0", self._batch_prompts[self._current_prompt_index])
            
            self._update_batch_file_indicator()
    
    def _toggle_same_prompt(self):
        """Toggle between same prompt for all files or individual prompts"""
        if self.use_same_prompt.get():
            # Copy current prompt to all files
            current_prompt = self.context_entry.get("1.0", tk.END).strip()
            self._batch_prompts = [current_prompt] * len(self.selected_files)
            # Disable navigation
            self.prev_file_btn.config(state='disabled')
            self.next_file_btn.config(state='disabled')
            self.batch_file_indicator.config(text="Same prompt for all files")
        else:
            # Re-enable navigation
            self._update_batch_file_indicator()
    
    def get_prompt_for_file(self, file_index):
        """Get the context prompt for a specific file"""
        if self.use_same_prompt.get() or not self._batch_prompts:
            return self.context_entry.get("1.0", tk.END).strip()
        
        if 0 <= file_index < len(self._batch_prompts):
            return self._batch_prompts[file_index]
        return ""
    
    def get_batch_files(self):
        """Get list of files to process"""
        if hasattr(self, 'selected_files') and self.selected_files:
            return [Path(f) for f in self.selected_files]
        return []
            
    def start_transcription(self):
        """Start the transcription process"""
        if self.processing:
            messagebox.showwarning("Processing", "A transcription is already in progress")
            return
        
        # Check if we have multiple files selected
        batch_files = self.get_batch_files()
        
        if batch_files:
            # Batch mode - multiple files or folder
            self.batch_files = batch_files
            self.batch_index = 0
            self.batch_total = len(batch_files)
            self.log(f"\n{'='*60}")
            self.log(f"🎬 Batch Mode: {self.batch_total} files to process")
            self.log(f"{'='*60}\n")
        else:
            # Single file validation
            if not self.video_path.get():
                messagebox.showerror("Error", "Please select a video file")
                return
            if not os.path.exists(self.video_path.get()):
                messagebox.showerror("Error", "Video file does not exist")
                return
            self.batch_files = None
            self.batch_index = 0
            self.batch_total = 1
            # Show filename for single file
            filename = Path(self.video_path.get()).name
            self.set_batch_file_text("📄 ", filename)
            
        # Create output directory
        os.makedirs(self.output_dir.get(), exist_ok=True)
        
        self.processing = True
        self.stop_requested = False
        self.process_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        self.progress['value'] = 0
        self.progress.start_animation()  # Start gradient animation
        self.set_progress_text("Starting...")
        
        # Start timer
        import time
        self.transcription_start_time = time.time()
        self.timer_running = True
        self.last_progress_value = None
        self.last_audio_position = None
        self.last_audio_duration = None
        self.update_timer()
        
        # Run in thread
        thread = threading.Thread(target=self.run_transcription, daemon=True)
        thread.start()
    
    def stop_transcription(self):
        """Stop the transcription process"""
        if self.processing:
            self.stop_requested = True
            self.stop_btn.config(state='disabled')
            self.set_progress_text("Stopping...")
            self.log("\n⚠️ Stop requested - cancelling transcription...")
            
            # Also terminate the subprocess if it's running
            if self.transcription_process is not None:
                try:
                    self.transcription_process.terminate()
                except:
                    pass
    
    def update_timer(self):
        """Update the progress label with elapsed time every second"""
        if not self.timer_running or not self.transcription_start_time:
            return
        
        # Update the progress label
        try:
            self._update_progress_label()
        except Exception as e:
            pass
        
        # Schedule next update
        if self.timer_running:
            self.root.after(1000, self.update_timer)
    
    def _update_progress_label(self):
        """Update progress label with current stats - only updates elapsed time"""
        import time
        
        if not self.transcription_start_time:
            return
        
        # Don't overwrite the "Done" message
        current_text = self.get_progress_text()
        if '✓ Done' in current_text or 'Cancelled' in current_text:
            self.timer_running = False
            return
        
        time_elapsed = time.time() - self.transcription_start_time
        elapsed_min = int(time_elapsed // 60)
        elapsed_sec = int(time_elapsed % 60)
        
        # Check if we're in proofreading mode - always update in this mode
        if getattr(self, '_proofreading_mode', False):
            batch_info = getattr(self, '_proofreading_batch_info', None)
            if batch_info:
                self.set_progress_text(f"AI Proofreading... {batch_info} | Elapsed: {elapsed_min}:{elapsed_sec:02d}")
            else:
                self.set_progress_text(f"AI Proofreading... | Elapsed: {elapsed_min}:{elapsed_sec:02d}")
            return
        
        # Don't update if detailed progress was recently set (within last 2 seconds)
        # This allows the transcription progress updates to control the display
        last_detail_update = getattr(self, '_last_detail_progress_time', 0)
        if time.time() - last_detail_update < 2.0:
            return
        
        # If we have ETA/Speed info, update the whole string with new elapsed time
        if 'ETA:' in current_text and 'Speed:' in current_text:
            # Parse and rebuild the detailed progress string
            try:
                import re
                # Extract parts: "Transcribing... X% (Ys / Zs) | Elapsed: ... | ETA: ... | Speed: ..."
                match = re.match(r'(.*?\(\d+s / \d+s\)).*ETA: (\d+:\d+).*Speed: ([\d.]+x)', current_text)
                if match:
                    base = match.group(1)
                    eta = match.group(2)
                    speed = match.group(3)
                    self.set_progress_text(f"{base} | Elapsed: {elapsed_min}:{elapsed_sec:02d} | ETA: {eta} | Speed: {speed}")
                    return
            except:
                pass
        
        # Simple format - just update elapsed time
        if 'Elapsed:' in current_text:
            base_text = current_text.split('  |  ')[0]  # Split on colored separator
            self.set_progress_text(f"{base_text} | Elapsed: {elapsed_min}:{elapsed_sec:02d}")
    
    def update_progress_with_time(self, value, status_text):
        """Update progress bar and label with elapsed time"""
        import time
        self.progress['value'] = value
        
        if self.transcription_start_time:
            elapsed = time.time() - self.transcription_start_time
            minutes = int(elapsed // 60)
            seconds = int(elapsed % 60)
            self.set_progress_text(f"{status_text} | Elapsed: {minutes}:{seconds:02d}")
        else:
            self.set_progress_text(status_text)
        
    def run_transcription(self):
        """Run the actual transcription"""
        try:
            # Batch mode - process multiple files
            if self.batch_files:
                for i, file_path in enumerate(self.batch_files):
                    if self.stop_requested:
                        self.root.after(0, lambda: self.log(f"\n❌ Batch processing stopped ({i}/{self.batch_total} completed)"))
                        break
                    
                    self.batch_index = i + 1
                    self._current_file_index = i  # Store for per-file context
                    self.root.after(0, lambda idx=i+1, total=self.batch_total, name=file_path.name: 
                                   self.log(f"\n📄 Processing file {idx}/{total}: {name}"))
                    
                    # Reset progress and update batch file label
                    def reset_for_next_file(idx=i+1, total=self.batch_total, name=file_path.name):
                        self.progress['value'] = 0
                        self.progress.start_animation()
                        self.set_batch_file_text(f"📄 File {idx}/{total}: ", name)
                    self.root.after(0, reset_for_next_file)
                    
                    # Set current file for transcription
                    self._current_batch_file = str(file_path)
                    
                    # For frozen EXE, run whisper via subprocess
                    if getattr(sys, 'frozen', False):
                        self._run_transcription_subprocess_single(str(file_path), file_index=i)
                    else:
                        self._run_transcription_direct_single(str(file_path), file_index=i)
                
                # Clear batch label when done
                self.root.after(0, lambda: self.clear_batch_file_text())
                
                if not self.stop_requested:
                    self.root.after(0, lambda: self.log(f"\n{'='*60}"))
                    self.root.after(0, lambda: self.log(f"✅ Batch complete! {self.batch_total} files processed"))
                    self.root.after(0, lambda: self.log(f"{'='*60}"))
            else:
                # Single file mode
                if getattr(sys, 'frozen', False):
                    self._run_transcription_subprocess()
                else:
                    self._run_transcription_direct()
        finally:
            # Always reset buttons when done
            if getattr(sys, 'frozen', False):
                self.root.after(0, self._reset_ui_after_transcription)
    
    def _reset_ui_after_transcription(self):
        """Reset UI elements after transcription completes"""
        was_stopped = self.stop_requested
        
        self.processing = False
        self.stop_requested = False
        self.timer_running = False
        self._proofreading_mode = False
        self.progress.stop_animation()
        self.process_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        self.clear_batch_file_text()  # Clear batch file label
        
        if was_stopped:
            # User cancelled - reset progress
            self.progress['value'] = 0
            self.set_progress_text("Cancelled")
            self.log("❌ Transcription cancelled")
        elif self.progress['value'] < 100:
            # Didn't complete successfully - reset
            self.progress['value'] = 0
            self.set_progress_text("")
    
    def _run_transcription_subprocess(self):
        """Run transcription via subprocess (for frozen EXE) - single file from UI"""
        video_file = self.video_path.get()
        self._run_transcription_subprocess_single(video_file, file_index=None)
    
    def _run_transcription_subprocess_single(self, video_file, file_index=None):
        """Run transcription via subprocess for a single file"""
        output_dir = self.output_dir.get()
        
        if self.stop_requested:
            self.log("❌ Transcription cancelled before starting")
            return
        
        # Show batch progress if in batch mode
        batch_prefix = ""
        if self.batch_files:
            batch_prefix = f"[{self.batch_index}/{self.batch_total}] "
            self.root.after(0, lambda: self.set_progress_text(
                f"File {self.batch_index}/{self.batch_total}: {Path(video_file).name[:30]}..."))
        
        self.log(f"\n{'='*60}")
        self.log(f"{batch_prefix}Starting transcription of: {os.path.basename(video_file)}")
        self.log(f"Model: {self.model.get()}")
        self.log(f"Language: {self.language.get()}")
        device = "cuda" if self.use_gpu.get() else "cpu"
        self.log(f"Device: {'GPU' if device == 'cuda' else 'CPU'}")
        self.log(f"{'='*60}\n")
        
        # Get audio duration for progress tracking
        audio_duration = None
        try:
            # Hide console window for ffprobe
            startupinfo = None
            creationflags = 0
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE
                creationflags = subprocess.CREATE_NO_WINDOW
            
            result = subprocess.run(
                ['ffprobe', '-v', 'error', '-show_entries', 'format=duration',
                 '-of', 'default=noprint_wrappers=1:nokey=1', video_file],
                capture_output=True, text=True, timeout=10,
                startupinfo=startupinfo,
                creationflags=creationflags
            )
            if result.returncode == 0:
                audio_duration = float(result.stdout.strip())
                self.log(f"Audio duration: {audio_duration:.1f} seconds\n")
                self.last_audio_duration = audio_duration
        except:
            pass
        
        video_name = Path(video_file).stem
        srt_file = Path(output_dir) / f"{video_name}.srt"
        json_file = Path(output_dir) / f"{video_name}_temp.json"
        
        # Create the transcription helper script inline
        helper_script = '''
import sys, json, argparse, os

# Force unbuffered output
sys.stdout = os.fdopen(sys.stdout.fileno(), 'w', buffering=1)
sys.stderr = os.fdopen(sys.stderr.fileno(), 'w', buffering=1)

# Fix Windows encoding issues
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    os.environ['PYTHONIOENCODING'] = 'utf-8'

parser = argparse.ArgumentParser()
parser.add_argument('video_file')
parser.add_argument('--model', default='turbo')
parser.add_argument('--language', default='da')
parser.add_argument('--device', default='cpu')
parser.add_argument('--output', required=True)
parser.add_argument('--initial_prompt', default='')
parser.add_argument('--word_timestamps', type=lambda x: x.lower() == 'true', default=True)
parser.add_argument('--condition_on_previous_text', type=lambda x: x.lower() == 'true', default=True)
parser.add_argument('--no_speech_threshold', type=float, default=0.6)
parser.add_argument('--hallucination_silence_threshold', type=float, default=0.0)
args = parser.parse_args()

import whisper

# Verify the file exists and print its path
video_path = os.path.abspath(args.video_file)
print(f"Input file: {video_path}", flush=True)
if not os.path.exists(video_path):
    print(f"ERROR: File not found: {video_path}", flush=True)
    sys.exit(1)
print(f"File size: {os.path.getsize(video_path)} bytes", flush=True)

# Check if model needs to be downloaded
model_name = args.model
whisper_cache = os.path.join(os.path.expanduser("~"), ".cache", "whisper")

# For "large" and "turbo", find the latest version available
def find_model_file(name, cache_dir):
    """Find the actual model file, checking for latest versions"""
    if not os.path.exists(cache_dir):
        return None
    
    cache_files = os.listdir(cache_dir)
    
    if name == "large":
        # Look for large-v* files, pick highest version
        large_files = [f for f in cache_files if f.startswith("large-v") and f.endswith(".pt") and "turbo" not in f]
        if large_files:
            large_files.sort(reverse=True)  # large-v3.pt > large-v2.pt > large-v1.pt
            return os.path.join(cache_dir, large_files[0])
    elif name == "turbo":
        # Look for turbo files, pick latest
        turbo_files = [f for f in cache_files if "turbo" in f and f.endswith(".pt")]
        if turbo_files:
            turbo_files.sort(reverse=True)
            return os.path.join(cache_dir, turbo_files[0])
    else:
        # Direct match for specific versions
        direct_file = os.path.join(cache_dir, f"{name}.pt")
        if os.path.exists(direct_file):
            return direct_file
        # Also check for files containing the name
        for f in cache_files:
            if name in f and f.endswith(".pt"):
                return os.path.join(cache_dir, f)
    
    return None

model_file = find_model_file(model_name, whisper_cache)
model_exists = model_file is not None and os.path.exists(model_file)

if not model_exists:
    print(f"Downloading model {model_name}... (this is a one-time download)", flush=True)

print(f"Loading model {model_name}...", flush=True)
model = whisper.load_model(args.model, device=args.device)
print("Model loaded", flush=True)

# Build transcription options
opts = {'language': args.language, 'word_timestamps': args.word_timestamps, 'verbose': True,
        'condition_on_previous_text': args.condition_on_previous_text,
        'no_speech_threshold': args.no_speech_threshold}
if args.initial_prompt:
    opts['initial_prompt'] = args.initial_prompt
if args.hallucination_silence_threshold > 0:
    opts['hallucination_silence_threshold'] = args.hallucination_silence_threshold

print("Transcribing...", flush=True)
print(f"Options: condition_on_previous_text={args.condition_on_previous_text}, no_speech_threshold={args.no_speech_threshold}", flush=True)

# Use file path directly - let Whisper handle audio loading
result = model.transcribe(video_path, **opts)
print("Transcription complete", flush=True)

# Debug: print first few segments
if result.get('segments'):
    first_seg = result['segments'][0]
    print(f"First segment: {first_seg.get('start', 'N/A')}s - {first_seg.get('end', 'N/A')}s: {first_seg.get('text', '')[:50]}", flush=True)
    print(f"Total segments: {len(result['segments'])}", flush=True)
else:
    print("WARNING: No segments in result!", flush=True)

output = {'text': result.get('text', ''), 'language': result.get('language', ''), 'segments': []}
for seg in result.get('segments', []):
    s = {'id': seg.get('id', 0), 'start': seg.get('start', 0), 'end': seg.get('end', 0), 'text': seg.get('text', ''), 'words': []}
    for w in seg.get('words', []):
        s['words'].append({'word': w.get('word', ''), 'start': w.get('start', 0), 'end': w.get('end', 0)})
    output['segments'].append(s)

with open(args.output, 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False)
print("DONE", flush=True)
'''
        
        # Build command to run helper script
        # Use -u for unbuffered output so we get line-by-line progress
        cmd = [
            'py', '-u', '-c', helper_script,
            video_file,
            '--model', self.model.get(),
            '--language', self.get_language_code(),
            '--device', device,
            '--output', str(json_file),
            '--word_timestamps', 'true' if self.use_word_timestamps.get() else 'false',
            '--condition_on_previous_text', 'true' if self.condition_on_previous_text.get() else 'false',
            '--no_speech_threshold', str(self.no_speech_threshold.get()),
            '--hallucination_silence_threshold', str(self.hallucination_silence_threshold.get()),
        ]
        
        # Add context/prompt if provided
        if file_index is not None:
            context = self.get_prompt_for_file(file_index)
        else:
            context = self.context_prompt.get().strip()
        if context:
            cmd.extend(['--initial_prompt', context])
            self.log(f"Using context: {context[:100]}{'...' if len(context) > 100 else ''}\n")
        
        # Log anti-hallucination settings if non-default
        if not self.condition_on_previous_text.get():
            self.log("Anti-hallucination: condition_on_previous_text = OFF")
        if self.no_speech_threshold.get() != 0.6:
            self.log(f"Anti-hallucination: no_speech_threshold = {self.no_speech_threshold.get()}")
        if self.hallucination_silence_threshold.get() > 0:
            self.log(f"Anti-hallucination: silence_threshold = {self.hallucination_silence_threshold.get()}s")
        
        self.update_progress_with_time(10, "Starting transcription...")
        
        try:
            # Set up environment with UTF-8 encoding
            env = os.environ.copy()
            env['PYTHONIOENCODING'] = 'utf-8'
            
            # Hide console window on Windows
            startupinfo = None
            creationflags = 0
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE
                creationflags = subprocess.CREATE_NO_WINDOW
            
            # Run transcription helper and capture output
            self.transcription_process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding='utf-8',
                errors='replace',
                bufsize=1,
                env=env,
                startupinfo=startupinfo,
                creationflags=creationflags
            )
            
            # Read output line by line
            for line in self.transcription_process.stdout:
                if self.stop_requested:
                    self.transcription_process.terminate()
                    self.log("\n❌ Transcription cancelled by user")
                    return
                
                line = line.strip()
                if not line:
                    continue
                
                import re  # Import at top of loop for all uses
                    
                # Filter out warnings
                if any(skip in line for skip in ['UserWarning', 'FutureWarning', 'Triton', 'warnings.warn']):
                    continue
                if line.startswith('C:\\') and 'site-packages' in line:
                    continue
                
                # Update progress based on status
                if 'Loading model' in line:
                    self.update_progress_with_time(15, "Loading model...")
                    self.log(f"📥 {line}")
                elif 'Downloading' in line or 'download' in line.lower():
                    # Model download in progress
                    self.update_progress_with_time(10, "Downloading model...")
                    self.log(f"📥 {line}")
                elif '%|' in line or 'B/s' in line or 'MB' in line:
                    # tqdm progress bar output - show download progress
                    # Extract percentage if present
                    pct_match = re.search(r'(\d+)%', line)
                    if pct_match:
                        pct = int(pct_match.group(1))
                        self.update_progress_with_time(5 + (pct * 0.15), f"Downloading model... {pct}%")
                    # Don't log every progress line to avoid spam, just update status
                elif 'Model loaded' in line:
                    self.update_progress_with_time(25, "Model loaded")
                    self.log(f"✅ {line}")
                elif 'Transcribing' in line and 'Options' not in line:
                    self.update_progress_with_time(30, "Transcribing...")
                    self.log(f"🎤 {line}")
                elif 'Transcription complete' in line:
                    self.update_progress_with_time(80, "Processing...")
                    self.log(f"✅ {line}")
                elif '-->' in line and '[' in line:
                    # Timestamp line like: [00:08.580 --> 00:11.460]  Text here
                    self.log(line)
                    try:
                        # Parse timestamp - format is [MM:SS.mmm --> MM:SS.mmm]
                        match = re.search(r'\[(\d+):(\d+)\.(\d+)\s*-->', line)
                        if match:
                            minutes = int(match.group(1))
                            seconds = int(match.group(2))
                            current_time = minutes * 60 + seconds
                            
                            # Update audio position for ETA calculation
                            self.last_audio_position = current_time
                            
                            # Calculate progress based on audio duration if available
                            if audio_duration and audio_duration > 0:
                                progress = 30 + (current_time / audio_duration) * 55
                                progress = min(85, progress)
                                
                                # Calculate ETA
                                import time
                                time_elapsed = time.time() - self.transcription_start_time
                                elapsed_min = int(time_elapsed // 60)
                                elapsed_sec = int(time_elapsed % 60)
                                
                                processing_speed = current_time / time_elapsed if time_elapsed > 0 else 0
                                remaining_audio = audio_duration - current_time
                                eta_seconds = remaining_audio / processing_speed if processing_speed > 0 else 0
                                eta_min = int(eta_seconds // 60)
                                eta_sec = int(eta_seconds % 60)
                                percent = int((current_time / audio_duration) * 100)
                                
                                # Update progress bar and label (use root.after for thread safety)
                                progress_text = (f"Transcribing... {percent}% ({current_time:.0f}s / {audio_duration:.0f}s) | "
                                               f"Elapsed: {elapsed_min}:{elapsed_sec:02d} | "
                                               f"ETA: {eta_min}:{eta_sec:02d} | "
                                               f"Speed: {processing_speed:.1f}x")
                                
                                def update_ui(p=progress, t=progress_text):
                                    self.progress['value'] = p
                                    self.set_progress_text(t)
                                self.root.after(0, update_ui)
                                
                                # Set timestamp so timer doesn't overwrite
                                self._last_detail_progress_time = time.time()
                                self.last_progress_value = progress
                            else:
                                # No duration info - just show basic progress
                                progress = min(75, 30 + minutes * 3)
                                self.root.after(0, lambda p=progress: self.progress.configure(value=p))
                                self.last_progress_value = progress
                    except Exception as e:
                        pass  # Don't let parsing errors stop transcription
                elif 'DONE' in line:
                    pass  # Skip final marker
                elif 'Error' in line or 'Traceback' in line or 'Exception' in line:
                    self.log(f"⚠️ {line}")
                else:
                    self.log(line)
            
            self.transcription_process.wait()
            
            # Check if stopped by user
            if self.stop_requested:
                return  # Exit cleanly - UI reset handled by finally block
            
            # Check if JSON was created (success indicator)
            if json_file.exists():
                # Success - JSON was created regardless of return code
                pass
            elif self.transcription_process.returncode != 0:
                self.log(f"\n❌ Transcription failed with code {self.transcription_process.returncode}")
                return
            
        except Exception as e:
            if not self.stop_requested:  # Don't log error if user cancelled
                self.log(f"\n❌ Error: {e}")
                import traceback
                self.log(traceback.format_exc())
            return
        finally:
            self.transcription_process = None
        
        # Check if JSON was created
        if not json_file.exists():
            self.log(f"⚠️ Transcription output not found at {json_file}")
            return
        
        self.update_progress_with_time(85, "Generating subtitles...")
        
        # Load the JSON result
        try:
            import json
            with open(json_file, 'r', encoding='utf-8') as f:
                result = json.load(f)
            
            # Clean up temp JSON file
            json_file.unlink()
            
        except Exception as e:
            self.log(f"❌ Error reading transcription result: {e}")
            return
        
        # Check if we have word-level timestamps
        has_word_timestamps = (
            self.use_word_timestamps.get() and 
            result.get('segments') and 
            len(result['segments']) > 0 and
            result['segments'][0].get('words') and
            len(result['segments'][0]['words']) > 0
        )
        
        if has_word_timestamps:
            # Use our smart SRT generator with word-level timestamps
            self.log("📝 Using word-level timestamps for smart subtitle splitting...")
            generate_smart_srt(
                result, 
                srt_file,
                max_chars_per_line=self.max_chars_per_line.get(),
                max_lines=self.max_line_count.get()
            )
            self.log("✅ Smart SRT generation complete!")
        else:
            # Generate basic SRT from segments
            self.log("📝 Generating SRT from segments...")
            with open(srt_file, 'w', encoding='utf-8') as f:
                for i, seg in enumerate(result.get('segments', []), 1):
                    start = self._format_srt_time(seg['start'])
                    end = self._format_srt_time(seg['end'])
                    text = seg.get('text', '').strip()
                    f.write(f"{i}\n{start} --> {end}\n{text}\n\n")
            self.log("✅ SRT generation complete!")
        
        # AI Proofreading
        proofread_file = None
        if self.use_ai_proofreading.get():
            self._proofreading_mode = True
            self._proofreading_batch_info = None
            try:
                # For frozen EXE, run proofreading via subprocess to get proper SSL support
                if getattr(sys, 'frozen', False):
                    proofread_file = self._proofread_via_subprocess(
                        srt_file=str(srt_file),
                        language=self.get_language_code(),
                        context=self.context_prompt.get().strip()
                    )
                else:
                    proofread_file = self.proofread_srt_with_ai(
                        srt_file=str(srt_file),
                        language=self.get_language_code(),
                        context=self.context_prompt.get().strip()
                    )
            except Exception as e:
                self.log(f"❌ AI Proofreading error: {e}")
                import traceback
                self.log(traceback.format_exc())
            finally:
                self._proofreading_mode = False
        
        # Done!
        self.progress['value'] = 100
        self.progress.stop_animation()  # Stop gradient animation
        self.timer_running = False  # Stop timer
        
        if self.transcription_start_time:
            import time
            total_time = time.time() - self.transcription_start_time
            minutes = int(total_time // 60)
            seconds = int(total_time % 60)
            self.set_progress_text(f"✓ Done in {minutes}:{seconds:02d}")
        else:
            self.set_progress_text("✓ Done")
        
        self.log(f"\n📁 Output saved to: {output_dir}")
        if proofread_file:
            self.log(f"   {srt_file.name} (proofread)")
        else:
            self.log(f"   {srt_file.name}")
        
        self.play_completion_chime()
    
    def _format_srt_time(self, seconds):
        """Format seconds to SRT timestamp"""
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        secs = seconds % 60
        return f"{hours:02d}:{minutes:02d}:{secs:06.3f}".replace('.', ',')
    
    def _run_transcription_direct(self):
        """Run transcription directly (for running as script) - single file from UI"""
        video_file = self.video_path.get()
        self._run_transcription_direct_single(video_file, file_index=None)
    
    def _run_transcription_direct_single(self, video_file, file_index=None):
        """Run transcription directly for a single file"""
        try:
            import whisper
            
            output_dir = self.output_dir.get()
            
            if self.stop_requested:
                self.log("❌ Transcription cancelled before starting")
                return
            
            self.log(f"\n{'='*60}")
            self.log(f"Starting transcription of: {os.path.basename(video_file)}")
            self.log(f"Model: {self.model.get()}")
            self.log(f"Language: {self.language.get()}")
            
            # Determine device
            device = "cpu"
            if self.use_gpu.get():
                try:
                    import torch
                    if torch.cuda.is_available():
                        device = "cuda"
                        self.log(f"Device: GPU ({torch.cuda.get_device_name(0)})")
                    else:
                        self.log(f"Device: CPU (GPU requested but not available)")
                except ImportError:
                    self.log(f"Device: CPU (PyTorch not installed)")
            else:
                self.log(f"Device: CPU")
            
            self.log(f"{'='*60}\n")
            
            if self.stop_requested:
                self.log("❌ Transcription cancelled")
                return
            
            # Load model
            self.log(f"Loading model {self.model.get()} on {device}...")
            self.update_progress_with_time(5, "Loading model...")
            
            model_name = self.model.get()
            
            # Check if model needs to be downloaded
            try:
                whisper_cache = os.path.join(os.path.expanduser("~"), ".cache", "whisper")
                # Check for various possible model file names
                possible_names = [
                    f"{model_name}.pt",
                    f"large-v3-turbo.pt" if model_name == "turbo" else None,
                ]
                model_exists = any(
                    name and os.path.exists(os.path.join(whisper_cache, name)) 
                    for name in possible_names
                )
                if not model_exists and os.path.exists(whisper_cache):
                    # Also check if any file contains the model name
                    cache_files = os.listdir(whisper_cache)
                    model_exists = any(model_name in f for f in cache_files)
                
                if not model_exists:
                    self.log(f"📥 Downloading model '{model_name}'...")
                    self.log("This is a one-time download and may take a few minutes.\n")
                    self.update_progress_with_time(5, f"Downloading {model_name}...")
            except:
                pass
            
            # Load the model (will download if needed)
            model = whisper.load_model(model_name, device=device)
            self.log("✅ Model loaded")
            
            if self.stop_requested:
                self.log("❌ Transcription cancelled")
                return
            
            # Get audio duration for progress tracking
            self.update_progress_with_time(10, "Analyzing audio...")
            
            try:
                import subprocess
                result = subprocess.run(
                    ['ffprobe', '-v', 'error', '-show_entries', 'format=duration',
                     '-of', 'default=noprint_wrappers=1:nokey=1', video_file],
                    capture_output=True, text=True, timeout=10
                )
                audio_duration = float(result.stdout.strip()) if result.returncode == 0 else None
            except:
                audio_duration = None
            
            # Transcribe
            self.log(f"\nTranscribing (this may take a few minutes)...")
            self.update_progress_with_time(15, "Transcribing...")
            
            if audio_duration:
                self.log(f"Audio duration: {audio_duration:.1f} seconds")
            
            self.log("Progress updates will appear below:\n")
            
            # Build transcribe options
            transcribe_options = {
                'language': self.get_language_code(),
                'word_timestamps': self.use_word_timestamps.get(),
                'verbose': True,  # Keep verbose for progress
                'condition_on_previous_text': self.condition_on_previous_text.get(),
                'no_speech_threshold': self.no_speech_threshold.get(),
            }
            
            # Add hallucination silence threshold if set
            if self.hallucination_silence_threshold.get() > 0:
                transcribe_options['hallucination_silence_threshold'] = self.hallucination_silence_threshold.get()
            
            # Add context/prompt if provided
            if file_index is not None:
                context = self.get_prompt_for_file(file_index)
            else:
                context = self.context_prompt.get().strip()
            if context:
                transcribe_options['initial_prompt'] = context
                self.log(f"Using context: {context[:100]}{'...' if len(context) > 100 else ''}\n")
            
            # Log anti-hallucination settings if non-default
            if not self.condition_on_previous_text.get():
                self.log("Anti-hallucination: condition_on_previous_text = OFF")
            if self.no_speech_threshold.get() != 0.6:
                self.log(f"Anti-hallucination: no_speech_threshold = {self.no_speech_threshold.get()}")
            if self.hallucination_silence_threshold.get() > 0:
                self.log(f"Anti-hallucination: silence_threshold = {self.hallucination_silence_threshold.get()}s")
            
            # Track progress (transcription_start_time already set in start_transcription)
            self.last_progress_update = 0
            
            # Capture stdout/stderr in real-time
            import io
            import sys
            import re
            import time
            
            class LogCapture(io.StringIO):
                def __init__(self, log_callback, stop_check, progress_callback, audio_duration):
                    super().__init__()
                    self.log_callback = log_callback
                    self.stop_check = stop_check
                    self.progress_callback = progress_callback
                    self.audio_duration = audio_duration
                    self.buffer = ""
                    self.start_time = time.time()
                
                def write(self, s):
                    # Check for stop request
                    if self.stop_check():
                        raise KeyboardInterrupt("Transcription stopped by user")
                    
                    self.buffer += s
                    if '\n' in self.buffer or '\r' in self.buffer:
                        lines = self.buffer.split('\n')
                        for line in lines[:-1]:
                            line = line.strip('\r').strip()
                            
                            # Filter out warnings and noise
                            if not line:
                                continue
                            if 'UserWarning' in line:
                                continue
                            if 'Triton' in line:
                                continue
                            if 'warnings.warn' in line:
                                continue
                            if line.startswith('C:\\') and 'site-packages' in line:
                                continue
                            if 'FutureWarning' in line:
                                continue
                                
                            # Log the line
                            self.log_callback(line)
                            
                            # Try to extract timestamp for progress
                            # Whisper outputs timestamps like [00:01.000 --> 00:05.000]
                            if self.audio_duration and '-->' in line:
                                try:
                                    match = re.search(r'\[(\d+):(\d+)\.(\d+)\s*-->', line)
                                    if match:
                                        minutes = int(match.group(1))
                                        seconds = int(match.group(2))
                                        current_time = minutes * 60 + seconds
                                        
                                        # Calculate progress (15% to 85% of progress bar)
                                        progress = 15 + (current_time / self.audio_duration) * 70
                                        progress = min(85, progress)  # Cap at 85%
                                        
                                        self.progress_callback(progress, current_time)
                                except Exception as e:
                                    # Log parsing errors for debugging
                                    import traceback
                                    self.log_callback(f"[DEBUG] Progress parsing error: {e}")
                        self.buffer = lines[-1]
                    return len(s)
                
                def flush(self):
                    if self.buffer.strip():
                        self.log_callback(self.buffer)
                        self.buffer = ""
            
            def update_progress(value, current_time=None):
                # Store the latest progress data for the timer to use
                self.last_progress_value = value
                self.last_audio_position = current_time
                self.last_audio_duration = audio_duration
                
                # Schedule GUI update in main thread
                def update_in_main_thread():
                    try:
                        self.progress['value'] = value
                        self._update_progress_label()
                    except Exception as e:
                        pass
                
                # Execute in main thread
                try:
                    self.root.after(0, update_in_main_thread)
                except Exception as e:
                    pass
            
            old_stdout = sys.stdout
            old_stderr = sys.stderr
            
            try:
                sys.stdout = LogCapture(self.log, lambda: self.stop_requested, update_progress, audio_duration)
                sys.stderr = LogCapture(self.log, lambda: self.stop_requested, update_progress, audio_duration)
                
                result = model.transcribe(video_file, **transcribe_options)
                
            except KeyboardInterrupt:
                self.log("\n❌ Transcription cancelled by user")
                self.set_progress_text("Cancelled")
                messagebox.showinfo("Cancelled", "Transcription was cancelled.")
                return
            finally:
                sys.stdout = old_stdout
                sys.stderr = old_stderr
            
            if self.stop_requested:
                self.log("\n❌ Transcription cancelled")
                self.set_progress_text("Cancelled")
                return
            
            self.progress['value'] = 90
            self.update_progress_with_time(90, "Generating SRT file...")
            self.log("\n✅ Transcription processing complete!")
            
            # Generate SRT
            self.log("\n📝 Generating SRT file...")
            video_name = Path(video_file).stem
            srt_file = Path(output_dir) / f"{video_name}.srt"
            
            # Check if we have word-level timestamps
            has_word_timestamps = (
                self.use_word_timestamps.get() and 
                result.get('segments') and 
                result['segments'][0].get('words')
            )
            
            if has_word_timestamps:
                # Use our smart SRT generator with word-level timestamps
                self.log("Using word-level timestamps for smart subtitle splitting...")
                generate_smart_srt(
                    result, 
                    srt_file,
                    max_chars_per_line=self.max_chars_per_line.get(),
                    max_lines=self.max_line_count.get()
                )
                self.log("✅ Smart SRT generation complete!")
            else:
                # Fallback to Whisper's WriteSRT
                self.log("Word timestamps not available, using standard SRT generation...")
                from whisper.utils import WriteSRT
                writer = WriteSRT(output_dir)
                writer(result, video_file, {
                    'max_line_width': self.max_chars_per_line.get(),
                    'max_line_count': self.max_line_count.get(),
                    'highlight_words': False
                })
            
            self.log(f"✅ Transcription complete!")
            
            if srt_file.exists():
                self.log(f"SRT file created: {srt_file}")
                
                # AI Proofreading (if enabled)
                proofread_file = None
                if self.use_ai_proofreading.get():
                    # Set proofreading mode for progress label
                    self._proofreading_mode = True
                    self.progress['value'] = 95
                    self.set_progress_text("AI Proofreading...")
                    
                    proofread_file = self.proofread_srt_with_ai(
                        str(srt_file), 
                        language=self.get_language_code(),
                        context=self.context_prompt.get().strip()
                    )
                    
                    # Clear proofreading mode
                    self._proofreading_mode = False
                
                # Set progress to 100% and show final time
                self.progress['value'] = 100
                
                # Calculate final time
                if self.transcription_start_time:
                    import time
                    total_time = time.time() - self.transcription_start_time
                    minutes = int(total_time // 60)
                    seconds = int(total_time % 60)
                    self.set_progress_text(f"✓ Done in {minutes}:{seconds:02d}")
                else:
                    self.set_progress_text("✓ Done")
                
                # Log output info
                self.log(f"\n📁 Output saved to: {output_dir}")
                if proofread_file:
                    self.log(f"   {srt_file.name} (proofread)")
                else:
                    self.log(f"   {srt_file.name}")
                
                # Play completion chime
                self.play_completion_chime()
            else:
                self.progress['value'] = 100
                self.set_progress_text("Complete (file not found)")
                self.log("⚠️ SRT file not found at expected location")
                
        except ImportError:
            self.log(f"\n❌ Whisper not installed!")
            self.log("Please click 'Check/Update Model' first to install Whisper.")
            messagebox.showerror("Error", "Whisper not installed.\n\nPlease click 'Check/Update Model' button first.")
        except Exception as e:
            self.log(f"\n❌ Error: {str(e)}")
            import traceback
            self.log(traceback.format_exc())
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        finally:
            # Save stop_requested state before resetting
            was_stopped = self.stop_requested
            
            self.processing = False
            self.stop_requested = False
            self.timer_running = False
            self._proofreading_mode = False  # Reset proofreading mode
            self.progress.stop_animation()  # Stop gradient animation
            self.process_btn.config(state='normal')
            self.stop_btn.config(state='disabled')
            if not was_stopped:
                # Keep progress at 100% if completed, otherwise reset
                if self.progress['value'] < 100:
                    self.progress['value'] = 0
                    self.set_progress_text("")
                else:
                    # Show final elapsed time
                    if self.transcription_start_time:
                        import time
                        total_time = time.time() - self.transcription_start_time
                        minutes = int(total_time // 60)
                        seconds = int(total_time % 60)
                        current_text = self.get_progress_text()
                        if 'Done' in current_text or 'Complete' in current_text:
                            self.set_progress_text(f"✓ Done in {minutes}:{seconds:02d}")
    
    def on_closing(self):
        """Handle window close - save settings"""
        self.save_settings()
        self.root.destroy()

def main():
    # On Windows, set app user model ID before creating window
    # This ensures the taskbar icon shows correctly
    import sys
    if sys.platform == 'win32':
        try:
            import ctypes
            myappid = 'finalfilm.finalwhisper.gui.1'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception:
            pass
    
    root = tk.Tk()
    app = WhisperGUI(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()
