#!/usr/bin/env python3
"""
Final Whisper - AI-powered transcription with smart subtitle formatting
By Final Film
"""

VERSION = "1.00"

import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import subprocess
import threading
import os
from pathlib import Path
import sys
import re


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
        self._value = max(0, min(self._maximum, value))
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
    
    def set_value(self, value):
        """Set the progress value (0-100)."""
        self._value = max(0, min(self._maximum, value))
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
        self.root.geometry("1100x900")
        
        # Set window icon
        self.set_window_icon()
        
        # Apply dark theme
        self.setup_dark_theme()
        
        # Variables
        self.video_path = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(Path.home() / "whisper_output"))
        self.language = tk.StringVar(value="da: Danish")
        self.model = tk.StringVar(value="turbo")
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
        
        # AI Proofreading settings
        self.use_ai_proofreading = tk.BooleanVar(value=False)
        self.anthropic_api_key = tk.StringVar(value="")
        
        # Load saved API key if exists
        self.load_api_key()
        
        self.create_widgets()
        self.check_whisper_installation()
        self.check_gpu_availability()
    
    def set_window_icon(self):
        """Set the window icon"""
        try:
            # Try to load icon from various locations
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
            
            for icon_path in icon_paths:
                if os.path.exists(icon_path):
                    self.root.iconbitmap(icon_path)
                    return
                    
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
            
            for logo_path in logo_paths:
                if os.path.exists(logo_path):
                    from PIL import Image, ImageTk
                    img = Image.open(logo_path)
                    # Resize to ~65% of original (was 180x65, now ~117x42)
                    new_width = int(img.width * 0.65)
                    new_height = int(img.height * 0.65)
                    img = img.resize((new_width, new_height), Image.LANCZOS)
                    self.logo_image = ImageTk.PhotoImage(img)
                    logo_label = ttk.Label(header_frame, image=self.logo_image)
                    logo_label.pack(side=tk.LEFT, padx=(0, 15), pady=(5, 5))
                    break
        except Exception:
            pass  # Logo is optional
        
        # Title (right of logo)
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(side=tk.LEFT, fill=tk.Y)
        
        title_label = ttk.Label(title_frame, text=f"Final Whisper v{VERSION}", font=("Segoe UI", 18, "bold"))
        title_label.pack(anchor=tk.W)
        
        subtitle_label = ttk.Label(title_frame, text="AI-powered transcription with smart formatting", 
                                   font=("Segoe UI", 9))
        subtitle_label.pack(anchor=tk.W)
        
        # Content area with two columns
        content_frame = ttk.Frame(main_frame)
        content_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Left column for settings
        left_frame = ttk.Frame(content_frame)
        left_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 15))
        
        # Right column for log
        right_frame = ttk.Frame(content_frame)
        right_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # File selection (row 0)
        file_section = CollapsibleFrame(left_frame, text="Files")
        file_section.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        file_frame = file_section.content
        
        ttk.Label(file_frame, text="Video File:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(file_frame, textvariable=self.video_path, width=38).grid(row=0, column=1, padx=8)
        ttk.Button(file_frame, text="Browse", command=self.browse_video).grid(row=0, column=2)
        
        ttk.Label(file_frame, text="Output Folder:").grid(row=1, column=0, sticky=tk.W, pady=(8,0))
        ttk.Entry(file_frame, textvariable=self.output_dir, width=38).grid(row=1, column=1, padx=8, pady=(8,0))
        ttk.Button(file_frame, text="Browse", command=self.browse_output).grid(row=1, column=2, pady=(8,0))
        
        # Transcription settings (row 1)
        settings_section = CollapsibleFrame(left_frame, text="Transcription")
        settings_section.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        settings_frame = settings_section.content
        settings_frame.columnconfigure(1, weight=1)  # Make column 1 expand
        
        ttk.Label(settings_frame, text="Language:").grid(row=0, column=0, sticky=tk.W)
        self.lang_combo = ttk.Combobox(settings_frame, textvariable=self.language,
                                 values=self.get_available_languages(),
                                 width=20, state="readonly")
        self.lang_combo.grid(row=0, column=1, sticky=tk.W, padx=8)
        
        ttk.Label(settings_frame, text="Context/Prompt:").grid(row=1, column=0, sticky=(tk.W, tk.N), pady=(8,0))
        
        # Create context entry with dark theme colors - full width and taller
        self.context_entry = tk.Text(settings_frame, height=5, wrap=tk.WORD,
                                     bg='#3c3c3c', fg='#ffffff', insertbackground='#ffffff',
                                     relief='flat', font=('Segoe UI', 9), padx=6, pady=6)
        self.context_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=8, pady=(8,0))
        
        # Bind text widget to variable
        def update_context(*args):
            self.context_prompt.set(self.context_entry.get("1.0", tk.END).strip())
        self.context_entry.bind("<KeyRelease>", update_context)
        
        hint_label = ttk.Label(settings_frame, text="Names, terms, companies to help recognition...", 
                 font=("Segoe UI", 8))
        hint_label.grid(row=2, column=1, sticky=tk.W, padx=8, pady=(2, 0))
        
        # Model section (row 2) - will be collapsed dynamically if GPU+CUDA detected
        self.model_section = CollapsibleFrame(left_frame, text="Model")
        self.model_section.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
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
        
        # CUDA Toolkit button (only shown when GPU is available but CUDA toolkit is missing)
        self.cuda_btn = ttk.Button(model_frame, text="Install CUDA Toolkit", 
                                   command=self.offer_cuda_toolkit)
        # Will be shown/hidden dynamically
        
        # Initial model status update
        self.update_model_status()
        
        # Subtitle formatting settings (row 3)
        format_section = CollapsibleFrame(left_frame, text="Subtitle Formatting", collapsed=True)
        format_section.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        format_frame = format_section.content
        
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
        self.proofread_section = CollapsibleFrame(left_frame, text="AI Proofreading", collapsed=True)
        self.proofread_section.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        proofread_frame = self.proofread_section.content
        
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
        
        # Progress
        self.progress = GradientProgressBar(left_frame)
        self.progress.grid(row=6, column=0, sticky=(tk.W, tk.E), pady=(0, 4))
        
        self.progress_label = ttk.Label(left_frame, text="", font=("Segoe UI", 10))
        self.progress_label.grid(row=7, column=0, sticky=tk.W, pady=(4, 0))
        
        # Output log (right side)
        log_frame = ttk.LabelFrame(right_frame, text="  Output Log  ", padding="12")
        log_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Create log with dark theme
        self.log_text = scrolledtext.ScrolledText(log_frame, height=32, width=55, state='disabled',
                                                   bg='#1e1e1e', fg='#d4d4d4', insertbackground='#ffffff',
                                                   relief='flat', font=('Consolas', 9), padx=8, pady=8)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
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
        """Add message to log (thread-safe)"""
        def do_log():
            self.log_text.config(state='normal')
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
        if hasattr(self, 'proofread_section'):
            if self.use_ai_proofreading.get():
                self.proofread_section.set_status("✓ ON")
            else:
                self.proofread_section.set_status("OFF")
    
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
    
    def get_config_path(self):
        """Get path to config file"""
        config_dir = Path.home() / ".final_whisper"
        config_dir.mkdir(exist_ok=True)
        return config_dir / "config.json"
    
    def load_api_key(self):
        """Load API key from config file"""
        try:
            config_path = self.get_config_path()
            if config_path.exists():
                import json
                with open(config_path, 'r') as f:
                    config = json.load(f)
                    if 'anthropic_api_key' in config and config['anthropic_api_key']:
                        self.anthropic_api_key.set(config['anthropic_api_key'])
                        # Auto-enable proofreading if we have a saved API key
                        self.use_ai_proofreading.set(True)
        except Exception:
            pass
    
    def save_api_key(self):
        """Save API key to config file"""
        try:
            import json
            config_path = self.get_config_path()
            config = {}
            if config_path.exists():
                with open(config_path, 'r') as f:
                    config = json.load(f)
            config['anthropic_api_key'] = self.anthropic_api_key.get()
            with open(config_path, 'w') as f:
                json.dump(config, f)
            self.log("✅ API key saved")
        except Exception as e:
            self.log(f"❌ Failed to save API key: {e}")
    
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
        # Try to get models from whisper module
        try:
            import whisper
            if hasattr(whisper, '_MODELS'):
                # Get model names from whisper's internal list
                models = list(whisper._MODELS.keys())
                # Sort with turbo and larger models first
                preferred_order = ['turbo', 'large-v3-turbo', 'large-v3', 'large-v2', 
                                   'large-v1', 'large', 'medium', 'small', 'base', 'tiny']
                sorted_models = []
                for m in preferred_order:
                    if m in models:
                        sorted_models.append(m)
                        models.remove(m)
                # Add any remaining models not in our preferred order
                sorted_models.extend(sorted(models, reverse=True))
                return sorted_models
        except ImportError:
            pass
        except Exception:
            pass
        
        # Fallback to known models if whisper not installed yet
        return ["turbo", "large-v3", "large-v2", "large", "medium", "small", "base", "tiny"]
    
    def get_available_languages(self):
        """Get list of available Whisper languages with display names"""
        # Try to get languages from whisper module
        try:
            import whisper
            if hasattr(whisper, 'tokenizer') and hasattr(whisper.tokenizer, 'LANGUAGES'):
                languages = whisper.tokenizer.LANGUAGES
                # Create list of "code: Name" entries, sorted by name
                lang_list = [f"{code}: {name.title()}" for code, name in sorted(languages.items(), key=lambda x: x[1])]
                # Move Danish and English to top
                prioritized = []
                for priority_code in ['da', 'en']:
                    for lang in lang_list[:]:
                        if lang.startswith(f"{priority_code}:"):
                            prioritized.append(lang)
                            lang_list.remove(lang)
                            break
                return prioritized + lang_list
        except ImportError:
            pass
        except Exception:
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
            try:
                import whisper
                self.log("✅ Whisper is installed")
            except ImportError:
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
    
    def check_gpu_availability(self):
        """Check if GPU is available for Whisper"""
        def check():
            gpu_ready = False
            try:
                import torch
                if torch.cuda.is_available():
                    gpu_name = torch.cuda.get_device_name(0)
                    self.device_info.set(f"✅ GPU: {gpu_name}")
                    self.use_gpu.set(True)
                    self.log(f"GPU detected: {gpu_name}")
                    gpu_ready = True
                    
                    # Check if CUDA toolkit is available (for Triton kernels)
                    try:
                        # Try to import triton to see if it works
                        import whisper
                        # If we can import whisper and we have a GPU, check for toolkit warnings
                        # We'll show the CUDA button if user reports Triton warnings
                        self.check_cuda_toolkit_needed()
                    except:
                        pass
                else:
                    # PyTorch is installed but no CUDA support
                    self.device_info.set("⚠️ CPU-only PyTorch detected")
                    self.use_gpu.set(False)
                    self.log("⚠️ PyTorch installed but without CUDA support")
                    
                    # Check if NVIDIA GPU exists
                    if self.check_nvidia_gpu():
                        self.offer_gpu_setup()
            except ImportError:
                # PyTorch not installed at all
                self.device_info.set("⚠️ PyTorch not installed")
                self.use_gpu.set(False)
                self.log("⚠️ PyTorch not found")
                
                # Check if NVIDIA GPU exists
                if self.check_nvidia_gpu():
                    self.offer_gpu_setup()
            
            # Collapse Model section if GPU is ready (everything is set up)
            if gpu_ready and hasattr(self, 'model_section'):
                self.root.after(100, self.model_section.collapse)
            
            # Update model status after GPU check completes
            if hasattr(self, 'model_section'):
                self.root.after(150, self.update_model_status)
        
        threading.Thread(target=check, daemon=True).start()
    
    def check_cuda_toolkit_needed(self):
        """Check if CUDA toolkit might be needed and show button"""
        # This is called after GPU is detected
        # We'll show a hint about CUDA toolkit for max performance
        def delayed_check():
            import time
            time.sleep(2)  # Give user time to read GPU detection message
            
            # Check if nvcc (CUDA compiler) is available
            try:
                result = subprocess.run(['nvcc', '--version'], capture_output=True, timeout=5)
                if result.returncode != 0:
                    # CUDA toolkit not found
                    self.show_cuda_toolkit_button()
            except:
                # CUDA toolkit not found
                self.show_cuda_toolkit_button()
        
        threading.Thread(target=delayed_check, daemon=True).start()
    
    def show_cuda_toolkit_button(self):
        """Show the CUDA toolkit installation button"""
        self.cuda_btn.grid(row=2, column=0, columnspan=4, pady=(5,0), sticky=(tk.W, tk.E))
        self.log("\n💡 Tip: Install CUDA Toolkit for maximum GPU performance (optional)")
        self.log("   Click 'Install CUDA Toolkit' button for ~2x faster transcription")
    
    def check_nvidia_gpu(self):
        """Check if NVIDIA GPU is present using nvidia-smi"""
        try:
            result = subprocess.run(['nvidia-smi'], capture_output=True, timeout=5)
            return result.returncode == 0
        except:
            return False
    
    def offer_gpu_setup(self):
        """Offer to install CUDA-enabled PyTorch"""
        response = messagebox.askyesno(
            "GPU Setup Available",
            "An NVIDIA GPU was detected but PyTorch with CUDA support is not installed.\n\n"
            "Would you like to automatically install GPU-accelerated PyTorch?\n\n"
            "This will make transcription ~10x faster!\n\n"
            "Installation may take a few minutes."
        )
        
        if response:
            self.install_cuda_pytorch()
    
    def install_cuda_pytorch(self):
        """Install CUDA-enabled PyTorch"""
        self.log("\n" + "="*60)
        self.log("Installing GPU-accelerated PyTorch...")
        self.log("="*60 + "\n")
        
        def install():
            try:
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
                subprocess.run([sys.executable, "-m", "pip", "uninstall", "-y", "torch", "torchvision", "torchaudio"],
                             capture_output=True)
                
                # Install CUDA PyTorch
                self.log(f"\nInstalling PyTorch from {torch_index}...")
                self.log("This may take several minutes, please wait...\n")
                
                install_cmd = [
                    sys.executable, "-m", "pip", "install", "--user",
                    "torch", "torchvision", "torchaudio",
                    "--index-url", torch_index
                ]
                
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
                    self.log("\n✅ GPU-accelerated PyTorch installed successfully!")
                    self.log("\nPlease restart the application to use GPU acceleration.")
                    messagebox.showinfo("Success", 
                        "GPU-accelerated PyTorch installed successfully!\n\n"
                        "Please restart this application to use GPU acceleration.")
                else:
                    self.log("\n❌ Installation failed!")
                    messagebox.showerror("Installation Failed",
                        "Failed to install GPU-accelerated PyTorch.\n\n"
                        "Check the log for details.")
                    
            except Exception as e:
                self.log(f"\n❌ Error during installation: {str(e)}")
                messagebox.showerror("Error", f"Installation error:\n{str(e)}")
        
        threading.Thread(target=install, daemon=True).start()
    
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
            
            self.offer_gpu_setup()
        else:
            messagebox.showwarning("No NVIDIA GPU", 
                "No NVIDIA GPU detected on this system.\n\n"
                "GPU acceleration requires an NVIDIA graphics card with CUDA support.")
    
    def offer_cuda_toolkit(self):
        """Offer to help install CUDA Toolkit"""
        response = messagebox.askyesnocancel(
            "Install CUDA Toolkit",
            "CUDA Toolkit provides additional GPU optimizations for ~2x faster transcription.\n\n"
            "Note: This is a large download (~3GB) from NVIDIA.\n\n"
            "Would you like to:\n"
            "• YES - Open NVIDIA download page (you install manually)\n"
            "• NO - Continue without CUDA Toolkit (GPU still works)\n"
            "• CANCEL - Do nothing",
            icon='question'
        )
        
        if response is True:
            # Open NVIDIA download page
            import webbrowser
            
            # Try to detect CUDA version for better link
            cuda_version = None
            try:
                result = subprocess.run(['nvidia-smi'], capture_output=True, text=True, timeout=5)
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
            
            self.log("\n📥 Opening NVIDIA CUDA Toolkit download page...")
            self.log("Please download and install the toolkit, then click 'Add CUDA to PATH' button below.")
            
            if cuda_version:
                cuda_major = cuda_version.split('.')[0]
                if cuda_major == "12":
                    url = "https://developer.nvidia.com/cuda-12-1-0-download-archive"
                    self.log(f"Detected CUDA {cuda_version} - opening CUDA 12.1 download page")
                elif cuda_major == "11":
                    url = "https://developer.nvidia.com/cuda-11-8-0-download-archive"
                    self.log(f"Detected CUDA {cuda_version} - opening CUDA 11.8 download page")
                else:
                    url = "https://developer.nvidia.com/cuda-downloads"
                    self.log("Opening latest CUDA Toolkit download page")
            else:
                url = "https://developer.nvidia.com/cuda-downloads"
                self.log("Opening CUDA Toolkit download page")
            
            webbrowser.open(url)
            
            # Replace the button with "Add CUDA to PATH" button
            self.cuda_btn.config(text="Add CUDA to PATH", command=self.add_cuda_to_path)
            
            messagebox.showinfo(
                "CUDA Toolkit Installation",
                "Steps:\n\n"
                "1. Download the installer from the opened webpage\n"
                "2. Run the installer (Express Installation recommended)\n"
                "3. After installation completes, click 'Add CUDA to PATH' button\n\n"
                "Installation may take 10-15 minutes."
            )
        elif response is False:
            self.log("\nℹ️ Continuing without CUDA Toolkit - GPU will still be used")
            self.cuda_btn.grid_remove()  # Hide the button
        # If None (cancelled), do nothing
    
    def add_cuda_to_path(self):
        """Find CUDA installation and add to PATH"""
        self.log("\n🔍 Searching for CUDA Toolkit installation...")
        
        def find_and_add():
            try:
                import winreg
                import os
                
                # Common CUDA installation path
                cuda_base = r"C:\Program Files\NVIDIA GPU Computing Toolkit\CUDA"
                
                if not os.path.exists(cuda_base):
                    self.log("❌ CUDA Toolkit not found in default location")
                    self.log(f"   Expected: {cuda_base}")
                    messagebox.showerror("CUDA Not Found", 
                        "CUDA Toolkit not found in the default installation location.\n\n"
                        "Please ensure CUDA Toolkit is installed first.")
                    return
                
                # Find installed CUDA versions
                cuda_versions = []
                for item in os.listdir(cuda_base):
                    version_path = os.path.join(cuda_base, item)
                    bin_path = os.path.join(version_path, "bin")
                    if os.path.isdir(version_path) and os.path.exists(bin_path):
                        cuda_versions.append((item, bin_path))
                
                if not cuda_versions:
                    self.log("❌ No valid CUDA installations found")
                    messagebox.showerror("CUDA Not Found", 
                        "CUDA Toolkit folder exists but no valid versions found.\n\n"
                        "Please reinstall CUDA Toolkit.")
                    return
                
                # Use the latest version
                cuda_versions.sort(reverse=True)
                version, bin_path = cuda_versions[0]
                
                self.log(f"✅ Found CUDA {version}")
                self.log(f"   Path: {bin_path}")
                
                # Check if already in PATH
                current_path = os.environ.get('PATH', '')
                if bin_path.lower() in current_path.lower():
                    self.log("ℹ️ CUDA is already in PATH")
                    self.log("Restarting application to detect CUDA...")
                    messagebox.showinfo("Already in PATH", 
                        "CUDA is already in your PATH.\n\n"
                        "Please restart this application to detect it.")
                    self.cuda_btn.grid_remove()
                    return
                
                # Add to user PATH (no admin required)
                self.log("Adding CUDA to user PATH...")
                
                try:
                    # Read current user PATH
                    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, 'Environment', 0, winreg.KEY_READ | winreg.KEY_WRITE)
                    try:
                        current_user_path, _ = winreg.QueryValueEx(key, 'PATH')
                    except WindowsError:
                        current_user_path = ''
                    
                    # Add CUDA if not already there
                    if bin_path.lower() not in current_user_path.lower():
                        new_path = current_user_path.rstrip(';') + ';' + bin_path
                        winreg.SetValueEx(key, 'PATH', 0, winreg.REG_EXPAND_SZ, new_path)
                        self.log("✅ CUDA added to user PATH successfully!")
                    
                    winreg.CloseKey(key)
                    
                    # Broadcast environment change
                    import ctypes
                    HWND_BROADCAST = 0xFFFF
                    WM_SETTINGCHANGE = 0x001A
                    SMTO_ABORTIFHUNG = 0x0002
                    result = ctypes.c_long()
                    ctypes.windll.user32.SendMessageTimeoutW(
                        HWND_BROADCAST, WM_SETTINGCHANGE, 0, "Environment",
                        SMTO_ABORTIFHUNG, 5000, ctypes.byref(result)
                    )
                    
                    self.log("\n✅ PATH updated successfully!")
                    self.log("Please restart this application for changes to take effect.")
                    
                    self.cuda_btn.grid_remove()  # Hide the button
                    
                    messagebox.showinfo("Success", 
                        f"CUDA {version} has been added to your PATH!\n\n"
                        "Please restart this application to use CUDA acceleration.")
                    
                except Exception as e:
                    self.log(f"❌ Failed to update PATH: {str(e)}")
                    self.log("\nManual steps:")
                    self.log(f'1. Press Win+R, type "sysdm.cpl" and press Enter')
                    self.log('2. Go to Advanced tab -> Environment Variables')
                    self.log('3. Under User variables, edit PATH')
                    self.log(f'4. Add: {bin_path}')
                    
                    messagebox.showerror("Failed to Update PATH",
                        f"Could not automatically update PATH.\n\n"
                        f"Please add manually:\n{bin_path}\n\n"
                        "See log for instructions.")
                
            except Exception as e:
                self.log(f"❌ Error: {str(e)}")
                messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        
        threading.Thread(target=find_and_add, daemon=True).start()
        
    def browse_video(self):
        """Browse for video file"""
        filename = filedialog.askopenfilename(
            title="Select Video File",
            filetypes=[
                ("Video files", "*.mp4 *.mkv *.avi *.mov *.m4v *.webm *.mp3 *.wav *.m4a *.flac *.ogg"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.video_path.set(filename)
            # Set output folder to same directory as input file
            input_dir = str(Path(filename).parent)
            self.output_dir.set(input_dir)
            
    def browse_output(self):
        """Browse for output directory"""
        dirname = filedialog.askdirectory(title="Select Output Folder")
        if dirname:
            self.output_dir.set(dirname)
            
    def start_transcription(self):
        """Start the transcription process"""
        if self.processing:
            messagebox.showwarning("Processing", "A transcription is already in progress")
            return
            
        if not self.video_path.get():
            messagebox.showerror("Error", "Please select a video file")
            return
            
        if not os.path.exists(self.video_path.get()):
            messagebox.showerror("Error", "Video file does not exist")
            return
            
        # Create output directory
        os.makedirs(self.output_dir.get(), exist_ok=True)
        
        self.processing = True
        self.stop_requested = False
        self.process_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        self.progress['value'] = 0
        self.progress.start_animation()  # Start gradient animation
        self.progress_label.config(text="Starting...")
        
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
            self.progress_label.config(text="Stopping...")
            self.log("\n⚠️ Stop requested - cancelling transcription...")
    
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
        """Update progress label with current stats"""
        import time
        
        if not self.transcription_start_time:
            return
        
        # Don't overwrite the "Done" message
        current_text = self.progress_label.cget('text')
        if '✓ Done' in current_text:
            self.timer_running = False
            return
        
        time_elapsed = time.time() - self.transcription_start_time
        elapsed_min = int(time_elapsed // 60)
        elapsed_sec = int(time_elapsed % 60)
        
        # Check if we're in proofreading mode
        if getattr(self, '_proofreading_mode', False):
            # During proofreading, show status with batch info if available
            batch_info = getattr(self, '_proofreading_batch_info', None)
            if batch_info:
                self.progress_label.config(text=f"AI Proofreading... {batch_info} | Elapsed: {elapsed_min}:{elapsed_sec:02d}")
            else:
                self.progress_label.config(text=f"AI Proofreading... | Elapsed: {elapsed_min}:{elapsed_sec:02d}")
            return
        
        # Check if we have detailed progress info
        audio_pos = getattr(self, 'last_audio_position', None)
        audio_dur = getattr(self, 'last_audio_duration', None)
        progress_val = getattr(self, 'last_progress_value', None)
        
        if audio_pos and audio_dur and audio_pos > 0:
            # Calculate speed and ETA
            processing_speed = audio_pos / time_elapsed if time_elapsed > 0 else 0
            remaining_audio = audio_dur - audio_pos
            eta_seconds = remaining_audio / processing_speed if processing_speed > 0 else 0
            
            # Format ETA
            eta_min = int(eta_seconds // 60)
            eta_sec = int(eta_seconds % 60)
            
            # Calculate percentage
            percent = int((audio_pos / audio_dur) * 100) if audio_dur > 0 else 0
            
            # Format: Transcribing... 15% (47s / 910s) | Elapsed: 2:21 | ETA: 27:52 | Speed: 1.0x
            progress_text = (f"Transcribing... {percent}% ({audio_pos:.0f}s / {audio_dur:.0f}s) | "
                           f"Elapsed: {elapsed_min}:{elapsed_sec:02d} | "
                           f"ETA: {eta_min}:{eta_sec:02d} | "
                           f"Speed: {processing_speed:.1f}x")
            self.progress_label.config(text=progress_text)
        elif progress_val is not None:
            # Just show percentage and elapsed
            self.progress_label.config(text=f"Transcribing... {progress_val:.0f}% | Elapsed: {elapsed_min}:{elapsed_sec:02d}")
        else:
            # Just show elapsed time
            current_text = self.progress_label.cget('text')
            if ' | Elapsed:' in current_text:
                base_text = current_text.split(' | Elapsed:')[0]
            else:
                base_text = current_text if current_text else "Transcribing..."
            
            if base_text:
                self.progress_label.config(text=f"{base_text} | Elapsed: {elapsed_min}:{elapsed_sec:02d}")
    
    def update_progress_with_time(self, value, status_text):
        """Update progress bar and label with elapsed time"""
        import time
        self.progress['value'] = value
        
        if self.transcription_start_time:
            elapsed = time.time() - self.transcription_start_time
            minutes = int(elapsed // 60)
            seconds = int(elapsed % 60)
            self.progress_label.config(text=f"{status_text} | Elapsed: {minutes}:{seconds:02d}")
        else:
            self.progress_label.config(text=status_text)
        
    def run_transcription(self):
        """Run the actual transcription"""
        try:
            import whisper
            
            video_file = self.video_path.get()
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
                'verbose': True  # Keep verbose for progress
            }
            
            # Add context/prompt if provided
            context = self.context_prompt.get().strip()
            if context:
                transcribe_options['initial_prompt'] = context
                self.log(f"Using context: {context[:100]}{'...' if len(context) > 100 else ''}\n")
            
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
                self.progress_label.config(text="Cancelled")
                messagebox.showinfo("Cancelled", "Transcription was cancelled.")
                return
            finally:
                sys.stdout = old_stdout
                sys.stderr = old_stderr
            
            if self.stop_requested:
                self.log("\n❌ Transcription cancelled")
                self.progress_label.config(text="Cancelled")
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
                    self.progress_label.config(text="AI Proofreading...")
                    
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
                    self.progress_label.config(text=f"✓ Done in {minutes}:{seconds:02d}")
                else:
                    self.progress_label.config(text="✓ Done")
                
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
                self.progress_label.config(text="Complete (file not found)")
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
            self.processing = False
            self.stop_requested = False
            self.timer_running = False
            self._proofreading_mode = False  # Reset proofreading mode
            self.progress.stop_animation()  # Stop gradient animation
            self.process_btn.config(state='normal')
            self.stop_btn.config(state='disabled')
            if not self.stop_requested:
                # Keep progress at 100% if completed, otherwise reset
                if self.progress['value'] < 100:
                    self.progress['value'] = 0
                    self.progress_label.config(text="")
                else:
                    # Show final elapsed time
                    if self.transcription_start_time:
                        import time
                        total_time = time.time() - self.transcription_start_time
                        minutes = int(total_time // 60)
                        seconds = int(total_time % 60)
                        current_text = self.progress_label.cget('text')
                        if 'Done' in current_text or 'Complete' in current_text:
                            self.progress_label.config(text=f"✓ Done in {minutes}:{seconds:02d}")

def main():
    root = tk.Tk()
    app = WhisperGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
