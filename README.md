# Final Whisper

**Professional Subtitle Generator with AI-Powered Transcription**

Final Whisper is a desktop application that automatically generates high-quality subtitles from video files using OpenAI's Whisper speech recognition model. Designed for film production workflows, it creates professionally formatted SRT subtitle files with intelligent text formatting and optional AI proofreading.

![Version](https://img.shields.io/github/v/release/Gerlif/final-whisper)
![Python](https://img.shields.io/badge/python-3.11-green)
![License](https://img.shields.io/badge/license-MIT-lightgrey)

---

## ✨ Features

### 🎯 Smart Subtitle Formatting
- **Intelligent text splitting** at natural sentence boundaries (periods, question marks, exclamation marks)
- **Orphan word prevention** - avoids leaving conjunctions and articles at the end of lines
- **Configurable character limits** (default: 40 chars/line, 2 lines max)
- **Automatic merging** of very short subtitles for better readability
- **Word-level timestamps** for precise subtitle timing

### 🤖 AI-Powered Transcription
- Multiple Whisper model support (tiny, base, small, medium, large, turbo)
- GPU acceleration for faster processing
- Multi-language support (optimized for English and Danish but supports all Whisper languages)
- Real-time progress tracking with animated UI
- **Anti-hallucination settings** to reduce repetitive/looping text errors

### 📁 Batch Processing
- **Process multiple files** at once with the "File(s)" button
- **Process entire folders** with the "Folder" button
- Sequential processing with progress tracking (e.g., "File 2/5")
- Stop button cancels entire batch

### 🔍 AI Proofreading (Optional)
- Integration with Claude (Anthropic API) for post-transcription corrections
- Fixes common transcription errors and improves subtitle quality
- Automatic retry with rate limit handling

---

## 📥 Installation

### Option 1: Download Pre-built EXE (Recommended)

1. Download `FinalWhisper.exe` from [Releases](https://github.com/Gerlif/final-whisper/releases/latest)
2. Install OpenAI Whisper:
   ```bash
   pip install openai-whisper
   ```
3. Run `FinalWhisper.exe`

> **Note:** The app will prompt you to set up GPU acceleration on first run if you have an NVIDIA GPU.

### Option 2: Run from Source

1. Clone the repository:
   ```bash
   git clone https://github.com/Gerlif/final-whisper.git
   cd final-whisper
   ```

2. Install dependencies:
   ```bash
   pip install openai-whisper sv-ttk Pillow
   ```

3. Run the application:
   ```bash
   python whisper_gui.py
   ```

---

## 🚀 Usage

1. **Select Input** - Choose a single file, multiple files, or an entire folder
2. **Choose Output Folder** - Where the SRT files will be saved
3. **Configure Settings** (optional):
   - Select Whisper model (larger = more accurate but slower)
   - Choose language
   - Adjust subtitle formatting rules
   - Enable AI proofreading (requires Anthropic API key)
4. **Click "Start Transcription"** - Watch the progress in real-time
5. **Done!** - Your SRT file(s) are ready to use

---

## ⚙️ Advanced Features

### Subtitle Formatting Options
- **Max characters per line**: Default 40
- **Max lines per subtitle**: Default 2
- **Minimum characters before split**: Prevents splitting too early in a sentence
- **Minimum subtitle length**: Merges very short subtitles with neighbors

### Anti-Hallucination Settings
Whisper can sometimes produce repetitive or looping text. These settings help prevent that:
- **Condition on previous text**: Toggle off to reduce looping hallucinations
- **No speech threshold**: Adjusts silence detection sensitivity (0.0-1.0)
- **Hallucination silence threshold**: Detects and handles silent hallucinations

### AI Proofreading
Enable optional AI proofreading to improve transcription accuracy:
1. Get an API key from [Anthropic](https://console.anthropic.com/)
2. Enable "AI proofreading" in the GUI
3. Enter your API key
4. Claude Sonnet will review and correct the transcription

The proofreading feature includes automatic retry with exponential backoff if rate limited.

---

## 🏗️ Building from Source

To build your own EXE:

```bash
python build_exe.py
```

The compiled EXE will be in the `dist/` folder.

---

## 🔧 Technical Details

### Smart Subtitle Algorithm

The subtitle formatting engine uses several intelligent rules:

1. **Sentence boundary detection** - Splits at natural pause points
2. **Orphan word prevention** - Moves conjunctions/articles to the next line
3. **Character limit enforcement** - Respects maximum line length
4. **Short subtitle merging** - Combines subtitles that are too brief
5. **Word-level timestamps** - Ensures accurate timing throughout

### Architecture

- **GUI Framework**: tkinter with sv-ttk theme
- **Transcription**: OpenAI Whisper
- **AI Proofreading**: Anthropic Claude API
- **Build System**: PyInstaller with custom spec file
- **Auto-Update**: GitHub Releases API with batch script installer

---

## 📦 Requirements

- **Python 3.11+** (for source installation)
- **OpenAI Whisper** and dependencies
- **GPU (Optional)**: NVIDIA GPU with CUDA support for faster processing
- **Anthropic API Key (Optional)**: For AI proofreading feature

---

## 🤖 AI Development Disclaimer

**This application was developed primarily using AI assistance** (Claude Code). While thoroughly tested, users should be aware that AI-generated code may contain unexpected behaviors or bugs. Please report any issues on the [GitHub Issues](https://github.com/Gerlif/final-whisper/issues) page.

---

## 🐛 Troubleshooting

**App won't start:**
- Ensure Python and Whisper are installed: `pip install openai-whisper`
- Check if antivirus is blocking the EXE

**Transcription is slow:**
- Use a smaller model (tiny, base, small)
- Enable GPU acceleration with CUDA-enabled PyTorch

**Out of memory:**
- Use a smaller Whisper model
- Close other applications
- Process shorter video segments

**Update won't install:**
- Run as administrator
- Check internet connection
- Manually download from [Releases](https://github.com/Gerlif/final-whisper/releases)

---

## 📝 License

MIT License - See LICENSE file for details

---

## 👏 Credits

- **OpenAI Whisper** - Speech recognition model
- **Anthropic Claude** - AI proofreading
- **sv-ttk** - Modern tkinter theme
- **PyInstaller** - EXE packaging

---

## 🔗 Links

- [Report Issues](https://github.com/Gerlif/final-whisper/issues)
- [Latest Release](https://github.com/Gerlif/final-whisper/releases/latest)
- [OpenAI Whisper](https://github.com/openai/whisper)
- [Anthropic Claude](https://www.anthropic.com/)

---

**Made with ❤️ by Final Film (and AI)**
