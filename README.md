# Final Whisper

**Professional Subtitle Generator with AI-Powered Transcription**

Final Whisper is a desktop application that automatically generates high-quality subtitles from video files using OpenAI's Whisper speech recognition model. Designed for film production workflows, it creates professionally formatted SRT subtitle files with intelligent text formatting and optional AI proofreading.

![Version](https://img.shields.io/badge/version-1.04-blue)
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
- Multi-language support (optimized for Danish)
- Real-time progress tracking with animated UI

### 🔍 AI Proofreading (Optional)
- Integration with Claude (Anthropic API) for post-transcription corrections
- Fixes common transcription errors and improves subtitle quality

### 🎨 Modern User Interface
- Dark theme with purple gradient accents
- Collapsible settings sections
- Real-time log output
- Built-in update checker with **auto-install functionality**

### 🔄 Auto-Update System
- Checks for new versions on GitHub
- One-click download and install
- Automatic restart after update
- No manual intervention required

---

## 📥 Installation

### Option 1: Download Pre-built EXE (Recommended)

1. Download the latest `Final Whisper.exe` from [Releases](https://github.com/Gerlif/final-whisper/releases/latest)
2. Install Whisper dependencies:
   ```bash
   pip install openai-whisper
   ```
3. (Optional) For GPU acceleration:
   ```bash
   pip install torch torchvision torchaudio --index-url https://download.pytorch.org/whl/cu118
   ```
4. Run `Final Whisper.exe`

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

1. **Select Video File** - Choose your video file (supports any format Whisper can process)
2. **Choose Output Folder** - Where the SRT file will be saved
3. **Configure Settings** (optional):
   - Select Whisper model (larger = more accurate but slower)
   - Choose language
   - Adjust subtitle formatting rules
   - Enable AI proofreading (requires Anthropic API key)
4. **Click "Start Transcription"** - Watch the progress in real-time
5. **Done!** - Your SRT file is ready to use

---

## ⚙️ Advanced Features

### Subtitle Formatting Options
- **Max characters per line**: Default 40
- **Max lines per subtitle**: Default 2
- **Minimum characters before split**: Prevents splitting too early in a sentence
- **Minimum subtitle length**: Merges very short subtitles with neighbors

### AI Proofreading
Enable optional AI proofreading to improve transcription accuracy:
1. Get an API key from [Anthropic](https://console.anthropic.com/)
2. Enable "AI proofreading" in the GUI
3. Enter your API key
4. Claude Sonnet will review and correct the transcription

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
