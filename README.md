# ATT&CKforge

![ATT&CKforge Banner](https://img.shields.io/badge/MITRE-ATT%26CKforge-blue)

**ATT&CKforge** is an interactive Python tool that generates professionally formatted Excel spreadsheets of MITRE ATT&CK matrices for any platform. Fetch real-time framework data, select specific matrices and platforms through an intuitive menu system, and create clean, organized references for cybersecurity professionals.
<p align="center">
  <img src="https://github.com/user-attachments/assets/e080fcd2-f087-41e7-8681-d54b8522a9ce">
</p>

## 🛡️ Features

- **Interactive Selection Menu**: Choose from Enterprise, Mobile, or ICS (Industrial Control Systems) matrices
- **Multi-Platform Support**: Generate matrices for Windows, macOS, Linux, Android, and more
- **Real-Time Data**: Pull the latest data directly from the official MITRE ATT&CK repository
- **Professional Formatting**: Create clean Excel spreadsheets with proper headers, borders, and organization
- **Comprehensive Structure**: Matrices include techniques, subtechniques, and direct links to MITRE references
- **Bulk Processing**: Generate multiple matrices in a single run
- **User-Friendly Experience**: Clear navigation and progress updates throughout the process

<p align="center">
  <img src="https://github.com/user-attachments/assets/4b6562a8-29ef-4fdf-b842-8dbdae83e778">
</p>


## 🚀 Installation

```bash
# Clone the repository
git clone https://github.com/usualdork/att-ckforge.git
cd att-ckforge

# Install required packages
pip install -r requirements.txt
```

## 📋 Requirements

- Python 3.7+
- Required packages:
    - requests
    - openpyxl
 
## 💻 Usage

```bash
# Run the interactive tool
python attckforge.py
```

Follow the on-screen prompts to:
  1. Select a matrix type (Enterprise, Mobile, or ICS)
  2. Choose one or more platforms to generate
  3. Wait for the Excel files to be created

## 📊 Example Output
ATT&CKforge generates Excel files with the following structure:
 -  Clear matrix title and header
 -   Tactics as section headers
 -   Techniques with their reference links
 -   Subtechniques organized under parent techniques
 -   Consistent formatting and borders
  
Files are saved as: MITRE_ATT&CK_<framework>_<platform>_<date>.xlsx

## 🔧 Advanced Usage

```bash
# Import the fetcher class into your own scripts
from attckforge import MitreAttackMatrixFetcher

# Create a fetcher instance
fetcher = MitreAttackMatrixFetcher()

# Process specific platforms programmatically
fetcher.process_selected_platforms(['Windows', 'macOS'])

# Or process all available platforms
fetcher.process_all_platforms()
```

## 🤝 Contributing

1. Fork the repository
2. Create your feature branch (git checkout -b feature/amazing-feature)
3. Commit your changes (git commit -m 'Add some amazing feature')
4. Push to the branch (git push origin feature/amazing-feature)
5. Open a Pull Request

## 📝 License
This project is licensed under the MIT License - see the LICENSE file for details.

## ✨Acknowledgements
- MITRE ATT&CK® is a registered trademark of The MITRE Corporation.
- Thanks to the MITRE team for maintaining the ATT&CK framework and making the data publicly available.
- Inspired by *sduff/mitre_attack_csv* for data structure reference.

## 📬 Contact
If you have any questions or feedback, please open an issue or contact me **@usualdork**
