# PowerPoint Automation Project

Automate extraction of images from a PowerPoint presentation, perform reverse image searches, and add URLs as footers to slides.

## Table of Contents
- [Setup](#setup)
  - [Clone the Repository](#clone-the-repository)
  - [Create a Virtual Environment](#create-a-virtual-environment)
  - [Activate the Virtual Environment](#activate-the-virtual-environment)
  - [Install Dependencies](#install-dependencies)
  - [Download and Set Up External Tools](#download-and-set-up-external-tools)
  - [Run the Main Script](#run-the-main-script)
- [Dependencies](#dependencies)
- [Usage](#usage)
- [Scripts](#scripts)
  - [extract_images.py](#extract_imagespy)
  - [selenium_automation.py](#selenium_automationpy)
  - [add_footer.py](#add_footerpy)
  - [main_script.py](#main_scriptpy)
- [Contributing](#contributing)
- [License](#license)

## Setup

Clone the Repository:

```bash
git clone https://github.com/your-username/powerpoint_citation_automated.git
cd powerpoint-automation
```

## Create a Virtual Environment:

```bash
python -m venv venv
```

## Activate the Virtual Environment:
On Windows:

```bash
.\venv\Scripts\activate
```
On macOS/Linux:

```bash
source venv/bin/activate
```

## Install Dependencies:

```bash
pip install -r requirements.txt
```

## Download and Set Up External Tools:
For Selenium, ensure ChromeDriver is installed or use chromedriver-autoinstaller as specified in requirements.txt.

## Run the Main Script:

```bash
python main_script.py
```

## Dependencies

Python (3.6 or higher)
See requirements.txt for Python packages.

## Usage

Place the PowerPoint file to be processed in the root directory. Run the main_script.py script to automate the image extraction, URL extraction, and adding URLs as footers to slides.

## Scripts

- extract_images.py: Extracts images from a PowerPoint presentation and saves them to the "output_images" folder.
- selenium_automation.py: Performs web automation to search for images and extract URLs using Selenium.
- add_footer.py: Adds URLs as footers to slides in a PowerPoint presentation based on CSV data.
- main_script.py: Orchestrates the execution of the three scripts in sequence.

## Contributing

Feel free to contribute by submitting issues, feature requests, or pull requests.

## License
This project is licensed under the MIT License - see the LICENSE file for details.
