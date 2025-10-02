# Project Title: AI-Powered Medical Data Extraction

**A collaborative project between Roskilde University (RUC) and the Technical University of Denmark (DTU) to automate the extraction of patient data from the Epic software.**

---

## Table of Contents
- [Project Overview](#project-overview)
- [For End-Users](#for-end-users)
- [For Developers](#for-developers)
- [Dependencies](#dependencies)
- [Project Status](#project-status)
- [License](#license)

---

## Project Overview

### The Challenge
Medical professionals and researchers at Roskilde University (RUC) and the Technical University of Denmark (DTU) frequently need to extract large volumes of patient data from the Epic electronic health record system for clinical studies. This process, when done manually, is repetitive, time-consuming, and prone to data entry errors.

### Our Solution
This project provides an intelligent automation tool that mimics human interaction to extract data from the Epic software. It navigates the Epic interface, identifies the correct patient records, and copies the required information into organized Excel files. The project is conducted with full approval from the Danish Board of Health, ensuring data privacy and compliance are respected.

### Business Value
- **Increased Efficiency:** Drastically reduces the time spent on manual data collection, allowing staff to perform higher-value tasks.
- **Improved Data Accuracy:** Eliminates human error from manual copy-pasting, ensuring the reliability of data for research and analysis.
- **Scalability:** The modular design allows the tool to be easily extended to extract new types of data as research needs evolve.
- **Cost Savings:** Reduces the person-hours required for data extraction, leading to significant operational cost savings.

---

## For End-Users

### How to Use the Software
To run the data extraction tool, please follow these steps carefully:

1. **Prepare Your CPR Numbers:**
   - Create an Excel file named `cprs.xlsx`.
   - In this file, create a single column with the header "cpr".
   - List all the patient CPR numbers you want to process in this column.
   ![example Excel file](images/example-of-cpr-nr-in-excel.jpg)

2. **Open Epic:**
   - Before running the software, make sure the Epic application is open on your computer.
   - Navigate to the correct starting page within Epic as shown in the image below. This is crucial for the software to find the correct buttons and fields.
   ![example Epic software](images/front-page-of-epic.jpg)

3. **Run the Software:**
   - Open your command line or terminal.
   - Navigate to the `master` folder of this project.
   - Run the command: `python open_extract_close_patient.py`

The software will then begin processing the patients one by one.

---

## For Developers

### Architecture Overview
The core of this project is a Robotic Process Automation (RPA) script that uses GUI automation to drive the Epic desktop application.

- **Orchestrator (`open_extract_close_patient.py`):** This is the main script. It reads a list of CPR numbers from `cprs.xlsx`, and for each patient, it automates the process of searching, opening the patient's file, and closing it after extraction is complete.

- **Extractor Modules:** The actual data extraction is handled by a series of "plug-in" modules located in their respective subdirectories (e.g., `blood_test`, `medicin`). The orchestrator calls these modules in sequence. Each module is responsible for navigating to a specific section of the Epic application and extracting the relevant data, typically using screen scraping (OCR) or by automating data export features if available.

- **GUI Automation:** The automation is powered by libraries like `PyAutoGUI` and `opencv-python`. It relies on finding specific images (buttons, icons, text fields) on the screen to navigate the application. These template images are stored in the `master/images` directory.

#### Flowchart
![flow-chart](images/extract-data-flow-chart.png)

### Project Structure
```
.
├── master/
│   ├── blood_test/            # Module for extracting blood test data.
│   ├── diagnose_liste/        # Module for extracting diagnosis lists.
│   ├── medicin/               # Module for extracting medication lists.
│   ├── ...                    # Other data extraction modules.
│   ├── images/                # Contains all the .jpg/.png image snippets used by PyAutoGUI to find and click UI elements.
│   ├── log/                   # Directory for log files (if any).
│   ├── cprs.xlsx              # (Input) A list of CPR numbers to be processed.
│   ├── global_functions.py    # (Optional) Shared functions used across multiple modules.
│   ├── open_extract_close_patient.py  # Main orchestration script that runs the end-to-end process.
│   └── requirements.txt       # A list of all Python packages required to run the project.
└── README.md                  # This file, providing project documentation.
```

### Getting Started for Developers
1.  **Clone the Repository:**
    `git clone <repository-url>`
2.  **Set Up a Virtual Environment (Recommended):**
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
    ```
3.  **Install Dependencies:**
    - Navigate to the `master` folder.
    - Run `pip install -r requirements.txt`.

### How to Contribute
The most common way to contribute is by adding a new module to extract a new piece of data.

1.  **Create a Module Directory:**
    - Inside the `master` directory, create a new folder with a descriptive name for the data you're extracting (e.g., `allergy_list`).

2.  **Develop Your Extraction Logic:**
    - Inside your new directory, create a Python script (e.g., `extract_allergies.py`).
    - This script should contain a primary function (e.g., `extract(cpr_nr)`) that performs the GUI automation to find and save the data.
    - Use the `images` folder to store any new UI screenshots needed for your automation.

3.  **Integrate into the Main Orchestrator:**
    - Open `master/open_extract_close_patient.py`.
    - Import your new function at the top of the file.
    - Find the main loop that processes each patient and add a call to your function in the appropriate sequence.
    ```python
    # ... inside the main processing loop ...

    # Call the existing blood test module
    from blood_test.extract_data_from_columns import extract_blood_test_data_from_columns
    extract_blood_test_data_from_columns(cpr_nr)
    time.sleep(2)

    # Add your new module call here
    from allergy_list.extract_allergies import extract
    extract(cpr_nr)
    time.sleep(2)
    ```

4.  **Update Requirements:**
    - If your new module requires any Python packages not already in `master/requirements.txt`, add them to the file.

---

## Dependencies
This project relies on a number of external Python libraries. Below is a list of the key dependencies. Please see the `master/requirements.txt` file for a complete list.

*   **GUI Automation:** `PyAutoGUI`, `opencv-python`, `Pillow`
*   **Data Handling:** `pandas`, `openpyxl`, `numpy`
*   **File Extraction:** `textract`, `pdfminer.six`, `python-docx`
*   **OCR:** `pytesseract`

**Note on Dependencies:** This project has a large number of dependencies, which can introduce security risks and maintenance overhead. Please use standard libraries where possible and carefully consider the need for any new external packages.

---

## Project Status
This project is actively being used and developed.

---

## License
[Specify a license, e.g., MIT, GPL, etc. If none, state that it is a proprietary/closed-source project.]