# **Excel Image Inserter & Compressor**  
This project was developed with AI assistance using ChatGPT and other AI tools.  

![AI-Assisted](https://img.shields.io/badge/AI-Assisted-blue?style=for-the-badge&logo=ai)  

## 📌 Overview  
**Excel Duplicate Row Remover** is a tool designed to remove duplicate rows from an Excel sheet. It prompts the user to specify a column containing quantity values, then removes duplicates based on all columns except for the specified quantity column, summing the values in that column. It is rather slow and inefficient but gets the job done

## ✨ Features  
- ✅ Removes duplicate rows while summing the quantity values
- ✅ Allows the user to select the column containing quantities
- ✅ Provides a progress bar for real-time updates
- ✅ GUI-based interface for user-friendly operation 
- ✅ Provides a progress bar for real-time updates  
- ✅ Compatible with .xlsx Excel files

## 🖥️ Installation  

1. **Clone the Repository:**  
   ```bash
   git clone https://github.com/Devbones/Excel-Duplicate-Row-Removal.git
   cd Excel-Duplicate-Row-Removal
   ```
2. **Install Dependencies:**  
   Ensure you have Python installed, then install the required libraries:  
   ```bash
   pip install -r requirements.txt
   ```
3. **Run the Application:**  
   ```bash
   python main.py
   ```

## 🚀 Usage  

1. **Select an Excel file** containing the data. 
2. **Enter the column letter** where quantities are stored (default is column "J").
3. **The script will process the data**, removing duplicate rows based on all columns except the quantity column and summing the quantity values. 
4. **View the progress** through the provided progress bar.  
5. **Enable/Disable Image Compression** as needed.  
6. **Once completed, the file will be saved** with the suffix _summed.xlsx added to the original file name.

## 📜 License  
© 2025 Artur Kuśmirek. All rights reserved.  
