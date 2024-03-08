# ght2txt_converter

This is a simple tool to convert **Glaucoma Hemifield Test data from Zeiss Humphrey Field Analyzer (HFA)** PDF files to Dataframe and extract meaningful information. The script is for converting PDF files to EXCEL files and extracting patient information from text files. The text files are generated from PDF files by using OCR software. The text files contain patient information and the PDF files contain Glaucoma Hemifield Test results. The script will create a dataframe and save it as an excel file. The script will also crop the PDF files and save them as JPG files. This repository will not contain  `.pdf` file, however, there will be `.txt` file converted and anonymized from the `.pdf` file.

## Installation

To install the required packages, run the following command:

```bash
conda env create -f environment.yml
conda activate ghtpdf
```

This will install the required packages, including `PyPDF2` and `pandas`.

## Usage

To convert a PDF file to excel file, run the following command:

```bash
python ght2pdf.py
```
