# Publisher_to_PDF.exe

## Function

Opens a file-selection window, converts the selected Microsoft Publisher files to PDF using Publisher automation, and displays a completion message.

## Overview

A Windows application for converting one or more Publisher documents to PDF without editing folder paths or running a Python script manually.

## Context and contribution

I developed this application for a personal project to simplify converting Publisher documents to PDF.
Developed with assistance from ChatGPT. I defined the requirements, adapted and tested the code, checked outputs for errors, and debugged issues.

## How to use

Use Windows with desktop Microsoft Publisher installed. Save your Publisher documents, then double-click Publisher_to_PDF.exe. In the file-selection window, select one or more .pub files and click Open. Each PDF is saved alongside its source file with the same base filename and a .pdf extension. A completion message appears after the conversions finish. Cancel the selection window to exit without converting files.

The application processes only the selected files. Choose files whose destination PDF names are not already in use if you want to retain earlier exports. The supplied source prints selected file paths to the console when one is available.
