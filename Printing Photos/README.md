# Photo-printing tools

## Function

Arrange photographs into a Word document for printing, applying image orientation and preserving aspect ratios while choosing layouts of up to four photographs per page.

## Overview

A Python tool for preparing photo layouts on A4 pages. It uses configurable page dimensions, margins, gaps and preferred photo sizes to reduce manual arrangement in Word.

## Contents

Fit_Photos_On_A4 - Arrange photographs in a Word document for printing.

## Context and contribution

I developed this tool as a personal project to automate the preparation of photographs for printing.

Developed with assistance from ChatGPT. I defined the requirements, adapted and tested the code, checked outputs for errors, and debugged issues.

## How to use

Open the Fit_Photos_On_A4 folder and follow its README. Install Pillow and python-docx, set PHOTO_FOLDER and OUTPUT_FILE, and adjust layout settings as needed. Place supported JPG, PNG, BMP, TIFF or WebP images directly in the input folder, then run the script. The default output is photo_layout.docx. Source images are not changed; review the resulting document before printing.
