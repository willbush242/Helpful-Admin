# Fit_Photos_On_A4.py

## Function

Applies EXIF orientation, preserves aspect ratios and chooses grids of up to four photographs using preferred sizes and displayed area. Creates temporary JPEGs for Word without changing source images.

## Overview

Arrange photographs in a Word document for printing.

## Context and contribution

I wrote this script for a personal project to automate a practical task.

Developed with assistance from ChatGPT. I defined the requirements, adapted and tested the code, checked outputs for errors, and debugged issues.

## How to use

Install Pillow and python-docx. Set PHOTO_FOLDER and OUTPUT_FILE; adjust page dimensions, margins, GAP_CM and target sizes if desired. Place supported JPG, PNG, BMP, TIFF or WebP images directly in PHOTO_FOLDER. Run `python Fit_Photos_On_A4.py`. Writes photo_layout.docx by default.
