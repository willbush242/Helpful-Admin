# File-conversion tools

## Function

Convert M4A audio files to MP3 and Microsoft Publisher documents to PDF.

## Overview

Personal tools for converting files individually or in batches. The collection includes Python scripts for audio and document conversion, plus a guide for a Publisher-to-PDF Windows application.

## Contents

Convert_m4a_to_MP3 - Convert a folder of M4A audio files to MP3.

Publisher_Files_to_PDF - Batch-convert Microsoft Publisher documents into PDFs.

Publisher_to_PDF - A Windows application for converting one or more Publisher documents to PDF without editing folder paths or running a Python script manually.

## Context and contribution

I developed these tools to automate practical tasks, including converting audio files from Microsoft Clipchamp and exporting Publisher documents to PDF.

Developed with assistance from ChatGPT. I defined the requirements, adapted and tested the code, checked outputs for errors, and debugged issues.

## How to use
Open the relevant folder and follow its README. The audio script requires imageio-ffmpeg; the Publisher Python script requires Windows, desktop Microsoft Publisher and pywin32. Set the input paths before running. The audio converter overwrites existing destination MP3s and deletes each M4A original after a successful conversion with a non-empty output. Choose unused PDF destination names if you want to retain earlier exports.

The Publisher_to_PDF folder currently contains documentation only; Publisher_to_PDF.exe is not included in this repository.
