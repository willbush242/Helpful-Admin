# Convert_m4a_to_MP3.py

## Function

Recursively finds M4A files and calls FFmpeg using libmp3lame at quality setting 0. Deletes each original only after a successful conversion with a non-empty MP3.

## Overview

Convert a folder of M4A audio files to MP3.

## Context and contribution

I wrote this script to convert audio files from Microsoft Clipchamp.

Developed with assistance from ChatGPT. I defined the requirements, adapted and tested the code, checked outputs for errors, and debugged issues.

## How to use

Install imageio-ffmpeg. Set root_folder, then run `python Convert_m4a_to_MP3.py`. MP3 files are written beside the originals. Existing destination MP3 files are overwritten, and successfully converted M4A originals are deleted.
