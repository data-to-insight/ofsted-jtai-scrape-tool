#!/bin/bash

# Install additional system-level dependencies/packages
sudo apt-get update    

# Install Python dependencies
pip install -r requirements.txt

# Ensure run permissions on scrape script
chmod +x ofsted_childrens_services_inspection_scrape.py

                                     
# Install the Python extension for Visual Studio Code
code --install-extension ms-python.python --force

