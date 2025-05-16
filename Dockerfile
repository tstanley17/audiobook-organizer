FROM jlesage/baseimage-gui:ubuntu-22.04-v4.5.3

# Install system dependencies, including OpenGL libraries
RUN apt-get update && apt-get install -y \
    python3 \
    python3-pip \
    qt6-base-dev \
    libxcb-cursor0 \
    libxcb1 \
    libxcb-render0 \
    libxcb-shm0 \
    libgl1-mesa-glx \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Copy the Python app
RUN pip3 install PySide6==6.2.4 mutagen requests

# Copy the application code
COPY audiobook_organizer.py /app/audiobook_organizer.py

# Create startapp.sh with explicit DISPLAY setting
RUN echo '#!/bin/sh\nexport DISPLAY=:0\npython3 /app/audiobook_organizer.py' > /startapp.sh && chmod +x /startapp.sh

# Set the application name
RUN set-cont-env APP_NAME "Audiobook Organizer"