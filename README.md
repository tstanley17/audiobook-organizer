# audiobook-organizer
Simple application to scan a directory for audiobook files. If metadata is missing, it will search for the metadata and allow you to select the files to write the metadata to the files. Write files out to a specific directory structure,previewing the changes prior to committing them.

Simple Docker Compose:

services: audiobook-organizer: image: tstanley17/audiobook-organizer:latest ports: - "5800:5800" volumes: - /home/plexadmin/appdata/downloads:/audiobooks - /home/plexadmin/appdata/audiobook-organizer:/config environment: - USER_ID=1000 - GROUP_ID=1000

Future:

Currently working on manual entry for searching metadata for specific audiobook files.
