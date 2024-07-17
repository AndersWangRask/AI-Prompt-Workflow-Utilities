# AI-Prompt-Workflow-Utilities

Some utils for Windows to aid in the workflow when working with AI Prompts.

## Overview

This repository contains utilities to enhance your workflow when working with AI prompts. The provided PowerShell scripts help in managing and converting documents to Markdown format, archiving files, and setting up the necessary environment.

### Scripts Included

1. **setup.ps1**: Sets up the environment by creating a symbolic link and adding context menu entries.
2. **remove.ps1**: Cleans up the environment by removing the symbolic link and context menu entries created by `setup.ps1`.
3. **create-markdown.ps1**: Converts `.docx` and `.pdf` files to Markdown format (`.md`).
4. **archive.ps1**: Archives files in a specified directory.
5. **install-pandoc.ps1**: Downloads and installs the latest version of Pandoc, and adds it to the system PATH.

## Setup

To get started, you need to set up the environment by running the `setup.ps1` script. This script will:

1. Create a symbolic link in the `Program Files` directory pointing to this repository.
2. Add context menu entries for `.docx` and `.pdf` files to allow easy conversion to Markdown format using the `create-markdown.ps1` script.

### Running the Setup Script

1. Open PowerShell as an Administrator.
2. Navigate to the directory where this repository is located.
3. Run the setup script:

   ```powershell
   .\setup.ps1
   ```

This will create a symbolic link in `C:\Program Files\AI-Prompt-Workflow-Utilities` and add context menu entries for `.docx` and `.pdf` files. The context menu entry "Create Markdown (AI Prompt)" will allow you to right-click on a `.docx` or `.pdf` file and convert it to a Markdown file.

## Removal

If you wish to remove the setup, you can run the `remove.ps1` script. This script will:

1. Remove the context menu entries created by `setup.ps1`.
2. Delete the symbolic link created in the `Program Files` directory.

### Running the Removal Script

1. Open PowerShell as an Administrator.
2. Navigate to the directory where this repository is located.
3. Run the removal script:

   ```powershell
   .\remove.ps1
   ```

This will clean up the context menu entries and remove the symbolic link, restoring your system to its previous state.

## Script Descriptions

### create-markdown.ps1

This script converts `.docx` and `.pdf` files to Markdown format (`.md`). It performs the following steps:

1. Checks if the provided file is a `.docx` or `.pdf` file.
2. If an `archive.ps1` script is found in the same directory, it runs the archive script to back up the current state of files.
3. Uses Pandoc to convert the file to Markdown format.
4. Outputs the result to the same directory with the `.md` extension.

### archive.ps1

This script archives files in a specified directory. It performs the following steps:

1. Checks if an `Archive` sub-directory exists in the current directory; if not, it creates it.
2. Copies files from the current directory to the `Archive` sub-directory, naming them with a timestamp to avoid overwriting existing files.
3. Provides feedback on the archiving process.

### install-pandoc.ps1

This script downloads and installs the latest version of Pandoc, a universal document converter. It performs the following steps:

1. Queries the GitHub API to get the latest release information for Pandoc.
2. Downloads the appropriate installer for the current platform.
3. Launches the installer and waits for the installation to complete.
4. Verifies that `pandoc.exe` is installed in the expected location.
5. Adds `pandoc.exe` to the system PATH if it is not already present.

### Running the Pandoc Installer Script

1. Open PowerShell as an Administrator.
2. Navigate to the directory where this repository is located.
3. Run the Pandoc installer script:

   ```powershell
   .\install-pandoc.ps1
   ```

This will ensure that the latest version of Pandoc is downloaded, installed, and added to the system PATH.

## Requirements

- Windows operating system
- PowerShell 5.0 or later
- Administrator privileges to run the setup and removal scripts
- [Pandoc](https://pandoc.org/installing.html) for document conversion (install using `install-pandoc.ps1` script)

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.