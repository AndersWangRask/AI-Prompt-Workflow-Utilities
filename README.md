# AI-Prompt-Workflow-Utilities

Some utils for Windows to aid in the workflow when working with AI Prompts.

## Overview

This repository contains utilities to enhance your workflow when working with AI prompts. The provided PowerShell scripts help in managing and converting documents to Markdown format, combining multiple Markdown files into a single DOCX, archiving files, and setting up the necessary environment.

## Scripts Included

1. **setup.ps1**: Sets up the environment by creating a symbolic link, adding context menu entries, and creating a SendTo shortcut.
2. **remove.ps1**: Cleans up the environment by removing the symbolic link, context menu entries, and SendTo shortcut created by `setup.ps1`.
3. **create-markdown.ps1**: Converts `.docx` and `.pdf` files to Markdown format (`.md`).
4. **create-docx.ps1**: Converts `.md` and `.pdf` files to DOCX format (`.docx`).
5. **combine-and-convert.ps1**: Combines multiple Markdown files and converts the result to a single DOCX file.
6. **combine-and-convert.bat**: A batch file that acts as an intermediary between the Windows SendTo menu and the combine-and-convert.ps1 script.
7. **archive.ps1**: Archives files in a specified directory.
8. **install-pandoc.ps1**: Downloads and installs the latest version of Pandoc, and adds it to the system PATH.

## Setup

To get started, you need to set up the environment by running the `setup.ps1` script. This script will:

1. Create a symbolic link in the `Program Files` directory pointing to this repository.
2. Add context menu entries for `.docx` and `.pdf` files to allow easy conversion to Markdown format using the `create-markdown.ps1` script.
3. Add context menu entries for `.md` and `.pdf` files to allow easy conversion to DOCX format using the `create-docx.ps1` script.
4. Create a SendTo shortcut for the combine-and-convert functionality.

### Running the Setup Script

1. Open PowerShell as an Administrator.
2. Navigate to the directory where this repository is located.
3. Run the setup script:

   ```powershell
   .\setup.ps1
   ```

This will create a symbolic link in `C:\Program Files\AI-Prompt-Workflow-Utilities`, add context menu entries, and create the SendTo shortcut.

## Removal

If you wish to remove the setup, you can run the `remove.ps1` script. This script will:

1. Remove the context menu entries created by `setup.ps1`.
2. Delete the symbolic link created in the `Program Files` directory.
3. Remove the SendTo shortcut.

### Running the Removal Script

1. Open PowerShell as an Administrator.
2. Navigate to the directory where this repository is located.
3. Run the removal script:

   ```powershell
   .\remove.ps1
   ```

This will clean up all changes made by the setup script, restoring your system to its previous state.

## Script Descriptions

### create-markdown.ps1

This script converts `.docx` and `.pdf` files to Markdown format (`.md`). It performs the following steps:

1. Checks if the provided file is a `.docx` or `.pdf` file.
2. If an `archive.ps1` script is found in the same directory, it runs the archive script to back up the current state of files.
3. Uses Pandoc to convert the file to Markdown format.
4. Outputs the result to the same directory with the `.md` extension.

### create-docx.ps1

This script converts `.md` and `.pdf` files to DOCX format (`.docx`). It performs the following steps:

1. Checks if the provided file is a `.md` or `.pdf` file.
2. If an `archive.ps1` script is found in the same directory, it runs the archive script to back up the current state of files.
3. Uses Pandoc to convert the file to DOCX format.
4. Outputs the result to the same directory with the `.docx` extension.

### combine-and-convert.ps1

This script combines multiple Markdown files and converts the result to DOCX format. It performs the following steps:

1. Accepts multiple file paths as input.
2. Sorts the input files alphabetically.
3. Combines the content of all input Markdown files into a single temporary file.
4. Uses the `create-docx.ps1` script to convert the combined Markdown to DOCX format.
5. Cleans up the temporary combined Markdown file.
6. Places the resulting DOCX file in the same directory as the first input file.

### combine-and-convert.bat

This batch file serves as an intermediary between the Windows SendTo menu and the `combine-and-convert.ps1` script. It:

1. Receives file paths from the SendTo menu.
2. Properly formats these paths for use with PowerShell.
3. Calls the `combine-and-convert.ps1` script with the formatted file paths.

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

## Using the Utilities

### Convert to Markdown or DOCX

After running the setup script, you can:

1. Right-click on a `.docx` or `.pdf` file.
2. Select "Create Markdown (AI Prompt)" from the context menu to convert to Markdown.

Or:

1. Right-click on a `.md` or `.pdf` file.
2. Select "Create DOCX (AI Prompt)" from the context menu to convert to DOCX.

### Combine and Convert to DOCX

To combine multiple Markdown files and convert them to a single DOCX:

1. Select multiple Markdown (.md) files in File Explorer.
2. Right-click on the selection.
3. Choose "Send to" > "Combine and Convert to DOCX" from the context menu.

The selected files will be combined in alphabetical order and converted to a single DOCX file. The resulting DOCX file will be placed in the same directory as the first selected Markdown file, with a filename based on the names of the first and last input files.

### Installing Pandoc

To install Pandoc:

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