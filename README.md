# AI-Prompt-Workflow-Utilities

Some utils for Windows to aid in the workflow when working with AI Prompts.

## Overview

This repository contains utilities to enhance your workflow when working with AI prompts. The provided PowerShell scripts help in managing and converting documents to Markdown format, making it easier to create and organize AI-related content.

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

## Requirements

- Windows operating system
- PowerShell 5.0 or later
- Administrator privileges to run the setup and removal scripts
- [Pandoc](https://pandoc.org/installing.html) installed and available in the system PATH for document conversion

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.