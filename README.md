# File Sequence Checker GUI

This PowerShell script provides a graphical user interface (GUI) to search for missing files in a numerical or alphabetical sequence. It is designed to help users quickly identify gaps in file sequences, such as missing files in a RAR archive, documents, or any other type of files.

## Features

- **User-Friendly Interface**: A simple and intuitive GUI built using Windows Forms.
- **Sequence Checking**: Allows users to specify a folder and file pattern to detect missing files in a sequence.
- **Customizable Search Parameters**: Supports various sequence formats, including numbers and letters.
- **Instant Results**: Quickly highlights missing files directly within the interface.

## Requirements

- Windows OS with PowerShell 5.0 or later.
- .NET Framework to support Windows Forms.

## How to Use

1. **Run the Script**: Open PowerShell and run the script `file-sequence-checker-gui.ps1`.
2. **Specify Search Criteria**:
   - Choose the directory containing the files.
   - Enter the expected file sequence pattern (e.g., `file001` to `file100`).
3. **Check Sequence**: Click the "Search" button to start checking the sequence.
4. **View Results**: The missing files will be displayed in the results section of the GUI.

## Installation

No installation is required. Simply download the script and execute it using PowerShell.

## Troubleshooting

- Ensure that PowerShell is running with sufficient permissions.
- If the GUI does not appear, make sure all required assemblies are accessible.

## License

This tool is open-source and free to use. Feel free to modify and distribute as needed.

## Contributing

Contributions are welcome! If you find any issues or have suggestions for improvements, please submit a pull request or open an issue.
