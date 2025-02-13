# MegaTool

MegaTool is a Python program designed to facilitate efficient data transfer and management across Windows applications to improve workflow. It aims to streamline operations such as file copying, moving, deleting, and basic Excel automation tasks.

## Features

- **File Copying**: Allows you to copy files from one directory to another.
- **File Moving**: Move files seamlessly between directories.
- **File Deletion**: Delete unnecessary files with ease.
- **Excel Automation**: Perform basic operations with Excel files, such as reading cell values.

## Requirements

- Python 3.x
- `pywin32` package for Windows COM support

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/megatoool.git
   ```
2. Install the required Python packages:
   ```bash
   pip install pywin32
   ```

## Usage

```python
from mega_tool import MegaTool

tool = MegaTool()

# Example usage
tool.copy_file('path/to/source/file.txt', 'path/to/destination/')
tool.move_file('path/to/source/file.txt', 'path/to/destination/')
tool.delete_file('path/to/destination/file.txt')
tool.excel_automation('path/to/excel/file.xlsx')
```

This code initializes the `MegaTool` class and demonstrates how to use its methods for file management and Excel automation.

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request for any improvements or bug fixes.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.