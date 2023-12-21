# PDF Index Creator

## Overview
PDF Index Creator is a dedicated Windows application for creating and managing PDF document indices. With a PySide6-powered graphical user interface, this tool provides an intuitive environment for users to add, edit, and organize index entries effectively.

![Image](https://github.com/kamalexandre/PDF-Index-Creator/blob/main/SampleView.png)

## Key Features
- **Interactive PDF Viewer**: Open and review PDFs directly within the application.
- **Index Management**: Intuitive addition, editing, and deletion of index entries.
- **Search Functionality**: Efficiently find terms with a built-in search feature.
- **Note-taking**: Annotate index entries with personalized notes and comments.
- **IDX File Generation**: Create `.idx` files compatible with LaTeX documentation.

## Windows Compatibility
The application is optimized for Windows OS and is undergoing fixes to support macOS and Linux in the future. Your contributions are appreciated!

## Installation
Follow these steps to set up the PDF Index Creator:

#### Step 1: Clone the Repository

Clone the repository to your local machine:

```bash
git clone https://github.com/kamalexandre/PDF-Index-Creator.git
```

#### Step 2: Navigate to the Application Directory

Change to the project directory:

```bash
cd PDF-Index-Creator
```

#### Step 3: Create a Virtual Environment

Create a virtual environment for the project:

```bash
python -m venv PDF-Index-Creator
```

#### Step 4: Activate the Virtual Environment

Activate the virtual environment:

```bash
PDF-Index-Creator\Scripts\activate
```

If you encounter any errors in PowerShell, try adjusting the execution policy:

```bash
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### Step 5: Install Dependencies

Install the necessary Python packages:

```bash
pip install -r requirements.txt
```

#### Step 6: Run the Application

Start the application:

```bash
python PDF-Index-Creator.py
```

### Optional: Compiling to a Windows Executable

To create a standalone Windows executable:

1. Install PyInstaller and openpyxl:

    ```bash
    pip install --upgrade pyinstaller openpyxl
    ```

2. Compile the application:

    ```bash
    pyinstaller main.spec
    ```

    The executable will be located in the `dist` directory.

## Note

PDF Index Creator is in passive development and done in my spare time, it may contain bugs. Your feedback and reports on any issues you encounter are invaluable.

## Contributing

Contributions are what make the open-source community such an inspiring place to learn, innovate, and create. Any contributions to improve PDF Index Creator or extend its compatibility are **greatly appreciated**.


## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contact

Akam Okokon - [@kam_alexandre](https://twitter.com/kam_alexandre)

Project Link: [PDF-Index-Creator](https://github.com/yourusername/pdf-index-creator)
