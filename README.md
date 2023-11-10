# docx-templater

## Description
This project is a Python-based application that automates the process of generating documents from templates. It uses PySimpleGUI for the graphical user interface, SQLite for database management, and the DocxTemplate library for handling DOCX templates. The tool allows users to input data, choose templates, and generate customized documents efficiently.

## Features
- User-friendly GUI for easy interaction.
- Ability to add, delete, and modify data keys in the database.
- Customizable document generation from DOCX templates.
- Efficient handling of file and database operations.

## Installation

### Prerequisites
- Python 3.x
- Pip (Python package manager)

### Dependencies
Install the required Python libraries using pip:
```bash
pip install PySimpleGUI
pip install python-docx
pip install docxtpl
```

### Cloning the Repository
Clone the repository to your local machine:
```bash
git clone https://github.com/mgsima/docx-templater.git
cd docx-templater
```

## Usage
1. **Starting the Application**: Run the application by executing the Python script.
   ```bash
   python main.py
   ```
2. **Configuring the Database**: The application uses an SQLite database. Make sure to configure the database according to your needs.

3. **Using the GUI**: The GUI is intuitive and user-friendly. Follow the on-screen instructions to select templates, input data, and generate documents.

4. **Customizing Templates**: Templates should be in the DOCX format and placed in the specified template directory.

## Contributing
Contributions to this project are welcome! Please follow these steps:
1. Fork the repository.
2. Create a new branch for your feature (`git checkout -b feature/AmazingFeature`).
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`).
4. Push to the branch (`git push origin feature/AmazingFeature`).
5. Open a Pull Request.

## License
This project is licensed under the MIT License - see the [LICENSE.md](LICENSE) file for details.

## Acknowledgments
- PySimpleGUI for the GUI framework.
- DocxTemplate for handling DOCX files.
- SQLite for database management.
