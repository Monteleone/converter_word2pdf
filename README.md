
<p align="center">
  <img src="https://github.com/Monteleone/converter_word2pdf/blob/main/convert-word-document-into-pdf.png" width="250">
</p>




This Python code creates a GUI application that converts Microsoft Word documents to PDF format. Here's what it does and what the user needs to execute it:

**Functionality**:

-   The application allows users to drag and drop both Microsoft Word documents (.docx) and shortcuts (.lnk) onto the window to select them for conversion.
-   Once files are dropped, their names are displayed in a list.
-   Clicking the "Convert" button initiates the conversion process, and each selected Word file is converted to PDF format.
-   After conversion, a message appears indicating that the process is complete.

**User Requirements**:

1.  **Python Environment**: The user needs to have Python installed on their system.
2.  **Python Packages**: The user must install the following Python packages:
    -   PyQt5
    -   docx2pdf
    -   win32com (for handling .lnk shortcuts)
3.  **Icon Files**: The application expects to find the following icon and image files in the specified paths:
    -   converter-ico.ico
    -   convert-word-document-into-pdf.png
4.  **Microsoft Word Documents**: Users should have Microsoft Word documents (*.docx) that they want to convert to PDF format.

**Changes Required**:

1.  **Python Packages**: Ensure that PyQt5, docx2pdf, and win32com packages are installed using pip.
2.  **Icon/Image Paths**: Modify the paths for the icon and image files if they are located elsewhere.
3.  **Shell Link Handling (Optional)**: If handling .lnk shortcuts is not needed, the section of code related to resolving shortcuts can be removed.
4.  **GUI Geometry**: Optionally adjust the dimensions and position of the GUI window using the `setGeometry()` method.
5.  **Execution**: Run the script using Python to start the application.

By following these steps and making the necessary changes, the user can successfully execute the Word to PDF conversion application, including the ability to drag and drop both Word documents and their shortcuts.
