# Section Break Add-in for Microsoft Word

An MS Word add-in that enables inserting section breaks — a functionality currently missing in the online version of Word. It allows you to add section breaks at the cursor position or at the end of the document.

## Features

- **Insert Section Break at Cursor:** Quickly insert a section break immediately after the current cursor position.
- **Insert Section Break at Document End:** Easily add a section break at the end of your document.

## Installation

### Prerequisites

- Microsoft Word 2016 or later.
- Office 365 subscription (for online Word).
- [Node.js](https://nodejs.org/) installed on your development machine (if you plan to modify or run the add-in locally).

### Steps

1. **Clone or Download the Repository:**

   ```bash
   git clone https://github.com/Alexandros-Gavriel/SectionBreak.git
   ```

   2. **Load the Add-in in Word:**

   - **For Word Online:**
     - Upload the `manifest.xml` file to a location accessible by Word Online (e.g., OneDrive).
     - Open Word Online.
     - Go to **Insert** > **Office Add-ins** > **Upload My Add-in**.
     - Select the `manifest.xml` file from your OneDrive.

   - **For Word Desktop:**
     - Open Word.
     - Go to **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.
     - Click **Add new catalog** and enter the folder path where the `manifest.xml` is located.
     - Check the box for **Show in Menu**.
     - Click **OK** to save settings and close the dialog.
     - Restart Word.
     - Go to **Insert** > **My Add-ins** > **Shared Folder**.
     - Select the **Section Break Add-in**.

## Usage

1. **Access the Add-in:**

   - Navigate to the **Insert** tab in Word.
   - Locate the **Section Breaks** group containing the add-in buttons.

2. **Insert Section Breaks:**

   - **At Cursor Position:**
     - Click on the **Cursor Position** button to insert a section break immediately after your cursor.
   - **At Document End:**
     - Click on the **End of Document** button to insert a section break at the end of your document.

## License

This project is licensed under the **GNU General Public License v3.0**. See the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please follow these steps:

1. Fork the repository.
2. Create a new branch for your feature or bug fix.
3. Commit your changes with clear messages.
4. Open a pull request describing your changes.

## Support

If you encounter any issues or have questions, please open an [issue](https://github.com/Alexandros-Gavriel/SectionBreak/issues) on GitHub.

## Acknowledgments

- Thanks to the open-source community for providing valuable resources and inspiration.
- [Office Add-ins Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)

---

*This add-in is not affiliated with or endorsed by Microsoft Corporation.*
