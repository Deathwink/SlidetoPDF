# PDF Slide Generator

This is a simple Windows Forms application that generates a PDF file from an online slide deck (such as from an Adobe InDesign link). The application allows users to specify a URL, the number of slides to process, and combines the slides into a single PDF.

## Features

- **Generate PDF from Online Slides**: Enter the URL of an online slide deck.
- **Slide Selection**: Specify how many slides to process from the deck.
- **Combine Slides**: The application will automatically navigate through the slides, generate PDFs for each slide, and combine them into one PDF document.
- **Log Output**: The application logs progress during the process, including which slides are being processed and any errors that occur.
- **Landscape Orientation**: Generated PDFs are in landscape mode to fit the slide content.
  
## Requirements

- **PuppeteerSharp**: A .NET port of the headless Chrome Node library, used to capture the slides.
- **PdfSharp**: A library used to combine individual slide PDFs into a single PDF.
- **.NET Framework**: The application is built using Visual Basic .NET and requires .NET Framework to run.

## Setup

### Prerequisites

1. Ensure you have [PuppeteerSharp](https://github.com/hardkoded/puppeteer-sharp) and [PdfSharp](https://github.com/empira/PDFsharp) installed.
2. Download and install [Chromium](https://www.chromium.org/) for PuppeteerSharp to use. It will be automatically fetched when the app runs.

### Getting Started

1. **Clone the repository**:
   git clone https://github.com/your-repository-name/SlidePDFGenerator.git

2. **Build the project**:
   Open the solution in Visual Studio and build the project.

3. **Run the Application**:
   - Open the application (`.exe` file generated after building).
   - Enter the URL for the slide deck in the provided text box.
   - Select the number of slides to process.
   - Click **Start Processing**.

4. **Log Output**:
   - The application will log each slide it processes in the **Log** box.
   - Once processing is complete, you will see a success message, and the combined PDF will be saved to your Desktop.

## How It Works

1. The user enters a URL to the online slide deck in the text box.
2. The application loads the URL and navigates through the slides using PuppeteerSharp.
3. Each slide is saved as a separate PDF in landscape orientation.
4. After all slides are processed, the PDFs are combined into a single PDF document using PdfSharp.
5. A log of the entire process is displayed in the log box for easy tracking.
6. The generated PDF is saved on the Desktop in the `SlidesPDF` folder.

## Troubleshooting

- **No Slides Generated**: Ensure the URL you provided is correct and the slides are publicly accessible.
- **Error During PDF Combination**: Make sure all generated slide PDFs exist before combining them.
