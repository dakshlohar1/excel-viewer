# Excel Preview and Editor

## Features

- Support for multiple Excel formats (CSV, XLSX, XLS, ODS, etc.)
- Preview multiple sheets in a single workbook
- Edit cells directly in the browser
- Export edited sheets

## Prerequisites

- Node.js (v16 or later)
- npm or yarn

## Installation

1. Clone the repository

```bash
git clone https://your-repository-url.git
cd excel-preview-editor
```

2. Install dependencies

```bash
npm install
# or
yarn install
```

3. Run the development server

```bash
npm run dev
# or
yarn dev
```

## Dependencies

- React
- xlsx (for Excel file parsing)
- Shadcn UI (for components)
- Tailwind CSS (for styling)

## Usage

1. Click on the file input to upload an Excel file
2. Navigate between sheets using tabs
3. Click on any cell to edit its content
4. Use the "Export Sheet" button to download the modified sheet

## Supported Formats

- .csv
- .xlsx
- .xls
- .ods
- .xlsm
- .xlsb
- .xml

## Customization

You can modify the `ExcelViewer` component to add more features like:

- More advanced cell formatting
- Formula support
- Additional export options

## Troubleshooting

- Ensure you have the latest version of Node.js
- Check console for any parsing errors
- Verify file compatibility
