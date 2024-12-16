import { useState } from "react";
import ExcelJS from "exceljs";
import "./App.css";

function App() {
  const [workbook, setWorkbook] = useState(null);
  const [activeSheet, setActiveSheet] = useState(null);
  const [sheetData, setSheetData] = useState(null);
  const [mergedCells, setMergedCells] = useState([]);
  const [images, setImages] = useState({});
  const [fileInfo, setFileInfo] = useState(null);
  const [headers, setHeaders] = useState([]);
  const [cellStyles, setCellStyles] = useState({});

  const getExcelVersionInfo = (file) => {
    return {
      name: file.name,
      size: file.size,
      type: file.type,
      lastModified: new Date(file.lastModified).toLocaleString(),
      format: file.name.endsWith(".xlsx")
        ? "XLSX (Office 2007+ XML Format)"
        : "Unknown Format",
    };
  };

  const extractCellStyle = (cell) => {
    const style = {};

    if (cell.fill) {
      if (cell.fill.type === "pattern" && cell.fill.pattern === "solid") {
        if (cell.fill.fgColor?.argb) {
          const color = `#${cell.fill.fgColor.argb.substring(2)}`;
          style.backgroundColor = color;

          // If background is very light, ensure text is dark
          const r = parseInt(color.substr(1, 2), 16);
          const g = parseInt(color.substr(3, 2), 16);
          const b = parseInt(color.substr(5, 2), 16);
          if (r + g + b > 600) {
            // Light background
            style.color = "#000000";
          } else {
            // Dark background
            style.color = "#FFFFFF";
          }
        }
      }
    }

    if (cell.font) {
      style.fontWeight = cell.font.bold ? "bold" : "normal";
      style.fontStyle = cell.font.italic ? "italic" : "normal";
      style.textDecoration = cell.font.underline ? "underline" : "none";
      if (cell.font.color?.argb) {
        style.color = `#${cell.font.color.argb.substring(2)}`;
      }
      if (cell.font.size) {
        style.fontSize = `${cell.font.size}px`;
      }
    }

    if (cell.alignment) {
      style.textAlign = cell.alignment.horizontal || "left";
      style.verticalAlign = cell.alignment.vertical || "middle";
    }

    if (cell.border) {
      const getBorderStyle = (border) => {
        if (!border) return "none";
        const width = border.style === "thick" ? "2px" : "1px";
        const style = border.style === "dotted" ? "dotted" : "solid";
        const color = border.color?.argb
          ? `#${border.color.argb.substring(2)}`
          : "#000000";
        return `${width} ${style} ${color}`;
      };

      style.borderTop = getBorderStyle(cell.border.top);
      style.borderRight = getBorderStyle(cell.border.right);
      style.borderBottom = getBorderStyle(cell.border.bottom);
      style.borderLeft = getBorderStyle(cell.border.left);
    }

    // Add padding
    style.padding = "8px";

    return style;
  };

  const handleFileUpload = async (event) => {
    const file = event.target.files[0];
    if (file) {
      try {
        const versionInfo = getExcelVersionInfo(file);
        setFileInfo(versionInfo);

        const workbook = new ExcelJS.Workbook();
        const arrayBuffer = await file.arrayBuffer();
        await workbook.xlsx.load(arrayBuffer);
        setWorkbook(workbook);

        const worksheet = workbook.worksheets[0];
        setActiveSheet(worksheet.name);

        // Get merged cells
        const mergedCells = worksheet.mergeCells;
        const mergedCellsArray = [];
        if (mergedCells) {
          Object.keys(mergedCells).forEach((key) => {
            const range = key.split(":").map((ref) => {
              const match = ref.match(/([A-Z]+)(\d+)/);
              const col =
                match[1]
                  .split("")
                  .reduce(
                    (acc, char) => acc * 26 + char.charCodeAt(0) - 64,
                    0
                  ) - 1;
              const row = parseInt(match[2]) - 1;
              return { r: row, c: col };
            });
            mergedCellsArray.push({
              s: range[0],
              e: range[1],
            });
          });
        }
        setMergedCells(mergedCellsArray);

        // Extract images
        const imageMap = {};
        worksheet.getImages().forEach((image) => {
          const imageId = image.imageId;
          const imageData = workbook.getImage(imageId);
          const key = `${image.range.tl.row}-${image.range.tl.col}`;

          imageMap[key] = {
            data: imageData.buffer,
            extension: imageData.extension,
          };
        });
        setImages(imageMap);

        // Get headers and data with styles
        const headers = [];
        const rows = [];
        const styles = {};

        worksheet.eachRow((row, rowNumber) => {
          const rowData = [];
          row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            const value = cell.text || "";
            const styleKey = `${rowNumber - 1}-${colNumber - 1}`;
            styles[styleKey] = extractCellStyle(cell);

            if (rowNumber === 1) {
              headers[colNumber - 1] = value;
            } else {
              rowData[colNumber - 1] = value;
            }
          });

          // Fill in any missing cells in the row with empty strings
          const maxCols = worksheet.columnCount;
          while (rowData.length < maxCols) {
            rowData.push("");
          }

          if (rowNumber > 1) {
            rows[rowNumber - 2] = rowData;
          }
        });

        setHeaders(headers);
        setSheetData(rows);
        setCellStyles(styles);
      } catch (error) {
        console.error("Error loading workbook:", error);
      }
    }
  };

  const handleSheetChange = async (sheetName) => {
    try {
      const worksheet = workbook.getWorksheet(sheetName);
      setActiveSheet(worksheet.name);

      // Get merged cells
      const mergedCells = worksheet.mergeCells;
      const mergedCellsArray = [];
      if (mergedCells) {
        Object.keys(mergedCells).forEach((key) => {
          const range = key.split(":").map((ref) => {
            const match = ref.match(/([A-Z]+)(\d+)/);
            const col =
              match[1]
                .split("")
                .reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0) -
              1;
            const row = parseInt(match[2]) - 1;
            return { r: row, c: col };
          });
          mergedCellsArray.push({
            s: range[0],
            e: range[1],
          });
        });
      }
      setMergedCells(mergedCellsArray);

      // Extract images
      const imageMap = {};
      worksheet.getImages().forEach((image) => {
        const imageId = image.imageId;
        const imageData = workbook.getImage(imageId);
        const key = `${image.range.tl.row}-${image.range.tl.col}`;

        imageMap[key] = {
          data: imageData.buffer,
          extension: imageData.extension,
        };
      });
      setImages(imageMap);

      // Get headers and data with styles
      const headers = [];
      const rows = [];
      const styles = {};

      worksheet.eachRow((row, rowNumber) => {
        const rowData = [];
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const value = cell.text || "";
          const styleKey = `${rowNumber - 1}-${colNumber - 1}`;
          styles[styleKey] = extractCellStyle(cell);

          if (rowNumber === 1) {
            headers[colNumber - 1] = value;
          } else {
            rowData[colNumber - 1] = value;
          }
        });

        // Fill in any missing cells in the row with empty strings
        const maxCols = worksheet.columnCount;
        while (rowData.length < maxCols) {
          rowData.push("");
        }

        if (rowNumber > 1) {
          rows[rowNumber - 2] = rowData;
        }
      });

      setHeaders(headers);
      setSheetData(rows);
      setCellStyles(styles);
    } catch (error) {
      console.error("Error changing sheet:", error);
    }
  };

  // Check if a cell is part of a merged range
  const isMergedCell = (rowIndex, colIndex) => {
    return mergedCells.find(
      (range) =>
        rowIndex >= range.s.r &&
        rowIndex <= range.e.r &&
        colIndex >= range.s.c &&
        colIndex <= range.e.c
    );
  };

  // Get merged cell dimensions
  const getMergedCellDimensions = (rowIndex, colIndex) => {
    const mergedCell = mergedCells.find(
      (range) => rowIndex === range.s.r && colIndex === range.s.c
    );

    if (mergedCell) {
      return {
        rowSpan: mergedCell.e.r - mergedCell.s.r + 1,
        colSpan: mergedCell.e.c - mergedCell.s.c + 1,
      };
    }

    return null;
  };

  // Render cell content (text or image)
  const renderCellContent = (cell, rowIndex, colIndex) => {
    // Special handling for PS DESIGN logo in the first row
    if (
      rowIndex === 0 &&
      colIndex === 0 &&
      cell?.toLowerCase().includes("psdesign")
    ) {
      return (
        <div className="flex items-center">
          <img
            src="/logo.png"
            alt="PS DESIGN"
            className="h-8 w-auto"
            onError={(e) => {
              e.target.style.display = "none";
              e.target.nextSibling.style.display = "block";
            }}
          />
          <span className="ml-2" style={{ display: "none" }}>
            {cell}
          </span>
        </div>
      );
    }

    // Check for images
    const imageKey = `${rowIndex}-${colIndex}`;
    const image = images[imageKey];

    if (image && image.data) {
      try {
        const uint8Array = new Uint8Array(image.data);
        const binaryString = uint8Array.reduce(
          (str, byte) => str + String.fromCharCode(byte),
          ""
        );
        const base64Data = btoa(binaryString);
        const imgSrc = `data:image/${image.extension};base64,${base64Data}`;

        return (
          <img
            src={imgSrc}
            alt={`Cell content at ${rowIndex},${colIndex}`}
            className="max-w-full h-auto"
            style={{ maxHeight: "100px" }}
          />
        );
      } catch (error) {
        console.error(
          `Error processing image at ${rowIndex}-${colIndex}:`,
          error
        );
        return "";
      }
    }

    return cell || "";
  };

  return (
    <div className="App p-4">
      <div className="max-w-6xl mx-auto">
        {/* Header with logo placeholder */}
        <div className="flex items-center justify-between mb-6 border-b pb-4">
          <div className="text-2xl font-bold">Document Log Sheet</div>
          <div className="text-gray-500">PS DESIGN</div>
        </div>

        {/* File Upload */}
        <div className="mb-6">
          <input
            type="file"
            onChange={handleFileUpload}
            accept=".xlsx,.xls,.csv,.ods,.xlsm,.xlsb,.xml"
            className="block w-full text-sm text-gray-500
              file:mr-4 file:py-2 file:px-4
              file:rounded-full file:border-0
              file:text-sm file:font-semibold
              file:bg-blue-50 file:text-blue-700
              hover:file:bg-blue-100"
          />
        </div>

        {/* File Information */}
        {fileInfo && (
          <div className="mb-6 p-4 bg-gray-50 rounded-lg">
            <h3 className="font-semibold mb-2">File Information:</h3>
            <ul className="space-y-1 text-sm">
              <li>
                <strong>Name:</strong> {fileInfo.name}
              </li>
              <li>
                <strong>Format:</strong> {fileInfo.format}
              </li>
              <li>
                <strong>Size:</strong> {(fileInfo.size / 1024).toFixed(2)} KB
              </li>
              <li>
                <strong>Last Modified:</strong> {fileInfo.lastModified}
              </li>
              <li>
                <strong>Type:</strong> {fileInfo.type}
              </li>
            </ul>
          </div>
        )}

        {/* Sheet Tabs */}
        {workbook && (
          <div className="flex gap-2 mb-4 overflow-x-auto">
            {workbook.worksheets.map((worksheet) => (
              <button
                key={worksheet.name}
                onClick={() => handleSheetChange(worksheet.name)}
                className={`px-4 py-2 rounded whitespace-nowrap ${
                  activeSheet === worksheet.name
                    ? "bg-blue-500 text-white"
                    : "bg-gray-200 hover:bg-gray-300"
                }`}
              >
                {worksheet.name}
              </button>
            ))}
          </div>
        )}

        {/* Sheet Data */}
        {sheetData && (
          <div className="overflow-x-auto border rounded-lg shadow">
            <table className="min-w-full border-collapse">
              <thead>
                <tr>
                  {headers.map((header, index) => {
                    const headerStyle = cellStyles[`0-${index}`] || {};
                    return (
                      <th
                        key={index}
                        className="border font-semibold text-left"
                        style={headerStyle}
                      >
                        {header}
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody>
                {sheetData.map((row, rowIndex) => (
                  <tr key={rowIndex}>
                    {row.map((cell, colIndex) => {
                      const mergedCell = isMergedCell(rowIndex, colIndex);
                      if (
                        mergedCell &&
                        (rowIndex !== mergedCell.s.r ||
                          colIndex !== mergedCell.s.c)
                      ) {
                        return null;
                      }

                      const dimensions = getMergedCellDimensions(
                        rowIndex,
                        colIndex
                      );
                      const styleKey = `${rowIndex}-${colIndex}`;
                      const cellStyle = cellStyles[styleKey] || {};

                      return (
                        <td
                          key={colIndex}
                          rowSpan={dimensions?.rowSpan}
                          colSpan={dimensions?.colSpan}
                          className="border"
                          style={cellStyle}
                        >
                          {renderCellContent(cell, rowIndex, colIndex)}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

export default App;
