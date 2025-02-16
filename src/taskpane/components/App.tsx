import * as React from "react";
import { makeStyles, Button, Tab, TabList, SelectTabEvent, SelectTabData } from "@fluentui/react-components";
import packageJson from "../../../package.json";

interface AppProps {
  title: string;
}

interface ImportData {
  resultCount: string;
  model: string;
  quantity: number;
  productCodeProvider: string;
  provider: string;
  manufacturer: string;
  minOrderQuantity: number;
  packaging: string;
  stockCount: number;
  price: number;
  lcscPartNumber: string;
  attributes: any[];
}

const useStyles = makeStyles({
  root: {
    padding: "10px",
    backgroundColor: "#fcfcfc",
    minHeight: "100vh",
  },
  button: {
    marginTop: "20px",
  },
  results: {
    marginTop: "20px",
    whiteSpace: "pre-wrap",
  },

  textarea: {
    minHeight: "100px",
    width: "100%",
    marginTop: "10px",
    boxSizing: "border-box",
    cursor: "text",
    userSelect: "text",
    WebkitUserSelect: "text",
  },
  checkboxContainer: {
    display: "flex",
    flexDirection: "column",
    gap: "10px",
    marginTop: "20px",
  },
  importTextarea: {
    width: "100%",
    height: "100px",
    marginBottom: "10px",
    boxSizing: "border-box",
    cursor: "text",
    userSelect: "text",
    WebkitUserSelect: "text",
  },
  debugOutput: {
    width: "100%",
    height: "200px",
    marginTop: "20px",
    backgroundColor: "#f0f0f0",
    fontFamily: "monospace",
    padding: "8px",
    overflow: "auto",
    border: "1px solid #ccc",
    boxSizing: "border-box",
  },
});

const App: React.FC<AppProps> = () => {
  const styles = useStyles();
  const [jsonResult, setJsonResult] = React.useState<string>("");
  const [selectedTab, setSelectedTab] = React.useState<string>("export");
  const [importData, setImportData] = React.useState<string>("");
  const [selectedFields, setSelectedFields] = React.useState<{
    [key: string]: boolean;
  }>({});
  const [availableFields, setAvailableFields] = React.useState<string[]>([]);
  const [debugLog, setDebugLog] = React.useState<string[]>([]);
  let headerRow = 0;
  let lastColumn = 0;

  React.useEffect(() => {
    // Clear formatting when component mounts
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const entireSheet = sheet.getUsedRange();
      entireSheet.format.fill.clear();
      entireSheet.format.font.color = "black";
      await context.sync();
    }).catch((error) => console.error("Error clearing formatting:", error));
  }, []);

  const copyToClipboard = () => {
    navigator.clipboard.writeText(jsonResult);
  };

  const handleTabSelect = (_event: SelectTabEvent, data: SelectTabData) => {
    setSelectedTab(data.value as string);
  };

  const handleImportDataChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setImportData(e.target.value);
    try {
      const parsed: ImportData[] = JSON.parse(e.target.value);
      if (parsed.length > 0) {
        const fields = Object.keys(parsed[0]).filter((key) => key !== "attributes");
        setAvailableFields(fields);
        const initialSelectedFields = fields.reduce(
          (acc, field) => ({
            ...acc,
            [field]: [
              "lcscComponentId",
              "overseasStockCount",
              "postStockCount",
              "privateStockCount",
              "idleStockCount",
              "status",
              "error",
              "lastOrdered",
              "startNumber",
              "jlcGoodsPrice",
              "gsGoodsPrice",
            ].includes(field),
          }),
          {}
        );
        setSelectedFields(initialSelectedFields);
      }
    } catch (error) {
      console.error("Invalid JSON:", error);
    }
  };

  const handleCheckboxChange = (field: string) => {
    setSelectedFields((prev) => ({
      ...prev,
      [field]: !prev[field],
    }));
  };

  const initializeHeaderRow = async () => {
    let foundHeaderRow = 0;
    let foundLastColumn = 0;

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange();
      usedRange.load(["values", "rowCount", "columnCount"]);
      await context.sync();
      addDebugLog(`Row count: ${usedRange.rowCount}`);
      addDebugLog(`Column count: ${usedRange.columnCount}`);

      if (!usedRange || !usedRange.values) {
        addDebugLog("ERROR: Worksheet is empty");
        return;
      }

      // Find the header row by looking for 'designators' in the first column only
      for (let i = 0; i < usedRange.values.length; i++) {
        const row = usedRange.values[i];
        if (row && row[0] && String(row[0]).toLowerCase() === "designators") {
          foundHeaderRow = i;
          foundLastColumn = row.length;
          break;
        }
      }
      addDebugLog(`Found header row at index: ${foundHeaderRow}`);
      addDebugLog(`Found last column: ${foundLastColumn}`);
    });

    // Always return these values, even if they're still 0
    return { headerRow: foundHeaderRow, lastColumn: foundLastColumn };
  };

  const addDebugLog = (message: string) => {
    setDebugLog((prev) => {
      const newLogs = [
        ...prev,
        `${new Date().toLocaleTimeString("de-DE", { timeZone: "Europe/Berlin" })} - ${message}`,
      ];
      // Auto-scroll after state update
      setTimeout(() => {
        const debugElement = document.querySelector(`.${styles.debugOutput}`);
        if (debugElement) {
          debugElement.scrollTop = debugElement.scrollHeight;
        }
      }, 0);
      return newLogs;
    });
  };

  const applyImportData = async () => {
    try {
      const { headerRow: foundHeaderRow, lastColumn: foundLastColumn } = await initializeHeaderRow();
      headerRow = foundHeaderRow;
      lastColumn = foundLastColumn;

      addDebugLog(`Working with headerRow: ${headerRow}, lastColumn: ${lastColumn}`);

      if (headerRow === 0) {
        addDebugLog("ERROR: Could not find header row with LCSC column!");
        return;
      }

      addDebugLog("Starting import process...");
      const data: ImportData[] = JSON.parse(importData);
      addDebugLog(`Parsed ${data.length} items from JSON`);

      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // Clear ALL background colors in the sheet first
        const entireSheet = sheet.getUsedRange();
        entireSheet.format.fill.clear();
        await context.sync();
        addDebugLog("Cleared all background colors");

        const dataRange = sheet.getRangeByIndexes(headerRow, 0, 100, lastColumn);
        dataRange.load("values");
        await context.sync();

        // Get original headers to identify read-only columns
        const originalHeaders = dataRange.values[0];

        // Mark read-only columns in light grey
        originalHeaders.forEach((header, columnIndex) => {
          if (String(header).startsWith("_")) {
            const columnRange = sheet.getRangeByIndexes(headerRow + 1, columnIndex, 100, 1);
            columnRange.format.fill.color = "#F5F5F5"; // Light grey for read-only columns
          }
        });

        // Color the header row
        const headerRange = sheet.getRangeByIndexes(headerRow, 0, 1, lastColumn);
        await context.sync();
        headerRange.format.fill.color = "#4472C4";
        headerRange.format.font.color = "white";
        addDebugLog("Applied header formatting");

        await context.sync();

        entireSheet.load("values");
        await context.sync();
        addDebugLog("Loaded entire sheet values");

        const headers = dataRange.values[0].map((h) =>
          h === null || h === undefined ? "" : String(h).toLowerCase().trim()
        );
        addDebugLog(`Found headers: ${headers.join(", ")}`);
        const lcscIndex = headers.findIndex((h) => h === "lcscpartNumber".toLowerCase());
        addDebugLog(`LCSC column index: ${lcscIndex}`);

        const existingHeaders = headers;
        const newColumns = Object.keys(selectedFields)
          .filter((field) => selectedFields[field])
          .filter((field) => !existingHeaders.includes(field.toLowerCase()));
        addDebugLog(`New columns to add: ${newColumns.join(", ")}`);

        if (newColumns.length > 0) {
          const newHeadersRange = sheet.getRangeByIndexes(headerRow, lastColumn, 1, newColumns.length);
          newHeadersRange.values = [newColumns];
          await context.sync();
        }

        let updatedRows = 0;
        for (const item of data) {
          addDebugLog(`Processing item with LCSC code: ${item.lcscPartNumber}`);

          const excelRows = dataRange.values.slice(1);
          const rowIndex = excelRows.findIndex((row) => {
            const cellValue = row[lcscIndex];
            return cellValue && String(cellValue).trim() === item.lcscPartNumber;
          });

          if (rowIndex !== -1) {
            const actualRowIndex = headerRow + rowIndex + 1;
            updatedRows++;
            addDebugLog(`Found match at row ${actualRowIndex}`);

            // Color the matched row green
            const updatedRowRange = sheet.getRangeByIndexes(actualRowIndex, 0, 1, lastColumn);
            updatedRowRange.format.fill.color = "#E2EFDA";

            // Get original header values (not lowercase) for underscore check
            const originalHeaders = dataRange.values[0];

            Object.keys(selectedFields).forEach((field) => {
              if (selectedFields[field]) {
                const headerIndex = headers.findIndex((h) => h === field.toLowerCase());
                if (headerIndex !== -1) {
                  // Check if the header starts with underscore
                  const originalHeader = originalHeaders[headerIndex];
                  if (!String(originalHeader).startsWith("_")) {
                    const range = sheet.getRangeByIndexes(actualRowIndex, headerIndex, 1, 1);
                    range.values = [[item[field as keyof ImportData]]];
                  } else {
                    addDebugLog(`Skipping column ${field} as it is read-only.`);
                  }
                }
              }
            });

            if (newColumns.length > 0) {
              const newDataRange = sheet.getRangeByIndexes(actualRowIndex, lastColumn, 1, newColumns.length);
              const newValues = newColumns.map((col) => item[col as keyof ImportData]);
              newDataRange.values = [newValues];
            }
          } else {
            addDebugLog(`No match found for LCSC code: ${item.lcscPartNumber}`);
          }
        }

        addDebugLog(`Updated ${updatedRows} rows in total`);
        await context.sync();
        addDebugLog("Import completed successfully");
      });
    } catch (error) {
      addDebugLog(`ERROR: ${error.message}`);
      console.error("Error applying import data:", error);
    }
  };

  const extractComponents = async () => {
    try {
      const { headerRow: foundHeaderRow, lastColumn: foundLastColumn } = await initializeHeaderRow();
      headerRow = foundHeaderRow;
      lastColumn = foundLastColumn;

      addDebugLog(`Working with headerRow: ${headerRow}, lastColumn: ${lastColumn}`);
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // Clear ALL background colors and font formatting in the sheet first
        const entireSheet = sheet.getUsedRange();
        entireSheet.format.fill.clear();
        entireSheet.format.font.color = "black";
        await context.sync();

        // Load the range to get the row count
        entireSheet.load("rowCount");
        await context.sync();

        // Color the header row
        const headerRange = sheet.getRangeByIndexes(headerRow, 0, 1, lastColumn);
        headerRange.format.fill.color = "#4472C4"; // Darker blue for header
        headerRange.format.font.color = "white";
        await context.sync();

        // Color the data rows that will be exported (from header to last used row)
        const dataRange = sheet.getRangeByIndexes(headerRow + 1, 0, entireSheet.rowCount - headerRow - 1, lastColumn);
        dataRange.format.fill.color = "#E6F3FF"; // Light blue for data
        await context.sync();

        entireSheet.load("values");
        await context.sync();

        if (!entireSheet.values || entireSheet.values.length < headerRow + 1) {
          setJsonResult("Not enough data in worksheet");
          return;
        }

        const headers = entireSheet.values[headerRow].map((h: any) =>
          h === null || h === undefined ? "" : String(h).toLowerCase().trim()
        );

        addDebugLog(`Found headers: ${headers.join(", ")}`);
        const lcscColumnIndex = headers.findIndex((h) =>
          ["lcsc", "lcsc #", "lcscpartnumber", "lcsc part number"].includes(h)
        );
        addDebugLog(`LCSC column index: ${lcscColumnIndex}`);

        const data = entireSheet.values
          .slice(headerRow + 1)
          .filter((row) => row.some((cell) => cell !== "" && cell !== null && cell !== undefined));

        addDebugLog(`Total rows after initial filter: ${data.length}`);

        const components = data
          .filter((row) => {
            const designator = getCellValue(row, headers, ["designators", "designator"]);
            const lcscPart = row[lcscColumnIndex] ? String(row[lcscColumnIndex]).trim() : "";

            addDebugLog(
              `Row data - Designator: ${designator}, LCSC (raw): ${row[lcscColumnIndex]}, LCSC (processed): ${lcscPart}`
            );

            const isValid = designator && designator.trim() !== "" && lcscPart !== "";

            if (!isValid) {
              addDebugLog(
                `Skipping row - Invalid conditions: ${
                  !designator
                    ? "No designator"
                    : designator.trim() === ""
                      ? "Empty designator"
                      : lcscPart === ""
                        ? "Empty LCSC part"
                        : "Unknown reason"
                }`
              );
            }

            return isValid;
          })
          .map((row) => ({
            designator: getCellValue(row, headers, ["designators", "designator"]),
            description: getCellValue(row, headers, ["desc", "description"]),
            qtyPerBoard: parseFloat(getCellValue(row, headers, ["qtyPerBoard", "_qtyPerBoard", "qty"])) || 0,
            qtyPerOrder: parseFloat(getCellValue(row, headers, ["qtyPerOrder", "_qtyPerOrder"])) || 0,
            qtyToConsign: parseFloat(getCellValue(row, headers, ["qtyToConsign", "_qtyToConsign"])) || 0,
            source: getCellValue(row, headers, ["source"]),
            provider: getCellValue(row, headers, ["provider"]),
            ordered: getCellValue(row, headers, ["ordered"]),
            minimumQty: parseFloat(getCellValue(row, headers, ["minimumQty", "_minimumQty", "min qty"])) || 0,
            bufferQty: parseFloat(getCellValue(row, headers, ["bufferQty", "_bufferQty"])) || 0,
            model: getCellValue(row, headers, ["model"]) || "",
            manufacturerPartNumber: getCellValue(row, headers, [
              "mfr #",
              "manufacturer part number",
              "model",
              "manufacturerPartNumber",
            ]),
            productCodeProvider: getCellValue(row, headers, ["productcodeprovider"]),
            lcscPartNumber: row[lcscColumnIndex] ? String(row[lcscColumnIndex]).trim() : "",
            jlcpcbPartNumber: getCellValue(row, headers, ["jlcpcb #", "jlcpcb", "jlcpcbPartNumber"]),
            mouserPartNumber: getCellValue(row, headers, ["mouser #", "mouser"]),
            digikeyPartNumber: getCellValue(row, headers, ["digikey #", "digikey"]),
            comment: getCellValue(row, headers, ["comment", "comments"]),
            packaging: getCellValue(row, headers, ["packaging"]),
          }));

        const jsonOutput = JSON.stringify(components, null, 2);
        setJsonResult(jsonOutput);
        addDebugLog(`Successfully extracted ${components.length} components`);
      });
    } catch (error) {
      console.error("Detailed error:", error);
      addDebugLog(`Error: ${error.message}`);
      setJsonResult(`Error processing components data: ${error.message}`);
    }
  };

  const getCellValue = (row: any[], headers: string[], possibleNames: string[]): string => {
    for (const name of possibleNames) {
      const index = headers.indexOf(name);
      if (index !== -1 && row[index]) {
        return row[index].toString().trim();
      }
    }
    return "";
  };

  return (
    <div className={styles.root}>
      <h1>BOM Assistant </h1>
      <p>
        Version {packageJson.version} (updated at {packageJson.date})
      </p>
      <TabList selectedValue={selectedTab} onTabSelect={handleTabSelect}>
        <Tab value="export">Export BOM</Tab>
        <Tab value="import">Update BOM</Tab>
      </TabList>

      {selectedTab === "export" ? (
        <>
          <Button appearance="primary" onClick={extractComponents} className={styles.button}>
            Create BOM
          </Button>
          {jsonResult && (
            <div>
              <textarea
                value={jsonResult}
                readOnly
                className={styles.textarea}
                onFocus={(e) => setTimeout(() => e.target.select(), 0)}
                aria-label="BOM Export Data"
              />
              <Button onClick={copyToClipboard} className={styles.button}>
                Copy to Clipboard
              </Button>
            </div>
          )}
        </>
      ) : (
        <div>
          <textarea
            value={importData}
            onChange={handleImportDataChange}
            className={styles.importTextarea}
            placeholder="Paste JSON data here..."
            onFocus={(e) => setTimeout(() => e.target.select(), 0)}
            aria-label="BOM Import Data"
          />
          <div className={styles.checkboxContainer}>
            {availableFields.map((field) => (
              <label key={field}>
                <input
                  type="checkbox"
                  checked={selectedFields[field] || false}
                  onChange={() => handleCheckboxChange(field)}
                />
                {field}
              </label>
            ))}
          </div>
          <Button
            appearance="primary"
            onClick={applyImportData}
            className={styles.button}
            disabled={!importData || Object.values(selectedFields).every((v) => !v)}
          >
            Apply Import Data
          </Button>
        </div>
      )}
      <div className={styles.debugOutput}>
        {debugLog.map((log, index) => (
          <div key={index}>{log}</div>
        ))}
      </div>
    </div>
  );
};

export default App;
