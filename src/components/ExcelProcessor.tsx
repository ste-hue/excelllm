import React, { useState, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  FileSpreadsheet,
  Download,
  RefreshCw,
  AlertCircle,
  CheckCircle,
} from "lucide-react";

interface ProcessedData {
  json: string;
  csv: string;
  markdown: string;
  text: string;
  explanation: string;
}

export const ExcelProcessor: React.FC = () => {
  const [file, setFile] = useState<File | null>(null);
  const [sheets, setSheets] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>("");
  const [range, setRange] = useState<string>("A1:Z100");
  const [processedData, setProcessedData] = useState<ProcessedData | null>(
    null,
  );
  const [outputFormat, setOutputFormat] = useState<
    "json" | "csv" | "markdown" | "text"
  >("json");
  const [isProcessing, setIsProcessing] = useState(false);
  const [dragActive, setDragActive] = useState(false);
  const [message, setMessage] = useState<{
    type: "success" | "error" | null;
    text: string;
  }>({ type: null, text: "" });
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [activeTab, setActiveTab] = useState<"preview" | "analysis">("preview");

  const showMessage = (type: "success" | "error", text: string) => {
    setMessage({ type, text });
    setTimeout(() => setMessage({ type: null, text: "" }), 3000);
  };

  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFile(e.dataTransfer.files[0]);
    }
  };

  const handleFile = useCallback(async (uploadedFile: File) => {
    // Validate file
    if (uploadedFile.size > 10 * 1024 * 1024) {
      showMessage("error", "File size must be less than 10MB");
      return;
    }

    const validTypes = [
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "application/vnd.ms-excel",
    ];

    if (
      !validTypes.includes(uploadedFile.type) &&
      !uploadedFile.name.match(/\.(xlsx|xls)$/)
    ) {
      showMessage("error", "Please upload a valid Excel file (.xlsx or .xls)");
      return;
    }

    setIsProcessing(true);
    setFile(uploadedFile);

    try {
      const data = await uploadedFile.arrayBuffer();
      const wb = XLSX.read(new Uint8Array(data), {
        type: "array",
        cellFormula: true,
      });
      setWorkbook(wb);
      setSheets(wb.SheetNames);
      setSelectedSheet(wb.SheetNames[0] || "");
      showMessage(
        "success",
        `File uploaded! Found ${wb.SheetNames.length} sheet(s)`,
      );
    } catch (error) {
      showMessage("error", "Error reading file. Please check the file format.");
      setFile(null);
    } finally {
      setIsProcessing(false);
    }
  }, []);

  const processSheet = useCallback(async () => {
    if (!workbook || !selectedSheet) return;

    setIsProcessing(true);
    try {
      const worksheet = workbook.Sheets[selectedSheet];
      const rangeObj = XLSX.utils.decode_range(range);

      // Extract data as 2D array
      const jsonData: any[][] = XLSX.utils.sheet_to_json(worksheet, {
        range: rangeObj,
        header: 1,
        raw: false,
        defval: "",
      }) as any[][];

      // Extract formulas
      const formulas: { [key: string]: string } = {};
      for (let R = rangeObj.s.r; R <= rangeObj.e.r; ++R) {
        for (let C = rangeObj.s.c; C <= rangeObj.e.c; ++C) {
          const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
          const cell = worksheet[cellAddress];
          if (cell && cell.f) {
            formulas[cellAddress] = cell.f;
          }
        }
      }

      // Generate different formats
      const csv = XLSX.utils.sheet_to_csv(worksheet, { FS: ",", RS: "\n" });

      const markdown =
        jsonData.length > 0
          ? (() => {
              const headers = jsonData[0];
              const rows = jsonData.slice(1);
              let md = `| ${headers.join(" | ")} |\n`;
              md += `| ${headers.map(() => "---").join(" | ")} |\n`;
              rows.forEach((row) => {
                md += `| ${row.join(" | ")} |\n`;
              });
              return md;
            })()
          : "";

      const text = jsonData.map((row) => row.join("\t")).join("\n");

      const explanation = [
        `# Excel Sheet Analysis: ${selectedSheet}`,
        `\n## Data Range: ${range}`,
        `\n## Structure:`,
        `- Total rows: ${jsonData.length}`,
        `- Total columns: ${jsonData[0]?.length || 0}`,
        jsonData.length > 0 ? `\n## Headers: ${jsonData[0].join(", ")}` : "",
        Object.keys(formulas).length > 0 ? "\n## Formulas Found:" : "",
        ...Object.entries(formulas).map(
          ([cell, formula]) => `- Cell ${cell}: ${formula}`,
        ),
        "\n## LLM Processing Notes:",
        "This Excel data has been converted for AI processing. The structure includes:",
        "1. Headers in the first row for context",
        "2. All formulas extracted and listed separately",
        "3. Data normalized to text format for better LLM understanding",
        "4. Range-specific extraction to focus on relevant data",
      ].join("\n");

      setProcessedData({
        json: JSON.stringify(jsonData, null, 2),
        csv,
        markdown,
        text,
        explanation,
      });

      showMessage(
        "success",
        "Processing complete! Data is ready for AI consumption.",
      );
    } catch (error) {
      showMessage(
        "error",
        "Error processing sheet. Please check your range and try again.",
      );
    } finally {
      setIsProcessing(false);
    }
  }, [workbook, selectedSheet, range]);

  const downloadData = () => {
    if (!processedData) return;

    const formatMap = {
      json: { data: processedData.json, ext: "json", mime: "application/json" },
      csv: { data: processedData.csv, ext: "csv", mime: "text/csv" },
      markdown: {
        data: processedData.markdown,
        ext: "md",
        mime: "text/markdown",
      },
      text: { data: processedData.text, ext: "txt", mime: "text/plain" },
    };

    const format = formatMap[outputFormat];
    const blob = new Blob([format.data], { type: format.mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `excel-export.${format.ext}`;
    a.click();
    URL.revokeObjectURL(url);

    showMessage("success", "Download started!");
  };

  const reset = () => {
    setFile(null);
    setSheets([]);
    setSelectedSheet("");
    setWorkbook(null);
    setProcessedData(null);
    setRange("A1:Z100");
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 p-6">
      <div className="max-w-6xl mx-auto">
        <div className="text-center mb-8">
          <h1 className="text-4xl font-bold text-gray-800 mb-2">
            Excel to LLM Converter
          </h1>
          <p className="text-gray-600">
            Transform Excel files into AI-friendly formats with formula
            extraction
          </p>
        </div>

        {/* Message Toast */}
        {message.type && (
          <div
            className={`fixed top-4 right-4 px-4 py-2 rounded-lg shadow-lg flex items-center gap-2 ${
              message.type === "success"
                ? "bg-green-500 text-white"
                : "bg-red-500 text-white"
            }`}
          >
            {message.type === "success" ? (
              <CheckCircle size={20} />
            ) : (
              <AlertCircle size={20} />
            )}
            {message.text}
          </div>
        )}

        <div className="grid md:grid-cols-2 gap-6">
          {/* Upload Section */}
          <div className="bg-white rounded-lg shadow-lg p-6">
            <div className="flex items-center justify-between mb-4">
              <h2 className="text-xl font-semibold text-gray-800">
                Upload Excel File
              </h2>
              {file && (
                <button
                  onClick={reset}
                  className="flex items-center gap-2 px-3 py-1 text-sm bg-gray-100 hover:bg-gray-200 rounded-md transition-colors"
                >
                  <RefreshCw size={16} />
                  Reset
                </button>
              )}
            </div>

            <div
              onDragEnter={handleDrag}
              onDragLeave={handleDrag}
              onDragOver={handleDrag}
              onDrop={handleDrop}
              className={`border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-all ${
                dragActive
                  ? "border-blue-500 bg-blue-50"
                  : "border-gray-300 hover:border-gray-400"
              } ${isProcessing ? "opacity-50 cursor-not-allowed" : ""}`}
            >
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) =>
                  e.target.files && handleFile(e.target.files[0])
                }
                className="hidden"
                id="file-upload"
                disabled={isProcessing}
              />
              <label htmlFor="file-upload" className="cursor-pointer">
                <FileSpreadsheet
                  className={`mx-auto h-12 w-12 mb-4 ${
                    dragActive ? "text-blue-500" : "text-gray-400"
                  }`}
                />
                {isProcessing ? (
                  <p className="text-gray-600">Processing...</p>
                ) : (
                  <>
                    <p className="font-medium text-gray-700">
                      {dragActive
                        ? "Drop file here"
                        : "Drag & drop Excel file here"}
                    </p>
                    <p className="text-sm text-gray-500 mt-2">
                      or click to browse • .xlsx/.xls • Max 10MB
                    </p>
                  </>
                )}
              </label>
            </div>

            {file && sheets.length > 0 && (
              <div className="mt-6 space-y-4">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Select Sheet
                  </label>
                  <select
                    value={selectedSheet}
                    onChange={(e) => setSelectedSheet(e.target.value)}
                    className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                  >
                    {sheets.map((sheet) => (
                      <option key={sheet} value={sheet}>
                        {sheet}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Cell Range
                  </label>
                  <input
                    type="text"
                    value={range}
                    onChange={(e) => setRange(e.target.value)}
                    placeholder="A1:Z100"
                    className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                  />
                  <p className="text-xs text-gray-500 mt-1">
                    Example: A1:D10 for columns A-D, rows 1-10
                  </p>
                </div>

                <button
                  onClick={processSheet}
                  disabled={isProcessing || !selectedSheet}
                  className="w-full bg-blue-500 text-white py-2 px-4 rounded-md hover:bg-blue-600 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors"
                >
                  {isProcessing ? "Processing..." : "Process Sheet"}
                </button>
              </div>
            )}
          </div>

          {/* Output Section */}
          <div className="bg-white rounded-lg shadow-lg p-6">
            <h2 className="text-xl font-semibold text-gray-800 mb-4">
              Output & Export
            </h2>

            {processedData ? (
              <div className="space-y-4">
                <div className="flex items-center gap-2">
                  <label className="text-sm font-medium text-gray-700">
                    Export Format:
                  </label>
                  <select
                    value={outputFormat}
                    onChange={(e) => setOutputFormat(e.target.value as any)}
                    className="px-3 py-1 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                  >
                    <option value="json">JSON</option>
                    <option value="csv">CSV</option>
                    <option value="markdown">Markdown</option>
                    <option value="text">Text</option>
                  </select>
                  <button
                    onClick={downloadData}
                    className="ml-auto flex items-center gap-2 bg-green-500 text-white px-4 py-1 rounded-md hover:bg-green-600 transition-colors"
                  >
                    <Download size={16} />
                    Download
                  </button>
                </div>

                <div className="border-t pt-4">
                  <div className="flex gap-2 mb-2">
                    <button
                      onClick={() => setActiveTab("preview")}
                      className={`px-3 py-1 text-sm font-medium border-b-2 transition-colors ${
                        activeTab === "preview"
                          ? "border-blue-500 text-blue-600"
                          : "border-transparent text-gray-500 hover:text-gray-700"
                      }`}
                    >
                      Preview
                    </button>
                    <button
                      onClick={() => setActiveTab("analysis")}
                      className={`px-3 py-1 text-sm font-medium border-b-2 transition-colors ${
                        activeTab === "analysis"
                          ? "border-blue-500 text-blue-600"
                          : "border-transparent text-gray-500 hover:text-gray-700"
                      }`}
                    >
                      LLM Analysis
                    </button>
                  </div>

                  <div className="bg-gray-50 rounded-md p-4 max-h-96 overflow-auto">
                    <pre className="text-sm whitespace-pre-wrap font-mono">
                      {activeTab === "preview"
                        ? processedData[outputFormat]
                        : processedData.explanation}
                    </pre>
                  </div>
                </div>
              </div>
            ) : (
              <div className="text-center py-12 text-gray-400">
                <FileSpreadsheet className="mx-auto h-12 w-12 mb-4" />
                <p>Upload and process an Excel file to see the output</p>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};
