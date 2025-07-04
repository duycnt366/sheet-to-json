"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import { saveAs } from "file-saver";
import {
  Box,
  Button,
  Typography,
  List,
  ListItem,
  ListItemText,
} from "@mui/material";

export default function ExcelToZipDownloader() {
  const [jsonFiles, setJsonFiles] = useState([]);
  const [isProcessing, setIsProcessing] = useState(false);

  const handleSelectFolder = async (e) => {
    const files = e.target.files;
    if (!files) return;

    setIsProcessing(true);
    const fileArray = Array.from(files);
    const resultList = [];
    const zip = new JSZip();

    for (const file of fileArray) {
      if (!file.name.endsWith(".xlsx")) continue;

      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array" });

      for (const sheetName of workbook.SheetNames) {
        const sheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });

        const baseName = file.name.replace(/\.xlsx$/, "");
        const fileName = `${baseName}__${sheetName}.json`;

        resultList.push({ filename: fileName, content: json });

        zip.file(fileName, JSON.stringify(json, null, 2));
      }
    }

    // Tạo và tải file zip
    const zipBlob = await zip.generateAsync({ type: "blob" });
    saveAs(zipBlob, "excel_json_export.zip");

    setJsonFiles(resultList);
    setIsProcessing(false);
  };

  return (
    <Box sx={{ maxWidth: 800, margin: "auto", p: 3 }}>
      <Typography variant="h6" gutterBottom>
        Convert Excel Files to JSON and Download as ZIP
      </Typography>

      <Button variant="contained" component="label" disabled={isProcessing}>
        {isProcessing ? "Processing..." : "Choose Folder"}
        <input
          type="file"
          webkitdirectory="true"
          multiple
          accept=".xlsx"
          hidden
          onChange={handleSelectFolder}
        />
      </Button>

      {jsonFiles.length > 0 && (
        <Box mt={4}>
          <Typography variant="subtitle1">
            Exported JSON Files ({jsonFiles.length}):
          </Typography>
          <List dense>
            {jsonFiles.map((file, idx) => (
              <ListItem key={idx}>
                <ListItemText
                  primary={file.filename}
                  secondary={`${file.content.length} rows`}
                />
              </ListItem>
            ))}
          </List>
        </Box>
      )}
    </Box>
  );
}
