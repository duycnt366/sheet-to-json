"use client";

import React, { useEffect, useState } from "react";
import {
  Typography,
  Button,
  Snackbar,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Paper,
  Select,
  MenuItem,
  InputLabel,
  FormControl,
  Stack,
  CircularProgress,
  Slider,
} from "@mui/material";
import UploadFileIcon from "@mui/icons-material/UploadFile";
import * as XLSX from "xlsx";
import axios from "axios";

const getSheetColumnHeaders = (sheet) => {
  const ref = sheet["!ref"];
  if (!ref) return [];

  const range = XLSX.utils.decode_range(ref);
  const headers = [];

  for (let col = range.s.c; col <= range.e.c; col++) {
    const colLetter = XLSX.utils.encode_col(col); // "A", "B", etc.
    headers.push(colLetter);
  }

  return headers;
};

export default function Home() {
  const [loading, setLoading] = useState(false);

  const [data, setData] = useState([]);
  const [range, setRange] = useState([0, 0]);
  const [headers, setHeaders] = useState([]);

  const [sheetUrl, setSheetUrl] = useState("");

  const [sheetName, setSheetName] = useState("");
  const [sheetNames, setSheetNames] = useState([]);
  const [workbook, setWorkbook] = useState(null);

  const [snackbar, setSnackbar] = useState({ open: false, message: "" });

  const showMessage = (msg) => {
    setSnackbar({ open: true, message: msg });
  };

  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setLoading(true);
    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: "buffer" });
      setWorkbook(wb);
      setSheetNames(wb.SheetNames);
      showMessage("Excel file parsed successfully");
    } catch (err) {
      console.error(err);
      showMessage("Error parsing file");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (sheetName && workbook) {
      const sheet = workbook.Sheets[sheetName];
      const columnHeaders = getSheetColumnHeaders(sheet);
      const json = XLSX.utils.sheet_to_json(sheet, { header: "A" });
      setData(json);
      setRange([1, json.length]);
      setHeaders(columnHeaders);
    } else {
      setData([]);
    }
  }, [sheetName, workbook]);

  const extractFileId = (url) => {
    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : null;
  };

  const handleGoogleSheet = async () => {
    const fileId = extractFileId(sheetUrl);
    if (!fileId) return showMessage("Invalid Google Sheet URL");

    const exportUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;

    try {
      const response = await axios.get(exportUrl, {
        responseType: "arraybuffer",
      });
      const wb = XLSX.read(response.data, { type: "buffer" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet);
      const withKeys = json.map((item, idx) => ({
        ...item,
        __key: `row-${idx}`,
      }));
      setData(withKeys);
      showMessage("Google Sheet loaded");
    } catch (err) {
      console.error(err);
      showMessage("Failed to fetch Google Sheet");
    }
  };

  const handleChange = (event, newValue) => {
    setRange(newValue);
  };

  const exportJsonFile = (data, range, filename = "data.json") => {
    const from = range[0] || 1;
    const end = range[1];
    const slice = data.slice(from - 1, end);
    const jsonString = JSON.stringify(slice, null, 2); // pretty print
    const blob = new Blob([jsonString], { type: "application/json" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.click();

    URL.revokeObjectURL(url);
  };

  return (
    <div style={{ padding: 24 }}>
      <Typography variant="h5" gutterBottom>
        Excel or Google Sheet Reader
      </Typography>
      <Stack direction="row" spacing={3}>
        {loading ? (
          <Button variant="outlined" disabled>
            <CircularProgress />
            Loading…
          </Button>
        ) : (
          <Button
            size="large"
            variant="contained"
            component="label"
            startIcon={<UploadFileIcon />}
            disabled={loading}
            loading={loading}
          >
            Upload Excel File
            <input type="file" hidden onChange={handleFileUpload} />
          </Button>
        )}

        {sheetNames.length > 0 && (
          <FormControl size="large" style={{ marginTop: 16, minWidth: 200 }}>
            <InputLabel id="sheet-select-label">Sheet Name</InputLabel>
            <Select
              labelId="sheet-select-label"
              value={sheetName}
              label="Sheet Name"
              onChange={(e) => setSheetName(e.target.value)}
            >
              {sheetNames.map((name) => (
                <MenuItem key={name} value={name}>
                  {name}
                </MenuItem>
              ))}
            </Select>
          </FormControl>
        )}

        {data.length ? (
          <div>
            <Typography gutterBottom>
              Price Range: {range[0]} - {range[1]}
            </Typography>
            <Slider
              value={range}
              onChange={handleChange}
              valueLabelDisplay="auto"
              min={0}
              max={100}
              step={1}
            />
          </div>
        ) : null}

        {data.length ? (
          <Button
            size="large"
            variant="outlined"
            onClick={() => {
              if (!data || data.length === 0) return;
              const from = range[0] || 1;
              const end = range[1];
              const slice = data.slice(from - 1, end);

              const jsonString = JSON.stringify(slice, null, 2);
              navigator.clipboard.writeText(jsonString).then(() => {
                setSnackbar({
                  open: true,
                  message: "JSON copied to clipboard",
                });
              });
            }}
          >
            Copy JSON
          </Button>
        ) : null}

        {data.length ? (
          <Button
            size="large"
            variant="outlined"
            color="primary"
            style={{ marginLeft: 16 }}
            onClick={() => exportJsonFile(data, range, "sheet-data.json")}
            disabled={!data || !data.length}
          >
            Export JSON
          </Button>
        ) : null}
      </Stack>

      {data.length ? (
        <TableContainer
          component={Paper}
          style={{ marginTop: 24, maxHeight: "75vh", overflow: "scroll" }}
        >
          <Table stickyHeader size="small">
            <TableHead>
              <TableRow>
                {headers.map((key) => (
                  <TableCell key={key}>{key}</TableCell>
                ))}
              </TableRow>
            </TableHead>
            <TableBody>
              {data
                .slice(
                  Math.max(range[0], 0),
                  Math.min(range[1], Math.max(range[0], 0) + 10)
                )
                .map((row, index) => (
                  <TableRow key={index}>
                    {headers.map((key) => (
                      <TableCell key={key}>{row[key]}</TableCell>
                    ))}
                  </TableRow>
                ))}
            </TableBody>
          </Table>
        </TableContainer>
      ) : null}

      <Snackbar
        open={snackbar.open}
        autoHideDuration={3000}
        onClose={() => setSnackbar({ open: false, message: "" })}
        message={snackbar.message}
      />
    </div>
  );
}
