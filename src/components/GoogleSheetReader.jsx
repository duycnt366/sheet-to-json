import React, { useState } from "react";
import { TextField, Button, Snackbar } from "@mui/material";
import * as XLSX from "xlsx";
import axios from "axios";

const extractFileId = (url) => {
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
};

const GoogleSheetReader = ({ onDataParsed }) => {
  const [sheetUrl, setSheetUrl] = useState("");
  const [snackbar, setSnackbar] = useState({ open: false, message: "" });

  const showMessage = (msg) => {
    setSnackbar({ open: true, message: msg });
  };

  const fetchGoogleSheet = async () => {
    const fileId = extractFileId(sheetUrl);
    if (!fileId) return showMessage("Invalid Google Sheet URL");

    const exportUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;

    try {
      const res = await axios.get(exportUrl, { responseType: "arraybuffer" });
      const workbook = XLSX.read(res.data, { type: "buffer" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(sheet);
      const json = raw.map((item, idx) => ({
        ...item,
        __rowKey: `row-${idx}`,
      }));
      onDataParsed(json);
      showMessage("Google Sheet loaded successfully");
    } catch (err) {
      console.error(err);
      showMessage("Failed to fetch or parse sheet");
    }
  };

  return (
    <div
      style={{ display: "flex", gap: 8, alignItems: "center", marginTop: 16 }}
    >
      <TextField
        label="Google Sheet URL"
        variant="outlined"
        value={sheetUrl}
        onChange={(e) => setSheetUrl(e.target.value)}
        fullWidth
      />
      <Button variant="contained" onClick={fetchGoogleSheet}>
        Load Sheet
      </Button>

      <Snackbar
        open={snackbar.open}
        autoHideDuration={3000}
        onClose={() => setSnackbar({ open: false, message: "" })}
        message={snackbar.message}
      />
    </div>
  );
};

export default GoogleSheetReader;
