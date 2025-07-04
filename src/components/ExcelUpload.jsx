import React from "react";
import { Button, Snackbar } from "@mui/material";
import UploadFileIcon from "@mui/icons-material/UploadFile";
import * as XLSX from "xlsx";

const ExcelUpload = ({ onDataParsed }) => {
  const [snackbar, setSnackbar] = React.useState({ open: false, message: "" });

  const showMessage = (msg) => {
    setSnackbar({ open: true, message: msg });
  };

  const handleFile = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "buffer" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(sheet);
      const json = raw.map((item, idx) => ({
        ...item,
        __rowKey: `row-${idx}`,
      }));
      onDataParsed(json);
      showMessage("Excel file parsed successfully");
    } catch (err) {
      console.error(err);
      showMessage("Failed to parse Excel file");
    }
  };

  return (
    <div>
      <Button
        variant="contained"
        component="label"
        startIcon={<UploadFileIcon />}
      >
        Upload Excel File
        <input type="file" hidden onChange={handleFile} />
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

export default ExcelUpload;
