import React, { useState } from "react";
import apiClient from "./apiClient"; // <-- import the axios instance

const ExcelUpload = () => {
  const [files, setFiles] = useState([]);
  const [isDragging, setIsDragging] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [message, setMessage] = useState("");

  // Backend only allows .xlsx now
  const allowedExtensions = [".xlsx"];

  const filterExcelFiles = (fileList) => {
    const arr = Array.from(fileList);
    return arr.filter((file) =>
      allowedExtensions.some((ext) =>
        file.name.toLowerCase().endsWith(ext)
      )
    );
  };

  const handleFileChange = (e) => {
    const selected = filterExcelFiles(e.target.files || []);
    if (selected.length === 0 && (e.target.files || []).length > 0) {
      setMessage("Only .xlsx files are allowed.");
      return;
    }
    setFiles((prev) => [...prev, ...selected]);
    setMessage("");
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  };

  const handleDragLeave = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);

    const dropped = filterExcelFiles(e.dataTransfer.files || []);
    if (dropped.length === 0) {
      setMessage("Only .xlsx files are allowed.");
      return;
    }
    setFiles((prev) => [...prev, ...dropped]);
    setMessage("");
  };

  const handleRemoveFile = (index) => {
    setFiles((prev) => prev.filter((_, i) => i !== index));
  };

  const handleUpload = async () => {
    if (files.length === 0) {
      setMessage("Please select at least one Excel file.");
      return;
    }

    setUploading(true);
    setMessage("");

    try {
      const formData = new FormData();
      files.forEach((file) => formData.append("files", file));

      // Backend now returns JSON only (no compare Excel path needed)
      const res = await apiClient.post("format-statement/", formData, {
        headers: {
          "Content-Type": "multipart/form-data",
        },
      });

      const data = res.data || {};
      // Backend already sends a good summary string
      let msg = data.message || "Upload & processing successful.";

      setMessage(msg);
      setFiles([]);

      // Optional: if you want to debug counts on frontend:
      console.log("Row count summary:", data.row_count_summary);
      console.log("Mismatch summary:", data.mismatch_summary);
    } catch (err) {
      console.error(err);
      const serverData = err.response?.data;

      let errorMsg = "An unknown error occurred.";

      // If backend sent JSON
      if (serverData && typeof serverData === "object") {
        errorMsg = serverData.error || serverData.message || errorMsg;
      } else if (typeof serverData === "string") {
        // If backend sent HTML (Django debug page)
        if (
          serverData.trim().startsWith("<!DOCTYPE html>") ||
          serverData.includes("<html")
        ) {
          errorMsg = "Server error (500). Please check backend logs.";
        } else {
          errorMsg = serverData;
        }
      } else if (err.message) {
        errorMsg = err.message;
      }

      setMessage(`Error: ${errorMsg}`);
    } finally {
      setUploading(false);
    }
  };

  return (
    <div style={styles.container}>
      <h2 style={styles.title}>Upload Excel Statements</h2>

      {/* Drag & Drop Area */}
      <div
        style={{
          ...styles.dropZone,
          borderColor: isDragging ? "#2563eb" : "#cccccc",
          backgroundColor: isDragging ? "#eff6ff" : "#fafafa",
        }}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
      >
        <p style={styles.dropText}>
          Drag & drop Excel files here
          <br />
          <span style={styles.dropHint}>(.xlsx only)</span>
        </p>
        <p style={styles.orText}>OR</p>

        {/* File Input */}
        <label style={styles.fileLabel}>
          Choose files
          <input
            type="file"
            multiple
            accept=".xlsx"
            style={{ display: "none" }}
            onChange={handleFileChange}
          />
        </label>
      </div>

      {/* Selected files list */}
      {files.length > 0 && (
        <div style={styles.filesContainer}>
          <h4 style={styles.filesTitle}>Selected files:</h4>
          <ul style={styles.filesList}>
            {files.map((file, index) => (
              <li key={index} style={styles.fileItem}>
                <span>{file.name}</span>
                <button
                  type="button"
                  style={styles.removeButton}
                  onClick={() => handleRemoveFile(index)}
                >
                  ✕
                </button>
              </li>
            ))}
          </ul>
        </div>
      )}

      <div style={styles.ButtonContainer}>
        {/* Upload button */}
        <button
          type="button"
          style={{
            ...styles.uploadButton,
            opacity: uploading ? 0.7 : 1,
            cursor: uploading ? "not-allowed" : "pointer",
          }}
          disabled={uploading}
          onClick={handleUpload}
        >
          {uploading ? "Processing..." : "Process"}
        </button>
      </div>

      {/* Status message */}
      {message && <p style={styles.message}>{message}</p>}
    </div>
  );
};

const styles = {
  container: {
    width: "700px",
    margin: "40px auto",
    padding: "24px",
    borderRadius: "12px",
    border: "1px solid #e5e7eb",
    backgroundColor: "#ffffff",
    fontFamily:
      "-apple-system, BlinkMacSystemFont, 'Segoe UI', system-ui, sans-serif",
    boxShadow: "0 10px 25px rgba(15, 23, 42, 0.05)",
  },
  title: {
    marginBottom: "16px",
    fontSize: "20px",
    fontWeight: 600,
    color: "#111827",
  },
  dropZone: {
    padding: "24px",
    borderRadius: "10px",
    border: "2px dashed #cccccc",
    textAlign: "center",
    transition: "all 0.15s ease-in-out",
  },
  dropText: {
    margin: 0,
    fontSize: "14px",
    color: "#374151",
  },
  dropHint: {
    fontSize: "12px",
    color: "#6b7280",
  },
  orText: {
    fontSize: "12px",
    color: "#9ca3af",
    margin: "10px 0",
  },
  fileLabel: {
    display: "inline-block",
    padding: "8px 16px",
    borderRadius: "999px",
    backgroundColor: "#2563eb",
    color: "#ffffff",
    fontSize: "14px",
    cursor: "pointer",
  },
  filesContainer: {
    marginTop: "16px",
  },
  filesTitle: {
    margin: "0 0 8px",
    fontSize: "14px",
    fontWeight: 500,
    color: "#374151",
  },
  filesList: {
    listStyle: "none",
    padding: 0,
    margin: 0,
  },
  fileItem: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "6px 10px",
    marginBottom: "4px",
    borderRadius: "6px",
    backgroundColor: "#e2e2e2",
    fontSize: "13px",
    color: "#555555",
  },
  removeButton: {
    border: "none",
    background: "transparent",
    color: "#ef4444",
    cursor: "pointer",
    fontSize: "14px",
  },
  ButtonContainer: {
    textAlign: "right",
  },
  uploadButton: {
    marginTop: "18px",
    padding: "10px 18px",
    borderRadius: "999px",
    border: "none",
    backgroundColor: "#16a34a",
    color: "#ffffff",
    fontSize: "14px",
    fontWeight: 500,
    textAlign: "right",
  },
  message: {
    marginTop: "12px",
    fontSize: "13px",
    color: "#374151",
  },
};

export default ExcelUpload;
