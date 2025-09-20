import React, { useState } from "react";
import {
  View,
  Text,
  TouchableOpacity,
  StyleSheet,
  FlatList,
  TextInput,
  Alert,
} from "react-native";
import { pick, types } from "@react-native-documents/picker";
import RNFS from "react-native-fs";
import XLSX from "xlsx";
import FileViewer from "react-native-file-viewer";
export default function App() {
  const [excelData, setExcelData] = useState([]); // 2D array (rows & columns)

  //  Pick Excel File & Parse
  const browseAndLoadExcel = async () => {
    try {
      const [res] = await pick({
        type: [types.xls, types.xlsx],
      });

      console.log("Picked file:", res);

      // Read file as base64
      const bstr = await RNFS.readFile(res.uri, "base64");
      const wb = XLSX.read(bstr, { type: "base64" });

      // Get first sheet
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];

      // Convert to JSON (array of arrays)
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
      setExcelData(data);
    } catch (err) {
      if (err?.message?.includes("User cancelled")) {
        console.log("User cancelled picker");
      } else {
        console.error("Error loading Excel:", err);
      }
    }
  };

  // Handle Editing
  const handleEdit = (rowIndex, colIndex, value) => {
    const updated = [...excelData];
    updated[rowIndex][colIndex] = value;
    setExcelData(updated);
  };

  // Save Edited Excel
  const saveExcel = async () => {
  try {
    const ws = XLSX.utils.aoa_to_sheet(excelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

    const wbout = XLSX.write(wb, { type: "base64", bookType: "xlsx" });

    
    const path =
      Platform.OS === "android"
        ? `${RNFS.DownloadDirectoryPath}/edited_excel.xlsx` // Android → Downloads
        : `${RNFS.DocumentDirectoryPath}/edited_excel.xlsx`; //  iOS → Documents

    await RNFS.writeFile(path, wbout, "base64");

    Alert.alert("Success", `Excel saved at:\n${path}`);

    // Try to open in Excel/Sheets/WPS
    await FileViewer.open(path);
  } catch (err) {
    console.error("Error saving Excel:", err);
  }
};

  // Render Excel Data
  const renderRow = ({ item, index: rowIndex }) => (
    <View style={styles.row}>
      {item.map((cell, colIndex) => (
        <TextInput
          key={colIndex}
          value={cell ? String(cell) : ""}
          onChangeText={(val) => handleEdit(rowIndex, colIndex, val)}
          style={styles.cell}
        />
      ))}
    </View>
  );

  return (
    <View style={styles.container}>
      <TouchableOpacity style={styles.button} onPress={browseAndLoadExcel}>
        <Text style={styles.buttonText}>📂 Browse Excel</Text>
      </TouchableOpacity>

      <FlatList
        data={excelData}
        keyExtractor={(_, i) => i.toString()}
        renderItem={renderRow}
      />

      {excelData.length > 0 && (
        <TouchableOpacity style={styles.saveButton} onPress={saveExcel}>
          <Text style={styles.buttonText}>💾 Save Excel</Text>
        </TouchableOpacity>
      )}
    </View>
  );
}

const styles = StyleSheet.create({
  container: { flex: 1, padding: 10, backgroundColor: "#f9f9f9" },
  button: {
    backgroundColor: "#4CAF50",
    padding: 12,
    borderRadius: 8,
    alignItems: "center",
    marginBottom: 10,
  },
  saveButton: {
    backgroundColor: "#2196F3",
    padding: 12,
    borderRadius: 8,
    alignItems: "center",
    marginTop: 10,
  },
  buttonText: { color: "#fff", fontSize: 16, fontWeight: "bold" },
  row: { flexDirection: "row" },
  cell: {
    borderWidth: 1,
    borderColor: "#ddd",
    padding: 6,
    minWidth: 80,
    textAlign: "center",
    backgroundColor: "#fff",
  },
});
