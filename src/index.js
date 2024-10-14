const { app, BrowserWindow, dialog, ipcMain, shell } = require("electron");
const path = require("node:path");
const XSLX = require("xlsx");
const fs = require("fs");
const { autoUpdater } = require("electron-updater");

let mainWindow;

if (require("electron-squirrel-startup")) {
  app.quit();
}

const createWindow = () => {
  mainWindow = new BrowserWindow({
    width: 1000,
    height: 750,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      nodeIntegration: false,
      contextIsolation: true,
      enableRemoteModule: false,
    },
  });

  mainWindow.setMenuBarVisibility(false);
  mainWindow.loadFile(path.join(__dirname, "index.html"));

  mainWindow.webContents.once("did-finish-load", () => {
    autoUpdater.checkForUpdatesAndNotify();
  });
};

app.whenReady().then(() => {
  createWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

// Sự kiện tự động cập nhật
autoUpdater.on("update-available", () => {
  dialog.showMessageBox({
    type: "info",
    title: "Update Available",
    message: "A new update is available. Downloading now...",
  });
});

autoUpdater.on("update-downloaded", () => {
  dialog
    .showMessageBox({
      type: "info",
      title: "Update Ready",
      message:
        "A new update has been downloaded. It will be installed on restart.",
      buttons: ["Restart Now", "Later"],
    })
    .then((result) => {
      if (result.response === 0) {
        autoUpdater.quitAndInstall(); // Cài đặt cập nhật ngay
      }
    });
});

ipcMain.handle("openFileDialog", async (event, name, extensions) => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ["openFile"],
    filters: [{ name: name, extensions: extensions }],
  });

  if (result.canceled) {
    return null;
  } else {
    return result.filePaths[0];
  }
});

ipcMain.handle("readExcelFile", async (event, filePath) => {
  if (!filePath) return null;

  try {
    const data = fs.readFileSync(filePath);
    const workbook = XSLX.read(data, { type: "buffer" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XSLX.utils.sheet_to_json(sheet, { header: 1 });
    return jsonData;
  } catch (err) {
    dialog.showMessageBox(mainWindow, {
      type: "error",
      title: "Error",
      message: "An error occurred",
      detail: err.message,
      buttons: ["OK"],
    });
    return null;
  }
});

ipcMain.handle("exportData", async (event, data) => {
  try {
    const excel1Data = data.excel1Data;
    const excel2Data = data.excel2Data;
    const selectedFields = data.selectedFields;
    const foreignKey1 = Number.parseInt(data.foreignKey1);
    const foreignKey2 = Number.parseInt(data.foreignKey2);

    if (
      !excel1Data ||
      !excel2Data ||
      !foreignKey1 ||
      !foreignKey2 ||
      !selectedFields
    )
      return;

    // Merge the two datasets based on the foreign keys
    const mergedData = await mergeTables(
      excel1Data,
      excel2Data,
      foreignKey1,
      foreignKey2,
      selectedFields
    );

    // Open Save Dialog
    const { filePath } = await dialog.showSaveDialog(mainWindow, {
      title: "Save Excel File",
      defaultPath: "exported_data.xlsx",
      filters: [{ name: "Excel Files", extensions: ["xlsx"] }],
    });

    if (!filePath) {
      return; // If user cancels the save dialog, exit the function.
    }

    // Create new workbook
    const newWorkbook = XSLX.utils.book_new();

    // Convert merged data to worksheet
    const newWorksheet = XSLX.utils.aoa_to_sheet(mergedData);

    // Append worksheet to workbook
    XSLX.utils.book_append_sheet(newWorkbook, newWorksheet, "Sheet1");

    // Write the file
    XSLX.writeFile(newWorkbook, filePath);

    // Inform the user that the file was saved successfully
    dialog.showMessageBox(mainWindow, {
      type: "info",
      title: "Success",
      message: "Excel file has been saved successfully.",
      buttons: ["OK"],
    });

    await shell.openPath(filePath);
  } catch (err) {
    // Display error message in case of failure
    dialog.showMessageBox(mainWindow, {
      type: "error",
      title: "Error",
      message: "An error occurred while exporting the data.",
      detail: err.message,
      buttons: ["OK"],
    });
  }
});

async function mergeTables(
  excel1Data,
  excel2Data,
  foreignKey1,
  foreignKey2,
  selectedFields
) {
  const mergedData = [];

  // Add headers (the first row will contain the headers selected by the user)
  mergedData.push(selectedFields);

  // Map excel2Data by foreign key for quick lookup
  const excel2Map = new Map();
  excel2Data.slice(1).forEach((row) => {
    excel2Map.set(row[foreignKey2], row);
  });

  // Iterate over excel1Data and merge with excel2Data based on foreign keys
  excel1Data.slice(1).forEach((row1) => {
    const matchingRow2 = excel2Map.get(row1[foreignKey1]);

    if (matchingRow2) {
      const mergedRow = [];

      selectedFields.forEach((field) => {
        const headerIndex1 = excel1Data[0].indexOf(field);
        const headerIndex2 = excel2Data[0].indexOf(field);

        // Nếu có cột tương ứng trong excel1Data
        if (headerIndex1 !== -1 && row1[headerIndex1] !== undefined) {
          mergedRow.push(row1[headerIndex1]);
        }
        // Nếu không, lấy từ excel2Data
        else if (
          headerIndex2 !== -1 &&
          matchingRow2[headerIndex2] !== undefined
        ) {
          mergedRow.push(matchingRow2[headerIndex2]);
        }
        // Nếu không tìm thấy ở cả hai bảng, thêm giá trị trống
        else {
          mergedRow.push("");
        }
      });

      mergedData.push(mergedRow);
    }
  });

  return mergedData;
}
