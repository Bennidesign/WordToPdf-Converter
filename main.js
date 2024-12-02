const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const edge = require('edge');

// Funktion zum Erstellen des Fensters
let win;
function createWindow() {
  win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true
    }
  });

  win.loadFile('index.html');
  win.on('closed', () => {
    win = null;
  });
}

// Initialisierung der Electron-Anwendung
app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

// Word zu PDF Konvertierungsfunktion (über COM-Interop)
const convertToPdf = edge.func({
  source: function () {/*
    using System;
    using System.Runtime.InteropServices;
    using Microsoft.Office.Interop.Word;

    public class WordToPdfConverter {
        public string Invoke(string docPath) {
            Application wordApp = new Application();
            object missing = System.Reflection.Missing.Value;
            object docPathObj = docPath;
            Document doc = wordApp.Documents.Open(ref docPathObj, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            string pdfPath = docPath.Replace(".docx", ".pdf").Replace(".doc", ".pdf");
            doc.SaveAs2(pdfPath, WdSaveFormat.wdFormatPDF);
            doc.Close();
            wordApp.Quit();
            return pdfPath;
        }
    }
  */}
});

// Führe die Konvertierung durch
ipcMain.handle('convert-to-pdf', async (event, filePath) => {
  try {
    const pdfPath = await convertToPdf(filePath);
    return pdfPath;
  } catch (err) {
    console.error('Fehler bei der Konvertierung:', err);
    return null;
  }
});
