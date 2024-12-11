const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const edge = require('edge');

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

ipcMain.handle('convert-to-pdf', async (event, filePath) => {
  try {
    const pdfPath = await convertToPdf(filePath);
    return pdfPath;
  } catch (err) {
    console.error('Fehler bei der Konvertierung:', err);
    return null;
  }
});
