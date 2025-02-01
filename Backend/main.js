import { app, BrowserWindow } from 'electron';
import path from 'path';
import './server.js';  // Import the server (this starts the Express server)

console.log("Starting Electron app...");
const __dirname = decodeURIComponent(path.dirname(new URL(import.meta.url).pathname));
function createWindow() {
    console.log("Creating window...");

    const win = new BrowserWindow({
        width: 800,
        height: 600,
        webPreferences: {
            nodeIntegration: true,  // Allow Node.js integration in the renderer process
        }
    });

    win.loadFile(path.join(__dirname, 'public', 'index.html').slice(1))
        .then(() => {
            console.log("Window loaded successfully");
        })
        .catch((err) => {
            console.error("Error loading HTML:", err);
        });

    win.webContents.openDevTools();
}

app.whenReady().then(() => {
    console.log("Electron app is ready");

    // Create the window
    createWindow();

    // macOS specific behavior
    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

// Quit when all windows are closed
app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});
