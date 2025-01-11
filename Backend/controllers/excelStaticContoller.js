import path from 'path';
import fs from 'fs';
import winax from 'winax'; // Import the winax module
import { exec } from 'child_process';

// Get the directory name for ES Modules
const __dirname = decodeURIComponent(path.dirname(new URL(import.meta.url).pathname));

const executeStaticExcelMacro = (filePath, macroName, callback) => {
    const psScriptPath = path.join(__dirname, 'runExcelMacro.ps1').slice(1);
    const command = `powershell -ExecutionPolicy ByPass -File "${psScriptPath}" "${filePath}" "${macroName}"`;

    console.log('Executing PowerShell script:', command);  // Debug print
    exec(command, (err, stdout, stderr) => {
        if (err) {
            console.error('Error executing PowerShell script:', stderr);
            return callback(err);
        }
        console.log('PowerShell script executed successfully:', stdout); // Debug print
        callback(null, stdout.trim());
    });
};



const uploadExcelStatic = async (req, res) => {
    console.log('Entering uploadExcel function...');
    const socket = req.app.get('socket'); // Assuming the socket instance is stored in the app object

    try {
        const file = req.file;
        if (!file) {
            console.error('No file uploaded');
            socket.emit('processing-status', 'No file uploaded');
            return res.status(400).json({ error: 'No file uploaded' });
        }

        if (!file.originalname.endsWith('.xlsm')) {
            console.error('Uploaded file is not an .xlsm file');
            socket.emit('processing-status', 'Uploaded file is not an .xlsm file');
            return res.status(400).json({ error: 'Uploaded file must be an .xlsm file containing macros' });
        }

        const tempFilePath = path.join(__dirname, '../uploads', file.filename).slice(1);
        const macroName = 'RunAllModules';

        // Emit progress: "Step 1 - Request Processed"
        socket.emit('processing-status', 'Request processed. Running macro...');

        executeStaticExcelMacro(tempFilePath, macroName, (err, result) => {
            if (err) {
                console.error('Macro execution error:', err);
                socket.emit('processing-status', `Error executing macro: ${err.message}`);
                return res.status(500).json({ error: 'Failed to execute macro', details: err.message });
            }

            // Emit progress: "Step 2 - Macro Executed"
            socket.emit('processing-status', 'Macro executed successfully. Preparing file for download.');

            const fileStream = fs.createReadStream(tempFilePath);

            // Emit progress: "Step 3 - Streaming Started"
            socket.emit('processing-status', 'Starting file download...');

            res.setHeader('Content-Disposition', `attachment; filename="${file.originalname}"`);
            res.setHeader('Content-Type', 'application/vnd.ms-excel.sheet.macroenabled.12');
            fileStream.pipe(res);

            fileStream.on('end', () => {
                console.log('File streaming completed. Cleaning up...');
                fs.unlink(tempFilePath, (unlinkErr) => {
                    if (unlinkErr) console.error('Failed to delete file:', unlinkErr);
                });
                socket.emit('processing-status', 'File download complete. Cleanup done.');
            });

            fileStream.on('error', (downloadErr) => {
                console.error('Error sending file:', downloadErr);
                fs.unlink(tempFilePath, (unlinkErr) => {
                    if (unlinkErr) console.error('Failed to delete file:', unlinkErr);
                });
                socket.emit('processing-status', `Error sending file: ${downloadErr.message}`);
                return res.status(500).json({ error: 'Failed to send file', details: downloadErr.message });
            });
        });
    } catch (err) {
        console.error('Error processing file:', err);
        socket.emit('processing-status', `Error processing file: ${err.message}`);
        res.status(500).json({ error: 'An error occurred while processing the file', details: err.message });
    }
};

export default uploadExcelStatic