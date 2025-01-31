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



const executeCommand = (tempFilePath, outputWordFilePath) => {
    return new Promise((resolve, reject) => {
        console.log("Entering word function");

        const imageAbove = "C:\\Users\\Sahil\\Documents\\TT_GEN_WEB\\Backend\\controllers\\Header.png"; // Update path
        const imageBelow = "C:\\Users\\Sahil\\Documents\\TT_GEN_WEB\\Backend\\controllers\\Footer.png"; // Update path

        // Path to the PowerShell script
        const psScriptPath = path.join(__dirname, 'WordDoc.ps1').slice(1);    

        // Construct the PowerShell command
        const psCommand = `powershell -ExecutionPolicy Bypass -File "${psScriptPath}" -ExcelFilePath "${tempFilePath}" -outputWordFilePath "${outputWordFilePath}" -ImageAbovePath "${imageAbove}" -ImageBelowPath "${imageBelow}"`;

        console.log('Executing PowerShell script...');

        // Execute the PowerShell command
        exec(psCommand, (err, stdout, stderr) => {
            if (err) {
                console.error('PowerShell script execution error:', err); // Debug print
                reject(new Error(`Execution failed: ${err.message}`)); // Reject the promise on error
            } else if (stderr) {
                console.error('PowerShell script stderr:', stderr); // Debug print
                reject(new Error(`PowerShell script error: ${stderr}`)); // Reject the promise on stderr
            } else {
                console.log('PowerShell script executed successfully:', stdout);  // Debug print
                resolve(stdout); // Resolve the promise when exec finishes successfully
            }
        });
    });
};


// Controller to handle file upload and processing
const uploadExcelStatic = async (req, res) => {
    console.log('Entering uploadExcel function...');  // Debug print
    try {
        const file = req.file; // File uploaded via middleware
        if (!file) {
            console.error('No file uploaded'); // Debug print
            return res.status(400).json({ error: 'No file uploaded' });
        }
        if (!file.originalname.endsWith('.xlsm')) {
            console.error('Uploaded file is not an .xlsm file'); // Debug print
            return res.status(400).json({ error: 'Uploaded file must be an .xlsm file containing macros' });
        }

        const tempFilePath = path.join(__dirname, '../uploads', file.filename).slice(1);
        const macroName = 'RunAllModules'; // Replace with your VBA macro name

        console.log('Temp file path:', tempFilePath);  // Debug print
        console.log('Running macro:', macroName);  // Debug print

        // Run the macro using winax
        executeStaticExcelMacro(tempFilePath, macroName, async (err, result) => {
            if (err) {
                console.error('Macro execution error:', err); // Debug print
                return res.status(500).json({ error: 'Failed to execute macro', details: err.message });
            }
            console.log('Macro executed successfully:', result);  // Debug print
            const outputWordFilePath = path.join(__dirname, '../uploads', 'outputDocument.docx').slice(1);
            console.log(outputWordFilePath)
           await executeCommand(tempFilePath,outputWordFilePath);
            // Create a readable stream from the modified Excel file
            const fileStream = fs.createReadStream(outputWordFilePath);
            console.log('Created file stream');  // Debug print

            // Set the appropriate headers for file download
            res.setHeader('Content-Disposition', `attachment; filename=document.docx`);
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.documen'); // MIME type for .xlsm files
            console.log('Headers set for file download');  // Debug print
            console.log(outputWordFilePath)
            // Pipe the file stream to the response
            fileStream.pipe(res);
            console.log('File streaming started');  // Debug print
            
            // Handle cleanup once the file download is complete
            fileStream.on('end', () => {
                console.log('File streaming completed. Cleaning up...');  // Debug print
                // fs.unlink(outputWordFilePath, (unlinkErr) => {
                //     if (unlinkErr) console.error('Failed to delete file:', unlinkErr);  // Debug print
                // });
            });

            // Handle potential errors during streaming
            fileStream.on('error', (downloadErr) => {
                console.error('Error sending file:', downloadErr);  // Debug print
                // fs.unlink(tempFilePath, (unlinkErr) => {
                //     if (unlinkErr) console.error('Failed to delete file:', unlinkErr);  // Debug print
                // });
                return res.status(500).json({ error: 'Failed to send file', details: downloadErr.message });
            });
        });
    } catch (err) {
        console.error('Error processing file:', err);  // Debug print
        res.status(500).json({ error: 'An error occurred while processing the file', details: err.message });
    }
};

export default uploadExcelStatic;
