import path from 'path';
import fs from 'fs';
import { exec } from 'child_process';
const __dirname = decodeURIComponent(path.dirname(new URL(import.meta.url).pathname));
const executeWord = (tempFilePath, outputWordFilePath) => {
    return new Promise((resolve, reject) => {
        console.log("Entering word function");

        // const imageAbove = "C:\\Users\\Sahil\\Documents\\TT_GEN_WEB\\Backend\\controllers\\Header.png"; // Update path
        // const imageBelow = "C:\\Users\\Sahil\\Documents\\TT_GEN_WEB\\Backend\\controllers\\Footer.png"; // Update path
       const imageAbove=path.join(__dirname, 'Header.png').slice(1);
       const imageBelow=path.join(__dirname, 'Footer.png').slice(1);
        // Path to the PowerShell script
        const psScriptPath = path.join(__dirname, 'WordDoc.ps1').slice(1);    
        console.log(imageAbove)
        console.log(imageBelow)
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
    const newRoomFilePath = path.join(__dirname, 'gen', 'room.xlsx').slice(1);
 const outputWordFilePath = path.join(__dirname, '../uploads', 'outputDocument.docx').slice(1);
 console.log(newRoomFilePath)
    await executeWord(newRoomFilePath,outputWordFilePath);