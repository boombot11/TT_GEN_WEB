import path from 'path';
import fs from 'fs';
import { exec } from 'child_process';

import AdmZip from 'adm-zip';

// Get the directory name for ES Modules
const __dirname = decodeURIComponent(path.dirname(new URL(import.meta.url).pathname));

const executeExcelMacro = (filePath, macroName, userInputLab, userInputLecture) => {
    return new Promise((resolve, reject) => {
        const psScriptPath = path.join(__dirname, 'runDynamicMacro.ps1').slice(1);
        console.log("Dynamic input lab: " + userInputLab);
        console.log("Dynamic input lecture: " + userInputLecture);
        
        // Construct the PowerShell command with all necessary arguments
        const command = `powershell -ExecutionPolicy ByPass -File "${psScriptPath}" "${filePath}" "${macroName}" "${userInputLab}" "${userInputLecture}"`;

        console.log('Executing PowerShell script:', command);  // Debug print
        const child = exec(command, { timeout:  300000 }, (err, stdout, stderr) => { // Timeout set to 60s
            if (err) {
                console.error('❌ PowerShell script failed:', stderr);
                return reject(err);
            }
            console.log('✅ PowerShell script executed successfully:', stdout);
            resolve(stdout.trim());
        });

        child.on('error', (error) => {
            console.error("❌ PowerShell process error:", error);
            reject(error);
        });
    });
};

const executeCommand = (command) => {
    return new Promise((resolve, reject) => {
        const child = exec(command, { timeout:  300000 }, (err, stdout, stderr) => { // Timeout set to 60s
            if (err) {
                console.error('❌ PowerShell script failed:', stderr);
                return reject(err);
            }
            console.log('✅ PowerShell script executed successfully:', stdout);
            resolve(stdout.trim());
        });

        child.on('error', (error) => {
            console.error("❌ PowerShell process error:", error);
            reject(error);
        });
    });
};


const executeWord = (tempFilePath, outputWordFilePath) => {
    return new Promise((resolve, reject) => {
        console.log("Entering word function");

        const imageAbove = path.join(__dirname, 'Header.png').slice(1);// Update path
        const imageBelow = path.join(__dirname, 'Footer.png').slice(1);// Update path

        // Path to the PowerShell script
        const psScriptPath = path.join(__dirname, 'WordDoc.ps1').slice(1);    

        // Construct the PowerShell command
        const psCommand = `powershell -ExecutionPolicy Bypass -File "${psScriptPath}" -ExcelFilePath "${tempFilePath}" -outputWordFilePath "${outputWordFilePath}" -ImageAbovePath "${imageAbove}" -ImageBelowPath "${imageBelow}"`;

        console.log('Executing PowerShell script...');

        // Execute the PowerShell command
        const child = exec(psCommand, { timeout:  300000 }, (err, stdout, stderr) => { // Timeout set to 60s
            if (err) {
                console.error('❌ PowerShell script failed:', stderr);
                return reject(err);
            }
            console.log('✅ PowerShell script executed successfully:', stdout);
            resolve(stdout.trim());
        });

        child.on('error', (error) => {
            console.error("❌ PowerShell process error:", error);
            reject(error);
        });
    });
};



const extractSheets = (filePath, rooms, labs) => {
    return new Promise (async(resolve, reject) => {
        const psScriptPath = path.join(__dirname, 'ExtractSheets.ps1').slice(1);
        const newRoomFilePath = path.join(__dirname, 'gen', 'room.xlsx').slice(1);
        const newLabFilePath = path.join(__dirname, 'gen', 'lab.xlsx').slice(1);
        const newTeacherFilePath = path.join(__dirname, 'gen', 'teachers.xlsx').slice(1);
       const TeacherScript=path.join(__dirname, 'teachersExtract.ps1').slice(1)
        console.log(`Room file path: ${newRoomFilePath}`);
        console.log(`Lab file path: ${newLabFilePath}`);
        console.log(`Teacher file path: ${newTeacherFilePath}`);
        console.log('FilePath:', filePath); // Debug print
        console.log("psScriptPath:", psScriptPath);
        console.log(rooms);
        console.log(labs)
        if (typeof rooms === 'string') {
            rooms = rooms.split(' '); // Split by space into an array
        }
        if (typeof labs === 'string') {
            labs = labs.split(' '); // Split by space into an array
        }
        

        for (let i = 0; i < Math.max(rooms.length, labs.length); i++) {
           
            const room = rooms[i] || null;  // If there's no room at this index, pass null
            const lab = labs[i] || null;    // If there's no lab at this index, pass null
            console.log(room, lab);
            console.log(i);
            // Construct the PowerShell command for each iteration
            const teacher = `powershell -ExecutionPolicy ByPass -File "${TeacherScript}" -excelFilePath "${filePath}" `
            + `-rooms "${room}" -labs "${lab}" `
            + `-newRoomFilePath "${newRoomFilePath}" `
            + `-newLabFilePath "${newLabFilePath}" `
            + `-newTeacherFilePath "${newTeacherFilePath}"`;
            if(i==0){
                await executeCommand(teacher);
                    }
            const command = `powershell -ExecutionPolicy ByPass -File "${psScriptPath}" -excelFilePath "${filePath}" `
            + `-rooms "${room}" -labs "${lab}" `
            + `-newRoomFilePath "${newRoomFilePath}" `
            + `-newLabFilePath "${newLabFilePath}" `
            + `-newTeacherFilePath "${newTeacherFilePath}"`;
                try {
            
                    await executeCommand(command);
                } catch (err) {
                    console.error('Error in extracting sheets:', err);
                    throw err; // Stop further execution if any command fails
                }
        }        
        // console.log('Executing PowerShell script to extract sheets:', command);
      
        resolve({ newRoomFilePath, newLabFilePath, newTeacherFilePath });
    });
};

export const uploadExcel = async (req, res) => {
    console.log('Entering uploadExcel function...');  // Debug print

    try {
        const file = req.file; // File uploaded via middleware
        const { classrooms, labs } = req.body; // Extract the user inputs from the body
        const userInputLab = labs;
        const userInputLecture = classrooms;

        if (!file) {
            console.error('No file uploaded'); // Debug print
            return res.status(400).json({ error: 'No file uploaded' });
        }

        if (!file.originalname.endsWith('.xlsm')) {
            console.error('Uploaded file is not an .xlsm file'); // Debug print
            return res.status(400).json({ error: 'Uploaded file must be an .xlsm file containing macros' });
        }
     const outputWordFilePath = path.join(__dirname, '../uploads', 'outputDocument.docx').slice(1);
            console.log(outputWordFilePath)
        const tempFilePath = path.join(__dirname, '../uploads', file.filename).slice(1);
        const macroName = 'RunAllWeb'; // The macro name to be executed

        console.log('Temp file path:', tempFilePath);  // Debug print
        console.log('Running macro:', macroName);  // Debug print

        // Pass the lab and lecture inputs to PowerShell when running the macro
        await executeExcelMacro(tempFilePath, macroName, userInputLab, userInputLecture);
        console.log('Macro executed successfully');  // Debug print

        // Run the extractSheets PowerShell script to create the three .xlsx files
        const { newRoomFilePath, newLabFilePath, newTeacherFilePath } = await extractSheets(tempFilePath, classrooms, labs);
        console.log("After resolving promise");
        console.log(`Room file path: ${newRoomFilePath}`);
        console.log(`Lab file path: ${newLabFilePath}`);
        console.log(`Teacher file path: ${newTeacherFilePath}`);
        await executeWord(tempFilePath,outputWordFilePath);
        // Check that the files exist before streaming them
        if (fs.existsSync(newRoomFilePath)) {
            console.log("Room file exists:", newRoomFilePath);
        } else {
            console.log("Room file does NOT exist:", newRoomFilePath);
        }

        if (fs.existsSync(newLabFilePath)) {
            console.log("Lab file exists:", newLabFilePath);
        } else {
            console.log("Lab file does NOT exist:", newLabFilePath);
        }

        if (fs.existsSync(newTeacherFilePath)) {
            console.log("Teacher file exists:", newTeacherFilePath);
        } else {
            console.log("Teacher file does NOT exist:", newTeacherFilePath);
        }

        // Create file streams for the three files
  
         
        console.log('Created file streams for room, lab, and teachers .xlsx');  // Debug print

        // Set the appropriate headers for file download
        res.setHeader('Content-Type', 'application/zip'); // We'll send a zip of all three files
        res.setHeader('Content-Disposition', 'attachment; filename="extracted_files.zip"');  // Zip download
        
        
        const zip = new AdmZip(); // Using archiver to zip the files
        const addFileToZipWithDelay = (filePath, name, delay) => {
            return new Promise((resolve) => {
                setTimeout(() => {
                    if (fs.existsSync(filePath)) {
                        zip.addLocalFile(filePath, undefined, name);
                        console.log(`Added ${name} to zip.`);
                    } else {
                        console.error(`${name} does NOT exist at ${filePath}`);
                    }
                    resolve();
                }, delay);
            });
        };

        // Add files one by one with a 2-second delay
        await addFileToZipWithDelay(newRoomFilePath, 'room.xlsx', 0);
        await addFileToZipWithDelay(newLabFilePath, 'lab.xlsx', 2000); // 2-second delay
        await addFileToZipWithDelay(newTeacherFilePath, 'teachers.xlsx', 4000); // Another 2-second delay
        await addFileToZipWithDelay(outputWordFilePath, 'outputDocument.docx', 6000);
        // Generate zip and send it in response
        const zipBuffer = zip.toBuffer();
        res.setHeader('Content-Disposition', 'attachment; filename=extracted_files.zip');
        res.setHeader('Content-Type', 'application/zip');
        res.send(zipBuffer);

        console.log('Zip file sent successfully.');
        fs.unlink(tempFilePath, (unlinkErr) => {
            if (unlinkErr) console.error('Failed to delete file:', unlinkErr);  // Debug print
        });
        fs.unlink(outputWordFilePath, (unlinkErr) => {
            if (unlinkErr) console.error('Failed to delete file:', unlinkErr);  // Debug print
        });
        const files = ['room.xlsx', 'teachers.xlsx', 'lab.xlsx'];
         replaceFiles(files);
    } catch (err) {
        console.error('Error processing file:', err);  // Debug print
        res.status(500).json({ error: 'An error occurred while processing the file', details: err.message });
    }
};



// Function to delete and replace files
function replaceFiles(files) {
  files.forEach(file => {
    const dirA = process.env.DIR_A || path.join(__dirname, 'gen').slice(1); // Default to 'gen' folder in current directory
    const dirB = process.env.DIR_B || path.join(__dirname, '..', 'routes').slice(1); 
    // Delete file in Directory A if it exists
    const fileInA = path.join(dirA, file);
    const fileInB = path.join(dirB, file);
    if (fs.existsSync(fileInA)) {
      fs.unlinkSync(fileInA);
      console.log(`Deleted: ${fileInA}`);
    } else {
      console.log(`File not found in A: ${fileInA}`);
    }

    // Replace with file from Directory B if it exists
    if (fs.existsSync(fileInB)) {
      fs.copyFileSync(fileInB, fileInA);
      console.log(`Replaced: ${fileInA}`);
    } else {
      console.log(`File not found in B: ${fileInB}`);
    }
  });
}
