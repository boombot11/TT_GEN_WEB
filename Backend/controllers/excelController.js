import path from 'path';
import fs from 'fs';
import { exec } from 'child_process';
import archiver from 'archiver';
import AdmZip from 'adm-zip';

// Get the directory name for ES Modules
const __dirname = decodeURIComponent(path.dirname(new URL(import.meta.url).pathname));

const executeExcelMacro = (filePath, macroName, userInputLab, userInputLecture, socket) => {
    return new Promise((resolve, reject) => {
        const psScriptPath = path.join(__dirname, 'runDynamicMacro.ps1').slice(1);
        console.log("Dynamic input lab: " + userInputLab);
        console.log("Dynamic input lecture: " + userInputLecture);

        // Construct the PowerShell command with all necessary arguments
        const command = `powershell -ExecutionPolicy ByPass -File "${psScriptPath}" "${filePath}" "${macroName}" "${userInputLab}" "${userInputLecture}"`;

        console.log('Executing PowerShell script:', command);  // Debug print
        exec(command, (err, stdout, stderr) => {
            if (err) {
                console.error('Error executing PowerShell script:', stderr);
                return reject(err);
            }
            console.log('PowerShell script executed successfully:', stdout); // Debug print
            resolve(stdout.trim());
        });
    });
};

const executeCommand = (command, socket, step) => {
    return new Promise((resolve, reject) => {
        socket.emit('processing-status', step);
        exec(command, (err, stdout, stderr) => {
            if (err) {
                console.error('Error executing PowerShell script:', stderr);
                return reject(err);
            }
            console.log('PowerShell script executed successfully:', stdout); // Debug print
            resolve(stdout);
        });
    });
};

const extractSheets = (filePath, rooms, labs, socket) => {
    return new Promise(async (resolve, reject) => {
        const psScriptPath = path.join(__dirname, 'ExtractSheets.ps1').slice(1);
        const newRoomFilePath = path.join(__dirname, 'gen', 'room.xlsx').slice(1);
        const newLabFilePath = path.join(__dirname, 'gen', 'lab.xlsx').slice(1);
        const newTeacherFilePath = path.join(__dirname, 'gen', 'teachers.xlsx').slice(1);
        const TeacherScript = path.join(__dirname, 'teachersExtract.ps1').slice(1);

        console.log(`Room file path: ${newRoomFilePath}`);
        console.log(`Lab file path: ${newLabFilePath}`);
        console.log(`Teacher file path: ${newTeacherFilePath}`);
        console.log('FilePath:', filePath); // Debug print
        console.log("psScriptPath:", psScriptPath);
        console.log(rooms);
        console.log(labs);

        if (typeof rooms === 'string') {
            rooms = rooms.split(' '); // Split by space into an array
        }
        if (typeof labs === 'string') {
            labs = labs.split(' '); // Split by space into an array
        }

        for (let i = 0; i < Math.max(rooms.length, labs.length); i++) {
            const room = rooms[i] || null;
            const lab = labs[i] || null;

            console.log(`Processing room: ${room}, lab: ${lab}`);

            // Construct the PowerShell command for teacher extraction
            const teacher = `powershell -ExecutionPolicy ByPass -File "${TeacherScript}" -excelFilePath "${filePath}" `
                + `-rooms "${room}" -labs "${lab}" `
                + `-newRoomFilePath "${newRoomFilePath}" `
                + `-newLabFilePath "${newLabFilePath}" `
                + `-newTeacherFilePath "${newTeacherFilePath}"`;

            if (i === 0) {
                await executeCommand(teacher, socket, `Step ${i + 1}: Executing teacher extraction for room ${room} and lab ${lab}`);
            }

            const command = `powershell -ExecutionPolicy ByPass -File "${psScriptPath}" -excelFilePath "${filePath}" `
                + `-rooms "${room}" -labs "${lab}" `
                + `-newRoomFilePath "${newRoomFilePath}" `
                + `-newLabFilePath "${newLabFilePath}" `
                + `-newTeacherFilePath "${newTeacherFilePath}"`;

            try {
                await executeCommand(command, socket, `Step ${i + 2}: Extracting sheets for room ${room} and lab ${lab}`);
            } catch (err) {
                console.error('Error in extracting sheets:', err);
                throw err; // Stop further execution if any command fails
            }
        }

        resolve({ newRoomFilePath, newLabFilePath, newTeacherFilePath });
    });
};

export const uploadExcel = async (req, res) => {
    console.log('Entering uploadExcel function...');  // Debug print
    const socket = req.app.get('socket');  // Assuming socket instance is stored in the app object

    try {
        const file = req.file;  // File uploaded via middleware
        const { classrooms, labs } = req.body;  // Extract the user inputs from the body
        const userInputLab = labs;
        const userInputLecture = classrooms;

        if (!file) {
            console.error('No file uploaded');  // Debug print
            socket.emit('processing-status', 'No file uploaded');
            return res.status(400).json({ error: 'No file uploaded' });
        }

        if (!file.originalname.endsWith('.xlsm')) {
            console.error('Uploaded file is not an .xlsm file');  // Debug print
            socket.emit('processing-status', 'Uploaded file is not an .xlsm file');
            return res.status(400).json({ error: 'Uploaded file must be an .xlsm file containing macros' });
        }

        const tempFilePath = path.join(__dirname, '../uploads', file.filename).slice(1);
        const macroName = 'RunAllModules';  // The macro name to be executed

        console.log('Temp file path:', tempFilePath);  // Debug print
        console.log('Running macro:', macroName);  // Debug print

        // Emit progress: Running the macro
        socket.emit('processing-status', 'Step 1: Running the macro...');

        // Pass the lab and lecture inputs to PowerShell when running the macro
        await executeExcelMacro(tempFilePath, macroName, userInputLab, userInputLecture, socket);
        console.log('Macro executed successfully');  // Debug print

        // Emit progress: Extracting sheets
        socket.emit('processing-status', 'Step 2: Extracting sheets...');

        // Run the extractSheets PowerShell script to create the three .xlsx files
        const { newRoomFilePath, newLabFilePath, newTeacherFilePath } = await extractSheets(tempFilePath, classrooms, labs, socket);
        console.log("After resolving promise");
        console.log(`Room file path: ${newRoomFilePath}`);
        console.log(`Lab file path: ${newLabFilePath}`);
        console.log(`Teacher file path: ${newTeacherFilePath}`);

        // Emit progress: Creating zip file
        socket.emit('processing-status', 'Step 3: Creating ZIP file...');

        // Create the zip file
        const zip = new AdmZip();
        await zip.addLocalFile(newRoomFilePath, 'room.xlsx');
        await zip.addLocalFile(newLabFilePath, 'lab.xlsx');
        await zip.addLocalFile(newTeacherFilePath, 'teachers.xlsx');

        fs.unlink(tempFilePath, (unlinkErr) => {
            if (unlinkErr) console.error('Failed to delete file:', unlinkErr);  // Debug print
        });

        const zipBuffer = zip.toBuffer();

        // Emit progress: Sending zip file
        socket.emit('processing-status', 'Step 4: Sending zip file to client...');

        res.setHeader('Content-Type', 'application/zip');
        res.setHeader('Content-Disposition', 'attachment; filename=extracted_files.zip');
        res.send(zipBuffer);

        console.log('Zip file sent successfully.');
    } catch (err) {
        console.error('Error processing file:', err);  // Debug print
        socket.emit('processing-status', `Error processing file: ${err.message}`);
        res.status(500).json({ error: 'An error occurred while processing the file', details: err.message });
    }
};
