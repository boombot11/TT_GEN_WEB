import path from 'path';
import fs from 'fs';
import AdmZip from 'adm-zip';
import {executeWord, executeExcelMacro, executeAddOns, extractSheets, fontResize} from './excelFunctions.js';

// Get the directory name for ES Modules
const __dirname = decodeURIComponent(path.dirname(new URL(import.meta.url).pathname));




export const uploadExcel = async (req, res) => {
    console.log('Entering uploadExcel function...');

    try {
        const file = req.files['file'] ? req.files['file'][0] : null; // .xlsm file
       
        const { classrooms, labs,HeaderText, config_data} = req.body; // Extract the user inputs and add-ons from the body
        const headerFile = req.files['HEADER'] ? req.files['HEADER'][0] : null; // HEADER image file
        const footerFile = req.files['FOOTER'] ? req.files['FOOTER'][0] : null; // FOOTER image file
        const headerText = Array.isArray(HeaderText) ? HeaderText.join(', ') : HeaderText;  // Join array of header text or use as string

        var temp=null;
        try{
            temp = JSON.parse(req.body.addOns);
        }catch(e){
            temp=null
        }  // If addOns is a string, parse it into an array
   const addOns=temp
        const userInputLab = labs;
        const userInputLecture = classrooms;
        console.log("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")
        console.log(addOns)
        var imageAbove = null;
        var imageBelow = null;
/*[
  {
    "day": "Monday",
    "sheetName": "64", 
    "content": "Some content related to Monday",
    "time": "8:00-10:00" 
  },
  {
    "day": "Tuesday",
    "sheetName": "65", 
    "content": "Some content related to Tuesday",
    "time": "10:30-12:00" 
  }
]
*/
        // Handle header and footer image files
        if(!headerFile ){
            imageAbove = path.join(__dirname, 'Header.png').slice(1);
        } else {
            imageAbove = path.join(__dirname, '../uploads', headerFile.filename).slice(1);
        }
        if(!footerFile){
            imageBelow = path.join(__dirname, 'Footer.png').slice(1);
        } else {
            imageBelow = path.join(__dirname, '../uploads', footerFile.filename).slice(1);
        }

        if (!file) {
            console.error('No file uploaded');
            return res.status(400).json({ error: 'No .xlsm file uploaded' });
        }

        if (!file.originalname.endsWith('.xlsm')) {
            console.error('Uploaded file is not an .xlsm file');
            return res.status(400).json({ error: 'Uploaded file must be an .xlsm file containing macros' });
        }

        // Read and parse the config.json file
        // const configData = JSON.parse(fs.readFileSync(configFile.path, 'utf-8'));
        const configData = JSON.parse(config_data);
        // Create Track dynamically based on config data (for flexibility)
        const TrackKeys = Object.keys(configData).join(" ");
        const mapValues = Object.values(configData).join(", ");
        console.log(TrackKeys);
        console.log(mapValues);

        // Prepare paths and variables
        const outputWordFilePath = path.join(__dirname, '../uploads', 'outputDocument.docx').slice(1);
        const tempFilePath = path.join(__dirname, '../uploads', file.filename).slice(1);
        const macroName = 'RunAllWeb'; // The macro name to be executed
        const tempWordExcel=path.join(__dirname, '../uploads', 'tempExcel.xlsx').slice(1);
        console.log('Temp file path:', tempFilePath);
        console.log('Running macro:', macroName);
        const AddOnEvents = [];
        await executeExcelMacro(tempFilePath,macroName, userInputLab, userInputLecture, TrackKeys, mapValues);
      
        console.log('Macro executed successfully');
        console.log(addOns)
        console.log('Processing firstttttttttttttt add-ons...');
        if (addOns ) {
            console.log('Processing add-ons...');
            addOns.forEach((addOn, index) => {
                const {day, sheetName, content, time} = addOn;
                console.log(`Add-On ${index + 1}: Day: ${day}, Sheet Name: ${sheetName}, Content: ${content}, Time: ${time}`);
                AddOnEvents.push({ day, sheetName, content, time });
            });
            await executeAddOns(tempFilePath, AddOnEvents, macroName);
        }
        console.log(AddOnEvents)

        // Other code remains the same (checking file, config, etc.)

        // Pass the addOnEvents array to the macro execution function
      
        // Run the extractSheets PowerShell script to create the three .xlsx files
        const { newRoomFilePath, newLabFilePath, newTeacherFilePath } = await extractSheets(tempFilePath, classrooms, labs);
        console.log("After resolving promise");
        console.log(`Room file path: ${newRoomFilePath}`);
        console.log(`Lab file path: ${newLabFilePath}`);
        console.log(`Teacher file path: ${newTeacherFilePath}`);
       await fontResize(tempFilePath,tempWordExcel,imageAbove,imageBelow)
        // Execute Word-related operations
        await executeWord(tempFilePath, outputWordFilePath, imageAbove, imageBelow);

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

        console.log('Created file streams for room, lab, and teachers .xlsx');
          

        // Set the appropriate headers for file download
        res.setHeader('Content-Type', 'application/zip'); // We'll send a zip of all three files
        res.setHeader('Content-Disposition', 'attachment; filename="extracted_files.zip"');  // Zip download

        const zip = new AdmZip(); // Using archiver to zip the files

        // Add files to zip with delay to avoid race conditions
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
        await addFileToZipWithDelay(tempFilePath, 'all_TimeTables.xlsm', 8000);  
        // Generate zip and send it in response
        const zipBuffer = zip.toBuffer();
        res.send(zipBuffer);

        console.log('Zip file sent successfully.');

        // // Clean up temporary files
        // fs.unlink(tempFilePath, (unlinkErr) => {
        //     if (unlinkErr) console.error('Failed to delete file:', unlinkErr);
        // });
        // fs.unlink(configFile.path, (unlinkErr) => {
        //     if (unlinkErr) console.error('Failed to delete file:', unlinkErr);
        // });
        // fs.unlink(outputWordFilePath, (unlinkErr) => {
        //     if (unlinkErr) console.error('Failed to delete file:', unlinkErr);
        // });

        const files = ['room.xlsx', 'teachers.xlsx', 'lab.xlsx'];
        replaceFiles(files);

    } catch (err) {
        console.error('Error processing file:', err);
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
