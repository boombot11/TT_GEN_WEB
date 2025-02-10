import path from 'path';
import fs from 'fs';
import executeCommand from '../components/CommandTemplate.js';

// Get the directory name for ES Modules
const __dirname = decodeURIComponent(path.dirname(new URL(import.meta.url).pathname));


export const executeExcelMacro = (filePath,macroName, userInputLab, userInputLecture,Track,map) => {
    return new Promise((resolve, reject) => {
        const psScriptPath = path.join(__dirname, 'runDynamicMacro.ps1').slice(1);
        const command = `powershell -ExecutionPolicy ByPass -File "${psScriptPath}" "${filePath}" "${macroName}" "${userInputLab}" "${userInputLecture}" "${Track}" "${map}" `;
        executeCommand(command)
        .then((stdout) => {
            console.log('✅ Excel macro executed successfully:', stdout);
            resolve(stdout);
        })
        .catch((err) => {
            console.error('❌ Excel macro execution failed:', err);
            reject(err);
        });
}); 
};


export const fontResize = (tempFilePath, outputExcelFilePath, imageAbove, imageBelow) => {
    return new Promise((resolve, reject) => {
        console.log("Entering Excel macro function");

      
            // Step 2: Construct the PowerShell command to execute the macro
            const psScriptPath = path.join(__dirname, 'excelMacro.ps1').slice(1);    

            // Construct the PowerShell command to run the Excel macro
            const psCommand = `powershell -ExecutionPolicy Bypass -File "${psScriptPath}" -ExcelFilePath "${tempFilePath}" -OutputExcelFilePath "${outputExcelFilePath}"  -ImageAbovePath "${imageAbove}" -ImageBelowPath "${imageBelow}"`;

            // Step 3: Execute the PowerShell command
            executeCommand(psCommand)
            .then((stdout) => {
                console.log('✅ Excel macro executed successfully:', stdout);
                resolve(stdout);
            })
            .catch((err) => {
                console.error('❌ Excel macro execution failed:', err);
                reject(err);
            });
       
    });
};

export const executeWord = (tempFilePath, outputWordFilePath,imageAbove,imageBelow) => {
    return new Promise((resolve, reject) => {
        console.log("Entering word function");
        const psScriptPath = path.join(__dirname, 'WordDoc.ps1').slice(1);    

        // Construct the PowerShell command
        const psCommand = `powershell -ExecutionPolicy Bypass -File "${psScriptPath}" -ExcelFilePath "${tempFilePath}" -outputWordFilePath "${outputWordFilePath}" -ImageAbovePath "${imageAbove}" -ImageBelowPath "${imageBelow}"`;
        executeCommand(psCommand)
        .then((stdout) => {
            console.log('✅ Excel macro executed successfully:', stdout);
            resolve(stdout);
        })
        .catch((err) => {
            console.error('❌ Excel macro execution failed:', err);
            reject(err);
        });
});

};



export const extractSheets = (filePath, rooms, labs) => {
    return new Promise (async(resolve, reject) => {
        const psScriptPath = path.join(__dirname, 'ExtractSheets.ps1').slice(1);
        const newRoomFilePath = path.join(__dirname, 'gen', 'room.xlsx').slice(1);
        const newLabFilePath = path.join(__dirname, 'gen', 'lab.xlsx').slice(1);
        const newTeacherFilePath = path.join(__dirname, 'gen', 'teachers.xlsx').slice(1);
       const TeacherScript=path.join(__dirname, 'teachersExtract.ps1').slice(1)
        if (typeof rooms === 'string') {
            rooms = rooms.split(' '); // Split by space into an array
        }
        if (typeof labs === 'string') {
            labs = labs.split(' '); // Split by space into an array
        }
        

        for (let i = 0; i < Math.max(rooms.length, labs.length); i++) {   
            const room = rooms[i] || null;  // If there's no room at this index, pass null
            const lab = labs[i] || null;    // If there's no lab at this index, pass null
    

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

export const executeAddOns = (filePath, AddOnEvents, macroName, userInputLab, userInputLecture, Track, map) => {
    return new Promise((resolve, reject) => {
        const psScriptPath = path.join(__dirname, 'addOn.ps1').slice(1);

        const addOnsParam = AddOnEvents.map(addOn => 
            `${addOn.day},${addOn.sheetName},${addOn.content},${addOn.time}`
        ).join(';'); 
        const command = `powershell -ExecutionPolicy ByPass -File "${psScriptPath}" "${filePath}" "${macroName}" "${userInputLab}" "${userInputLecture}" "${Track}" "${map}" "${addOnsParam}"`;

 
        executeCommand(command)
            .then((stdout) => {
                console.log('✅ Excel macro executed successfully:', stdout);
                resolve(stdout);
            })
            .catch((err) => {
                console.error('❌ Excel macro execution failed:', err);
                reject(err);
            });
    });
};
