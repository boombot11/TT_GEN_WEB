
import { exec } from 'child_process';

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
export default executeCommand;