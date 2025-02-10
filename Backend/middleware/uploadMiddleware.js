import multer from 'multer';
import path from 'path';

// Configure storage for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, './uploads'); // Save to the uploads folder
    },
    filename: (req, file, cb) => {
        // Generate unique names for .xlsm, .json, HEADER, and FOOTER files
        const ext = path.extname(file.originalname);
        const baseName = Date.now() + `-${Math.random().toString(36).substr(2, 9)}`; // Unique identifier

        if (file.fieldname === 'file') {
            // For the .xlsm file
            cb(null, `file-${baseName}${ext}`);
        } else if (file.fieldname === 'config') {
            // For the config.json file
            cb(null, `config-${baseName}${ext}`);
        } else if (file.fieldname === 'HEADER') {
            // For the HEADER image
            cb(null, `HEADER_${baseName}${ext}`);
        } else if (file.fieldname === 'FOOTER') {
            // For the FOOTER image
            cb(null, `FOOTER_${baseName}${ext}`);
        } else {
            cb(new Error('Unsupported file type'), false); // Reject unsupported files
        }
    },
});

// File filter for .xlsm, .json, and image files (assuming .png, .jpg, etc.)
const fileFilter = (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (
        ext === '.xlsm' || // Accept .xlsm files
        ext === '.json' || // Accept .json files
        ['.jpg', '.jpeg', '.png'].includes(ext) // Accept image files (header/footer)
    ) {
        cb(null, true); // Accept these file types
    } else {
        cb(new Error('Only .xlsm, .json, .jpg, .jpeg, .png files are allowed'), false); // Reject other file types
    }
};

// Middleware to handle multiple files
const upload = multer({ storage, fileFilter }).fields([
    { name: 'file', maxCount: 1 }, // For the .xlsm file
    { name: 'config', maxCount: 1 }, // For the config.json file
    { name: 'HEADER', maxCount: 1 }, // For the HEADER image
    { name: 'FOOTER', maxCount: 1 }, // For the FOOTER image
]);

export default upload; // Export the middleware
