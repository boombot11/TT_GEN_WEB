import multer from 'multer';
import path from 'path';

// Configure storage for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, './uploads'); // Save to the uploads folder
    },
    filename: (req, file, cb) => {
        const uniqueName = `${Date.now()}-${file.originalname}`;
        cb(null, uniqueName); // Unique filename to avoid overwriting
    },
});

// File filter for .xlsm and .json files
const fileFilter = (req, file, cb) => {
    const ext = path.extname(file.originalname);
    if (ext === '.xlsm' || ext === '.json') {
        cb(null, true); // Accept .xlsm and .json files
    } else {
        cb(new Error('Only .xlsm and .json files are allowed'), false); // Reject other file types
    }
};

// Middleware to handle multiple files
const upload = multer({ storage, fileFilter }).fields([
    { name: 'file', maxCount: 1 }, // For the .xlsm file
    { name: 'config', maxCount: 1 }, // For the config.json file
]);

export default upload; // Export the middleware
