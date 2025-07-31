import multer from 'multer';
import path from 'path';

// Configure storage for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, './uploads'); // Save to the uploads folder
    },
    filename: (req, file, cb) => {
        // Generate a unique name using the original fieldname as a prefix
        const ext = path.extname(file.originalname);
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        // Creates a filename like: my_data_file-1677654321098-123456789.json
        cb(null, `${file.fieldname}-${uniqueSuffix}${ext}`);
    },
});

// File filter remains the same
const fileFilter = (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (
        ext === '.xlsm' ||
        ext === '.json' ||
        ['.jpg', '.jpeg', '.png'].includes(ext)
    ) {
        cb(null, true); // Accept these file types
    } else {
        cb(new Error('Only .xlsm, .json, and image files are allowed'), false);
    }
};

// Use .any() to accept all files that pass the filter
const upload = multer({ storage, fileFilter }).any();

export default upload;