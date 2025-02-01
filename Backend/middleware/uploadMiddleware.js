import multer from "multer";
import path from 'path';
// Configure storage for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, './uploads');
    },
    filename: (req, file, cb) => {
        const uniqueName = `${Date.now()}-${file.originalname}`;
        cb(null, uniqueName);
    },
});

// File filter for .xlsm files
const fileFilter = (req, file, cb) => {
    const ext = path.extname(file.originalname);
    if (ext === '.xlsm') {
        cb(null, true);
    } else {
        cb(new Error('Only .xlsm files are allowed'), false);
    }
};

const upload = multer({ storage, fileFilter });

export default upload.single('file'); // Middleware for handling single file upload
