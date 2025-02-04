import express from "express";
// import { register, login } from "../controllers/authController.js";
// import { verifyToken } from "../middleware/authMiddleware.js";
import uploadMiddleware from "../middleware/uploadMiddleware.js";
import { uploadExcel} from "../controllers/excelController.js";


const router = express.Router();

// router.post("/register", register);
// router.post("/login", login);
// router.post("/verify-token", verifyToken, (req, res) => {
//   res.status(200).json({
//     success: true,
//     message: "Token is valid",
//     user: req.user,
//   });
// });
router.post('/upload-excel', uploadMiddleware, uploadExcel);

export default router;
