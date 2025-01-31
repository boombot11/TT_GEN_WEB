import express from "express";
import cors from "cors";
import router from "./routes/authRoutes.js";
// import authRoutes from "./routes/authRoutes.js";

const app = express();

// Middleware
app.use(express.json());
app.use(cors({ origin: '*' }));
// Routes
// app.use("/auth", authRoutes);
// Routes
const PORT = process.env.PORT || 5000; // Use the environment variable for PORT or fallback to 5000
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
app.use((req, res, next) => {
    res.setTimeout(10 * 60 * 1000, () => { // 10 minutes timeout
        console.log('Request timed out.');
        res.status(408).send('Request Timeout');
    });
    next();
});
app.use(router);
export default app;
