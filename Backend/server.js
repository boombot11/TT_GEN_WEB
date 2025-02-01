import express from "express";
import cors from "cors";
import router from "./routes/authRoutes.js";

// Ensure it’s running on Windows before starting the server
if (process.platform !== 'win32') {
    console.log('This application is only supported on Windows.');
    process.exit(1); // Exit the app if not on Windows
}

const app = express();

// Middleware
app.use(express.json());
app.use(cors({ origin: '*' }));

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
