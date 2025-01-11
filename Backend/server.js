import express from "express";
import cors from "cors";
import http from 'http';
import { Server } from 'socket.io';


import router from "./routes/authRoutes.js";
// import authRoutes from "./routes/authRoutes.js";

const app = express();
const server = http.createServer(app);
const io = new Server(server,{
    cors: {
        origin: '*', // Replace with your frontend's URL
        methods: ['GET', 'POST'],
   
    },
}); // Initialize socket.io

// Store the socket instance in app.locals or app object so that it can be accessed in the controllers
app.set('socket', io);
// Middleware
app.use(express.json());
app.use(cors({ origin: '*' }));
// Routes
// app.use("/auth", authRoutes);
// Routes
app.use((req, res, next) => {
    res.setTimeout(10 * 60 * 1000, () => { // 10 minutes timeout
        console.log('Request timed out.');
        res.status(408).send('Request Timeout');
    });
    next();
});
app.use(router);
export default app;
