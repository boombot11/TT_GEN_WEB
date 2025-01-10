import dotenv from "dotenv";
import app from "./server.js";

dotenv.config({ path: ".env" });

const PORT = process.env.PORT || 5000;

const server = app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});

// Graceful shutdown
process.on("SIGINT", () => {
  console.log("Shutting down gracefully");
  server.close(() => {
    process.exit(0);
  });
});
