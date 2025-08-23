require("dotenv").config();
const express = require("express");
const cors = require("cors");
const connectDB = require("./config/db");
const path = require("path");

const app = express();
app.use(express.json());
app.use(cors());

// Middleware to log requests
app.use((req, res, next) => {
  const now = new Date();
  console.log(`[${now.toISOString()}] ${req.method} ${req.url}`);
  next(); // pass control to next middleware/route handler
});

// Serve static files from uploads folder
app.use("/public/uploads", express.static(path.join(__dirname, "/public/uploads")));

// Routes
app.use("/api/auth", require("./routes/auth"));
app.use("/api/posts", require("./routes/posts"));

// Connect to MongoDB
connectDB();

app.get("/", (req, res) => {
  res.send("ShiningStar Backend API");
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
