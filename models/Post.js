const mongoose = require("mongoose");

const PostSchema = new mongoose.Schema({
  title: { type: String, required: true },
  description: { type: String, required: true },
  category: { type: String, required: true },
  participantType: { type: String, required: true },
  department: { type: String },
  tags: { type: String },
  status: { type: String, default: "published" },
  type: { type: String, enum: ["image", "video"], default: "image" },
  content: { type: String },
  image: { type: String },
  author: {
    name: { type: String, required: true },
    department: { type: String, required: true },
  },
  timestamp: { type: String },
  likes: { type: Number, default: 0 },
  comments: { type: Number, default: 0 },
  createdAt: { type: Date, default: Date.now },
});

module.exports = mongoose.model("Post", PostSchema);
