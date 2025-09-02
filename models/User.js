const mongoose = require("mongoose");

const UserSchema = new mongoose.Schema({
  name: { type: String, required: true },
  email: { type: String, required: true, unique: true },
  password: { type: String, required: true },
  avatar: { type: String },
  employeeId: { type: String, required: true, unique: true },
  gender: { type: String },
  dateOfBirth: { type: Date },
  department: { type: String, required: true },
  contactNo: { type: String },
  createdAt: { type: Date, default: Date.now },
});

module.exports = mongoose.model("User", UserSchema);
