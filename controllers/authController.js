const User = require("../models/User");
const bcrypt = require("bcryptjs");
const jwt = require("jsonwebtoken");

exports.register = async (req, res) => {
  try {
    const { name, email, password, employeeId, gender, dateOfBirth, department, contactNo } = req.body;
    const existingUser = await User.findOne({ $or: [{ email }, { name }] });
    if (existingUser) return res.status(400).json({ message: "User already exists" });
    const hashedPassword = await bcrypt.hash(password, 10);
    const user = new User({
      name,
      email,
      password: hashedPassword,
      employeeId,
      gender,
      dateOfBirth,
      department,
      contactNo,
    });
    await user.save();
    res.status(201).json({ message: "User registered successfully" });
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};

exports.login = async (req, res) => {
  try {
    const { email, password } = req.body;
    const user = await User.findOne({ email });
    if (!user) return res.status(400).json({ message: "Invalid credentials" });
    const isMatch = await bcrypt.compare(password, user.password);
    if (!isMatch) return res.status(400).json({ message: "Invalid credentials" });
    const token = jwt.sign({ userId: user._id }, process.env.JWT_SECRET, { expiresIn: "1d" });
    // res.json({ token, user: { id: user._id, name: user.name, email: user.email, avatar: user.avatar } });
    res.json({ token, user });
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
};
