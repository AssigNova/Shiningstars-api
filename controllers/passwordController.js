const User = require("../models/User");
const bcrypt = require("bcryptjs");
const crypto = require("crypto");
const nodemailer = require("nodemailer");

// In-memory store for OTPs (replace with DB/Redis for production)
const otpStore = {};

function sendOtpEmail(email, otp) {
  // Configure your SMTP transport here
  const transporter = nodemailer.createTransport({
    host: "smtp.hostinger.com", // Hostinger SMTP server
    port: 587, // use 587 (STARTTLS)
    secure: false, // true = port 465, false = other ports
    auth: {
      user: "noreply.itcshiningstars@cosmosevents.in", // your full Hostinger email
      pass: "Cosmos&Assignova@123", // password you set for this mailbox
    },
  });
  const mailOptions = {
    from: process.env.EMAIL_USER,
    to: email,
    subject: "Your OTP for Password Reset",
    text: `Your OTP is: ${otp}`,
  };
  return transporter.sendMail(mailOptions);
}

exports.requestReset = async (req, res) => {
  const { email } = req.body;
  const user = await User.findOne({ email });
  if (!user) return res.status(404).json({ message: "Email not found" });
  const otp = Math.floor(100000 + Math.random() * 900000).toString();
  otpStore[email] = { otp, expires: Date.now() + 1000 * 60 * 10 };
  try {
    await sendOtpEmail(email, otp);
    res.json({ message: "OTP sent to email" });
  } catch {
    res.status(500).json({ message: "Failed to send OTP" });
  }
};

exports.verifyOtp = (req, res) => {
  const { email, otp } = req.body;
  const entry = otpStore[email];
  if (!entry || entry.expires < Date.now()) return res.status(400).json({ message: "OTP invalid" });
  if (entry.otp !== otp) return res.status(400).json({ message: "Invalid OTP" });
  res.json({ message: "OTP verified" });
};

exports.resetPassword = async (req, res) => {
  const { email, password } = req.body;
  const entry = otpStore[email];
  if (!entry || entry.expires < Date.now()) return res.status(400).json({ message: "OTP expired" });
  const user = await User.findOne({ email });
  if (!user) return res.status(404).json({ message: "User not found" });
  user.password = await bcrypt.hash(password, 10);
  await user.save();
  delete otpStore[email];
  res.json({ message: "Password reset successful" });
};
