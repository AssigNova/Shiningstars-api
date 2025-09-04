const User = require("../models/User");
const Otp = require("../models/Otp");
const bcrypt = require("bcryptjs");
const nodemailer = require("nodemailer");

function sendOtpEmail(email, otp) {
  const transporter = nodemailer.createTransport({
    host: "smtp.hostinger.com",
    port: 587,
    secure: false,
    auth: {
      user: "noreply.itcshiningstars@cosmosevents.in",
      pass: "Cosmos&Assignova@123",
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

  try {
    const user = await User.findOne({ email });
    if (!user) return res.status(404).json({ message: "Email not found" });

    // Delete any existing OTPs for this email
    await Otp.deleteMany({ email });

    const otp = Math.floor(100000 + Math.random() * 900000).toString();
    const expiresAt = new Date(Date.now() + 10 * 60 * 1000); // 10 minutes

    // Create new OTP document
    await Otp.create({
      email,
      otp,
      expiresAt,
    });

    await sendOtpEmail(email, otp);
    res.json({ message: "OTP sent to email" });
  } catch (error) {
    res.status(500).json({ message: "Failed to process request" });
  }
};

exports.verifyOtp = async (req, res) => {
  const { email, otp } = req.body;

  try {
    const otpDoc = await Otp.findOne({ email });

    if (!otpDoc || otpDoc.expiresAt < new Date()) {
      return res.status(400).json({ message: "OTP invalid or expired" });
    }

    if (otpDoc.otp !== otp) {
      return res.status(400).json({ message: "Invalid OTP" });
    }

    // Mark as verified
    otpDoc.verified = true;
    await otpDoc.save();

    res.json({ message: "OTP verified" });
  } catch (error) {
    res.status(500).json({ message: "Failed to verify OTP" });
  }
};

exports.resetPassword = async (req, res) => {
  const { email, password } = req.body;

  try {
    const otpDoc = await Otp.findOne({ email });

    if (!otpDoc || otpDoc.expiresAt < new Date()) {
      return res.status(400).json({ message: "OTP expired" });
    }

    if (!otpDoc.verified) {
      return res.status(400).json({ message: "OTP not verified" });
    }

    const user = await User.findOne({ email });
    if (!user) return res.status(404).json({ message: "User not found" });

    user.password = await bcrypt.hash(password, 10);
    await user.save();

    // Delete the used OTP
    await Otp.deleteOne({ email });

    res.json({ message: "Password reset successful" });
  } catch (error) {
    res.status(500).json({ message: "Failed to reset password" });
  }
};
