const { S3Client } = require("@aws-sdk/client-s3");

const s3Client = new S3Client({
  region: "us-east-1", // 👈 N. Virginia = us-east-1
  // No need for credentials here if EC2 IAM role is attached
});

module.exports = s3Client;
