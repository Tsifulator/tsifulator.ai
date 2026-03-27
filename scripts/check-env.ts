import "dotenv/config";

console.log("SID:", process.env.TWILIO_ACCOUNT_SID);
console.log("TOKEN_LEN:", process.env.TWILIO_AUTH_TOKEN?.length);
console.log("FROM:", process.env.TWILIO_WHATSAPP_FROM);
console.log("TO:", process.env.YOUR_WHATSAPP_TO);
