import "dotenv/config";
import twilio from "twilio";

const client = twilio(process.env.TWILIO_ACCOUNT_SID!, process.env.TWILIO_AUTH_TOKEN!);

client.messages.create({
  from: process.env.TWILIO_WHATSAPP_FROM!,
  to: process.env.YOUR_WHATSAPP_TO!,
  body: "Tsifulator outbound test"
}).then((m) => {
  console.log("OK", m.sid);
}).catch((e) => {
  console.error("FAIL", e);
  process.exit(1);
});
