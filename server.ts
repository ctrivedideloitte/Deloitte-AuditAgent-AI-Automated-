import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import { fileURLToPath } from "url";
import session from "express-session";
import cookieParser from "cookie-parser";
import axios from "axios";
import { google } from "googleapis";
import "dotenv/config";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json());
  app.use(cookieParser());
  app.use(
    session({
      secret: process.env.SESSION_SECRET || "deloitte-audit-secret",
      resave: false,
      saveUninitialized: false,
      cookie: {
        secure: true,
        sameSite: "none",
        httpOnly: true,
      },
    })
  );

  // --- OAuth Helpers ---
  const getRedirectUri = (type: "microsoft" | "google") => {
    return `${process.env.APP_URL || `http://localhost:${PORT}`}/auth/${type}/callback`;
  };

  // --- Microsoft Auth (Outlook) ---
  app.get("/api/auth/microsoft/url", (req, res) => {
    const params = new URLSearchParams({
      client_id: process.env.MICROSOFT_CLIENT_ID!,
      response_type: "code",
      redirect_uri: getRedirectUri("microsoft"),
      response_mode: "query",
      scope: "openid profile email User.Read Mail.Read Mail.Send Files.Read",
    });
    res.json({ url: `https://login.microsoftonline.com/${process.env.MICROSOFT_TENANT_ID || 'common'}/oauth2/v2.0/authorize?${params}` });
  });

  app.get("/api/microsoft/fetch-sharepoint", async (req, res) => {
    const tokens = (req.session as any).msTokens;
    if (!tokens) return res.status(401).json({ error: "Not authenticated with Microsoft" });

    const url = req.query.url as string;
    if (!url) return res.status(400).json({ error: "URL required" });

    try {
      // 1. Convert SharePoint URL to Graph-compatible sharing token
      const encodedUrl = Buffer.from(url).toString('base64').replace(/=/g, '').replace(/\//g, '_').replace(/\+/g, '-');
      const shareToken = `u!${encodedUrl}`;

      // 2. Fetch the file content
      const response = await axios.get(
        `https://graph.microsoft.com/v1.0/shares/${shareToken}/driveItem/content`,
        { 
          headers: { Authorization: `Bearer ${tokens.access_token}` },
          responseType: 'arraybuffer'
        }
      );

      res.send(Buffer.from(response.data));
    } catch (error: any) {
      console.error("SharePoint Fetch Error:", error.response?.data || error.message);
      res.status(500).json({ error: "Failed to fetch from SharePoint. Ensure the file is shared with you." });
    }
  });

  app.get("/auth/microsoft/callback", async (req, res) => {
    const { code } = req.query;
    try {
      const response = await axios.post(
        `https://login.microsoftonline.com/${process.env.MICROSOFT_TENANT_ID || 'common'}/oauth2/v2.0/token`,
        new URLSearchParams({
          client_id: process.env.MICROSOFT_CLIENT_ID!,
          client_secret: process.env.MICROSOFT_CLIENT_SECRET!,
          code: code as string,
          redirect_uri: getRedirectUri("microsoft"),
          grant_type: "authorization_code",
        }),
        { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
      );
      const sess = req.session as any;
      sess.msTokens = response.data;
      res.send(`<html><body><script>window.opener.postMessage({ type: 'MS_AUTH_SUCCESS' }, '*');window.close();</script></body></html>`);
    } catch (error: any) {
      console.error("MS Auth Error:", error.response?.data || error.message);
      res.status(500).send("Auth Failed");
    }
  });

  // --- Google Auth (Sheets) ---
  app.get("/api/auth/google/url", (req, res) => {
    const oauth2Client = new google.auth.OAuth2(
      process.env.GOOGLE_CLIENT_ID,
      process.env.GOOGLE_CLIENT_SECRET,
      getRedirectUri("google")
    );
    const url = oauth2Client.generateAuthUrl({
      access_type: "offline",
      scope: ["https://www.googleapis.com/auth/spreadsheets.readonly", "https://www.googleapis.com/auth/userinfo.profile", "https://www.googleapis.com/auth/userinfo.email"],
      prompt: "consent",
    });
    res.json({ url });
  });

  app.get("/auth/google/callback", async (req, res) => {
    const { code } = req.query;
    const oauth2Client = new google.auth.OAuth2(
      process.env.GOOGLE_CLIENT_ID,
      process.env.GOOGLE_CLIENT_SECRET,
      getRedirectUri("google")
    );
    try {
      const { tokens } = await oauth2Client.getToken(code as string);
      const sess = req.session as any;
      sess.googleTokens = tokens;
      res.send(`<html><body><script>window.opener.postMessage({ type: 'GOOGLE_AUTH_SUCCESS' }, '*');window.close();</script></body></html>`);
    } catch (error) {
      console.error("Google Auth Error:", error);
      res.status(500).send("Auth Failed");
    }
  });

  // --- API Routes ---
  app.get("/api/status", (req, res) => {
    const sess = req.session as any;
    res.json({
      microsoft: !!sess.msTokens,
      google: !!sess.googleTokens,
    });
  });

  app.get("/api/outlook/fetch-attachment", async (req, res) => {
    const sess = req.session as any;
    const tokens = sess.msTokens;
    if (!tokens) return res.status(401).json({ error: "Not authenticated with Microsoft" });

    const subject = req.query.subject as string;
    if (!subject) return res.status(400).json({ error: "Subject required" });

    try {
      // 1. Find the email
      const searchResponse = await axios.get(
        `https://graph.microsoft.com/v1.0/me/messages?$filter=subject eq '${subject}'&$orderby=receivedDateTime desc&$top=1`,
        { headers: { Authorization: `Bearer ${tokens.access_token}` } }
      );

      const message = searchResponse.data.value[0];
      if (!message) return res.status(404).json({ error: "No email matching subject found" });

      // 2. Get attachments
      const attachmentsResponse = await axios.get(
        `https://graph.microsoft.com/v1.0/me/messages/${message.id}/attachments`,
        { headers: { Authorization: `Bearer ${tokens.access_token}` } }
      );

      const attachment = attachmentsResponse.data.value[0];
      if (!attachment) return res.status(404).json({ error: "No attachment found in email" });

      res.json({
        name: attachment.name,
        contentType: attachment.contentType,
        contentBytes: attachment.contentBytes, // Base64
      });
    } catch (error: any) {
      console.error("Outlook Fetch Error:", error.response?.data || error.message);
      res.status(500).json({ error: "Failed to fetch from Outlook" });
    }
  });

  app.get("/api/sheets/fetch", async (req, res) => {
    const sess = req.session as any;
    const tokens = sess.googleTokens;
    if (!tokens) return res.status(401).json({ error: "Not authenticated with Google" });

    const sheetId = req.query.sheetId as string;
    if (!sheetId) return res.status(400).json({ error: "Sheet ID required" });

    try {
      const oauth2Client = new google.auth.OAuth2(
        process.env.GOOGLE_CLIENT_ID,
        process.env.GOOGLE_CLIENT_SECRET,
        getRedirectUri("google")
      );
      oauth2Client.setCredentials(tokens);

      const sheets = google.sheets({ version: "v4", auth: oauth2Client });
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: sheetId,
        range: "A1:Z500", // Fetch first 500 rows
      });

      res.json({ values: response.data.values });
    } catch (error: any) {
      console.error("Sheets Fetch Error:", error.response?.data || error.message);
      res.status(500).json({ error: "Failed to fetch from Google Sheets" });
    }
  });

  app.post("/api/outlook/send", async (req, res) => {
    const sess = req.session as any;
    const tokens = sess.msTokens;
    if (!tokens) return res.status(401).json({ error: "Not authenticated with Microsoft" });

    const { to, subject, body, cc } = req.body;

    try {
      await axios.post(
        "https://graph.microsoft.com/v1.0/me/sendMail",
        {
          message: {
            subject: subject,
            body: {
              contentType: "Text",
              content: body,
            },
            toRecipients: to.split(";").map((email: string) => ({
              emailAddress: { address: email.trim() },
            })),
            ccRecipients: cc ? cc.split(";").map((email: string) => ({
              emailAddress: { address: email.trim() },
            })) : [],
          },
        },
        { headers: { Authorization: `Bearer ${tokens.access_token}` } }
      );
      res.json({ success: true });
    } catch (error: any) {
      console.error("Outlook Send Error:", error.response?.data || error.message);
      res.status(500).json({ error: "Failed to send email" });
    }
  });

  // --- Vite / Production Serving ---
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(path.join(__dirname, "dist")));
    app.get("*", (req, res) => {
      res.sendFile(path.join(__dirname, "dist", "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
