/* ================================================================
   DMS Backend – Premier Energies
   CommonJS Express server  ·  HTTPS  ·  MSSQL  ·  Graph email
   ================================================================ */

// ─── 1. ENV & IMPORTS ───────────────────────────────────────────
const path = require("path");
require("dotenv").config({ path: path.resolve(__dirname, "../.env") });

const fs = require("fs");
const https = require("https");
const crypto = require("crypto");
const express = require("express");
const cors = require("cors");
const cookieParser = require("cookie-parser");
const compression = require("compression");
const jwt = require("jsonwebtoken");
const sql = require("mssql");
const multer = require("multer");
const { v4: uuidv4 } = require("uuid");
const { Client } = require("@microsoft/microsoft-graph-client");
const { ClientSecretCredential } = require("@azure/identity");
require("isomorphic-fetch");

// ─── Config from env ────────────────────────────────────────────
const PORT = Number(process.env.PORT) || 42443;
const HOST = process.env.HOST || "0.0.0.0";
const DEV_BYPASS_AUTH = String(process.env.DMS_DEV_BYPASS_AUTH || "").toLowerCase() === "true";
const DEV_BYPASS_ADMIN = String(process.env.DMS_DEV_BYPASS_ADMIN || "").toLowerCase() === "true";
const DEV_BYPASS_EMAIL = process.env.DMS_DEV_EMAIL || "aarnav.singh@premierenergies.com";
const AUTH_PUBLIC_KEY = fs.readFileSync(
  path.resolve(__dirname, "../", process.env.AUTH_PUBLIC_KEY_FILE || "./server/keys/auth-public.pem"),
  "utf8"
);
const UPLOAD_DIR = path.resolve(
  __dirname,
  "../",
  process.env.UPLOAD_DIR || "./server/uploads"
);
const OBSOLETE_DIR = path.join(UPLOAD_DIR, "obsolete");

// Ensure upload directory exists
if (!fs.existsSync(UPLOAD_DIR)) {
  fs.mkdirSync(UPLOAD_DIR, { recursive: true });
}
if (!fs.existsSync(OBSOLETE_DIR)) {
  fs.mkdirSync(OBSOLETE_DIR, { recursive: true });
}

// ─── 2. MSSQL CONNECTION POOLS ──────────────────────────────────
const dmsConfig = {
  user: process.env.MSSQL_USER || "PEL_DB",
  password: process.env.MSSQL_PASSWORD || "V@aN3#@VaN",
  server: process.env.MSSQL_SERVER || "10.0.50.17",
  port: Number(process.env.MSSQL_PORT) || 1433,
  database: process.env.MSSQL_DB || "dms",
  connectionTimeout: 60000,
  requestTimeout: 60000,
  pool: { max: 10, min: 0, idleTimeoutMillis: 30000 },
  options: { trustServerCertificate: true, encrypt: false },
};

const spotConfig = {
  user: process.env.MSSQL_USER || "PEL_DB",
  password: process.env.MSSQL_PASSWORD || "V@aN3#@VaN",
  server: process.env.MSSQL_SERVER || "10.0.50.17",
  port: Number(process.env.MSSQL_PORT) || 1433,
  database: process.env.MSSQL_DB_SPOT || "SPOT",
  connectionTimeout: 60000,
  requestTimeout: 60000,
  pool: { max: 10, min: 0, idleTimeoutMillis: 30000 },
  options: { trustServerCertificate: true, encrypt: false },
};

let dmsPool;
let spotPool;

async function initPools() {
  try {
    dmsPool = await new sql.ConnectionPool(dmsConfig).connect();
    console.log("[DB] DMS pool connected");
  } catch (err) {
    console.error("[DB] DMS pool error:", err.message);
  }
  try {
    spotPool = await new sql.ConnectionPool(spotConfig).connect();
    console.log("[DB] SPOT pool connected");
  } catch (err) {
    console.error("[DB] SPOT pool error:", err.message);
  }
}

// ─── 3. DATABASE TABLES ─────────────────────────────────────────
async function ensureTables() {
  try {
    const req = dmsPool.request();

    // DmsAdmins
    await req.query(`
      IF OBJECT_ID('DmsAdmins','U') IS NULL
      BEGIN
        CREATE TABLE DmsAdmins (
          Email NVARCHAR(256) PRIMARY KEY,
          Active BIT DEFAULT 1,
          CreatedAt DATETIME2 DEFAULT SYSUTCDATETIME(),
          UpdatedBy NVARCHAR(256)
        );
        INSERT INTO DmsAdmins (Email) VALUES
          ('prakash.chandra@premierenergies.com'),
          ('baskara.pandian@premierenergies.com'),
          ('ramesh.t@premierenergies.com'),
          ('aarnav.singh@premierenergies.com');
      END
    `);

    // DmsDocuments
    await dmsPool.request().query(`
      IF OBJECT_ID('DmsDocuments','U') IS NULL
      BEGIN
        CREATE TABLE DmsDocuments (
          Id INT IDENTITY(1,1) PRIMARY KEY,
          Title NVARCHAR(500) NOT NULL,
          Description NVARCHAR(MAX),
          FileName NVARCHAR(500) NOT NULL,
          FilePath NVARCHAR(1000) NOT NULL,
          FileSize BIGINT,
          MimeType NVARCHAR(256),
          CurrentVersion INT DEFAULT 1,
          CurrentVersionLabel NVARCHAR(20) DEFAULT '1.0',
          IsControlled BIT DEFAULT 0,
          Status NVARCHAR(50) DEFAULT 'active',
          ShareScope NVARCHAR(50) DEFAULT 'private',
          ShareGroupId INT NULL,
          CreatorEmail NVARCHAR(256) NOT NULL,
          CreatorEmpId NVARCHAR(50),
          Department NVARCHAR(256),
          Location NVARCHAR(256),
          FileHash NVARCHAR(128),
          IsObsolete BIT DEFAULT 0,
          ParentDocId INT NULL,
          ApprovalStatus NVARCHAR(50) DEFAULT 'none',
          HodSkipped BIT DEFAULT 0,
          SearchContent NVARCHAR(MAX),
          MetadataJson NVARCHAR(MAX),
          ValidFrom DATE NULL,
          ValidTo DATE NULL,
          ValidityReminderSentAt DATETIME2 NULL,
          CreatedAt DATETIME2 DEFAULT SYSUTCDATETIME(),
          UpdatedAt DATETIME2 DEFAULT SYSUTCDATETIME(),
          CONSTRAINT FK_DmsDocuments_Parent FOREIGN KEY (ParentDocId) REFERENCES DmsDocuments(Id)
        );
        CREATE INDEX IX_DmsDocuments_Creator ON DmsDocuments(CreatorEmail);
        CREATE INDEX IX_DmsDocuments_Status ON DmsDocuments(Status);
        CREATE INDEX IX_DmsDocuments_Dept ON DmsDocuments(Department);
        CREATE INDEX IX_DmsDocuments_Location ON DmsDocuments(Location);
      END
    `);

    await dmsPool.request().query(`
      IF OBJECT_ID('DmsDocuments','U') IS NOT NULL
      BEGIN
        IF COL_LENGTH('DmsDocuments', 'MetadataJson') IS NULL
          ALTER TABLE DmsDocuments ADD MetadataJson NVARCHAR(MAX) NULL;
        IF COL_LENGTH('DmsDocuments', 'CurrentVersionLabel') IS NULL
          ALTER TABLE DmsDocuments ADD CurrentVersionLabel NVARCHAR(20) NULL;
        IF COL_LENGTH('DmsDocuments', 'ValidFrom') IS NULL
          ALTER TABLE DmsDocuments ADD ValidFrom DATE NULL;
        IF COL_LENGTH('DmsDocuments', 'ValidTo') IS NULL
          ALTER TABLE DmsDocuments ADD ValidTo DATE NULL;
        IF COL_LENGTH('DmsDocuments', 'ValidityReminderSentAt') IS NULL
          ALTER TABLE DmsDocuments ADD ValidityReminderSentAt DATETIME2 NULL;
      END
    `);

    // DmsDocVersions
    await dmsPool.request().query(`
      IF OBJECT_ID('DmsDocVersions','U') IS NULL
      BEGIN
        CREATE TABLE DmsDocVersions (
          Id INT IDENTITY(1,1) PRIMARY KEY,
          DocId INT NOT NULL,
          Version INT NOT NULL,
          FileName NVARCHAR(500),
          FilePath NVARCHAR(1000),
          FileSize BIGINT,
          FileHash NVARCHAR(128),
          UploadedBy NVARCHAR(256),
          Reason NVARCHAR(MAX),
          CreatedAt DATETIME2 DEFAULT SYSUTCDATETIME(),
          CONSTRAINT FK_DmsDocVersions_Doc FOREIGN KEY (DocId) REFERENCES DmsDocuments(Id)
        );
      END
    `);

    // DmsApprovals
    await dmsPool.request().query(`
      IF OBJECT_ID('DmsApprovals','U') IS NULL
      BEGIN
        CREATE TABLE DmsApprovals (
          Id INT IDENTITY(1,1) PRIMARY KEY,
          DocId INT NOT NULL,
          Version INT NOT NULL,
          Stage NVARCHAR(50) NOT NULL,
          ApproverEmail NVARCHAR(256) NOT NULL,
          Status NVARCHAR(50) DEFAULT 'pending',
          Comments NVARCHAR(MAX),
          ActionAt DATETIME2,
          CreatedAt DATETIME2 DEFAULT SYSUTCDATETIME(),
          CONSTRAINT FK_DmsApprovals_Doc FOREIGN KEY (DocId) REFERENCES DmsDocuments(Id)
        );
      END
    `);

    // DmsHods
    await dmsPool.request().query(`
      IF OBJECT_ID('DmsHods','U') IS NULL
      BEGIN
        CREATE TABLE DmsHods (
          Id INT IDENTITY(1,1) PRIMARY KEY,
          Location NVARCHAR(256) NOT NULL,
          Department NVARCHAR(256) NOT NULL,
          HodEmail NVARCHAR(256) NOT NULL,
          HodName NVARCHAR(256),
          Active BIT DEFAULT 1,
          UpdatedBy NVARCHAR(256),
          UpdatedAt DATETIME2 DEFAULT SYSUTCDATETIME(),
          CONSTRAINT UQ_DmsHods_LocDept UNIQUE(Location, Department)
        );
      END
    `);

    // DmsShareGroups
    await dmsPool.request().query(`
      IF OBJECT_ID('DmsShareGroups','U') IS NULL
      BEGIN
        CREATE TABLE DmsShareGroups (
          Id INT IDENTITY(1,1) PRIMARY KEY,
          Name NVARCHAR(256) NOT NULL,
          CreatorEmail NVARCHAR(256) NOT NULL,
          CreatedAt DATETIME2 DEFAULT SYSUTCDATETIME()
        );
      END
    `);

    // DmsShareGroupMembers
    await dmsPool.request().query(`
      IF OBJECT_ID('DmsShareGroupMembers','U') IS NULL
      BEGIN
        CREATE TABLE DmsShareGroupMembers (
          GroupId INT NOT NULL,
          Email NVARCHAR(256) NOT NULL,
          CONSTRAINT PK_DmsShareGroupMembers PRIMARY KEY (GroupId, Email),
          CONSTRAINT FK_DmsShareGroupMembers_Group FOREIGN KEY (GroupId) REFERENCES DmsShareGroups(Id)
        );
      END
    `);

    // DmsDocAccess
    await dmsPool.request().query(`
      IF OBJECT_ID('DmsDocAccess','U') IS NULL
      BEGIN
        CREATE TABLE DmsDocAccess (
          DocId INT NOT NULL,
          Email NVARCHAR(256) NOT NULL,
          AccessType NVARCHAR(50) DEFAULT 'view',
          GrantedAt DATETIME2 DEFAULT SYSUTCDATETIME(),
          CONSTRAINT PK_DmsDocAccess PRIMARY KEY (DocId, Email),
          CONSTRAINT FK_DmsDocAccess_Doc FOREIGN KEY (DocId) REFERENCES DmsDocuments(Id)
        );
      END
    `);

    // DmsPublicLinks
    await dmsPool.request().query(`
      IF OBJECT_ID('DmsPublicLinks','U') IS NULL
      BEGIN
        CREATE TABLE DmsPublicLinks (
          Id INT IDENTITY(1,1) PRIMARY KEY,
          DocId INT NOT NULL,
          LinkToken NVARCHAR(128) NOT NULL UNIQUE,
          CreatedBy NVARCHAR(256),
          ExpiresAt DATETIME2,
          Active BIT DEFAULT 1,
          CreatedAt DATETIME2 DEFAULT SYSUTCDATETIME(),
          CONSTRAINT FK_DmsPublicLinks_Doc FOREIGN KEY (DocId) REFERENCES DmsDocuments(Id)
        );
      END
    `);

    // DmsAuditLog
    await dmsPool.request().query(`
      IF OBJECT_ID('DmsAuditLog','U') IS NULL
      BEGIN
        CREATE TABLE DmsAuditLog (
          Id INT IDENTITY(1,1) PRIMARY KEY,
          Action NVARCHAR(100) NOT NULL,
          EntityType NVARCHAR(50),
          EntityId INT,
          UserEmail NVARCHAR(256),
          Reason NVARCHAR(MAX),
          BeforeState NVARCHAR(MAX),
          AfterState NVARCHAR(MAX),
          IpAddress NVARCHAR(50),
          CreatedAt DATETIME2 DEFAULT SYSUTCDATETIME()
        );
        CREATE INDEX IX_DmsAuditLog_Entity ON DmsAuditLog(EntityType, EntityId);
        CREATE INDEX IX_DmsAuditLog_User ON DmsAuditLog(UserEmail);
      END
    `);

    console.log("[DB] Tables ensured");
  } catch (err) {
    console.error("[DB] ensureTables error:", err.message);
  }
}

// ─── 4. MICROSOFT GRAPH EMAIL ───────────────────────────────────
const credential = new ClientSecretCredential(
  process.env.AZURE_TENANT_ID,
  process.env.AZURE_CLIENT_ID,
  process.env.AZURE_CLIENT_SECRET
);

function getGraphClient() {
  return Client.init({
    authProvider: async (done) => {
      try {
        const token = await credential.getToken("https://graph.microsoft.com/.default");
        done(null, token.token);
      } catch (err) {
        done(err, null);
      }
    },
  });
}

async function sendEmail(to, subject, html) {
  try {
    const client = getGraphClient();
    const toRecipients = Array.isArray(to) ? to : [to];
    await client
      .api(`/users/${process.env.SENDER_EMAIL}/sendMail`)
      .post({
        message: {
          subject,
          body: { contentType: "HTML", content: html },
          toRecipients: toRecipients.map((e) => ({
            emailAddress: { address: e },
          })),
        },
        saveToSentItems: false,
      });
    console.log(`[Email] Sent to ${toRecipients.join(", ")}: ${subject}`);
  } catch (err) {
    console.error("[Email] sendEmail error:", err.message);
  }
}

async function sendEmailWithCc(to, cc, subject, html) {
  try {
    const client = getGraphClient();
    const toRecipients = Array.isArray(to) ? to : [to];
    const ccRecipients = Array.isArray(cc) ? cc : [cc];
    await client
      .api(`/users/${process.env.SENDER_EMAIL}/sendMail`)
      .post({
        message: {
          subject,
          body: { contentType: "HTML", content: html },
          toRecipients: toRecipients.map((e) => ({
            emailAddress: { address: e },
          })),
          ccRecipients: ccRecipients.filter(Boolean).map((e) => ({
            emailAddress: { address: e },
          })),
        },
        saveToSentItems: false,
      });
    console.log(`[Email] Sent to ${toRecipients.join(", ")} cc ${ccRecipients.join(", ")}: ${subject}`);
  } catch (err) {
    console.error("[Email] sendEmailWithCc error:", err.message);
  }
}

// ─── 5. AUTH MIDDLEWARE ─────────────────────────────────────────
function getISTDay() {
  const now = new Date();
  const ist = new Date(now.getTime() + 5.5 * 60 * 60 * 1000);
  return ist.toISOString().slice(0, 10);
}

function requireAuth(req, res, next) {
  if (DEV_BYPASS_AUTH) {
    req.user = {
      sub: "dev-user",
      email: DEV_BYPASS_EMAIL,
      roles: [],
      apps: ["dms"],
      day: getISTDay(),
    };
    return next();
  }

  try {
    const token = req.cookies && req.cookies.sso;
    if (!token) {
      return res.status(401).json({ error: "Not authenticated – no SSO cookie" });
    }

    const decoded = jwt.verify(token, AUTH_PUBLIC_KEY, {
      algorithms: ["RS256"],
      issuer: process.env.ISSUER || "auth.premierenergies.com",
      audience: process.env.AUDIENCE || "apps.premierenergies.com",
    });

    // Check IST day matches
    if (decoded.day && decoded.day !== getISTDay()) {
      return res.status(401).json({ error: "Token expired – day mismatch" });
    }

    req.user = decoded; // { sub, email, roles, apps, day, ... }
    next();
  } catch (err) {
    console.error("[Auth] JWT verify error:", err.message);
    return res.status(401).json({ error: "Invalid or expired token" });
  }
}

async function requireAdmin(req, res, next) {
  if (DEV_BYPASS_AUTH && DEV_BYPASS_ADMIN) {
    req.user = {
      sub: "dev-admin",
      email: DEV_BYPASS_EMAIL,
      roles: ["admin"],
      apps: ["dms"],
      day: getISTDay(),
      isAdmin: true,
    };
    return next();
  }

  requireAuth(req, res, async () => {
    try {
      const result = await dmsPool
        .request()
        .input("email", sql.NVarChar, req.user.email)
        .query("SELECT Email FROM DmsAdmins WHERE Email = @email AND Active = 1");
      if (result.recordset.length === 0) {
        return res.status(403).json({ error: "Admin access required" });
      }
      req.user.isAdmin = true;
      next();
    } catch (err) {
      console.error("[Auth] requireAdmin error:", err.message);
      return res.status(500).json({ error: "Admin check failed" });
    }
  });
}

// ─── 6. FILE UPLOAD (MULTER) ────────────────────────────────────
const storage = multer.diskStorage({
  destination: (_req, _file, cb) => cb(null, UPLOAD_DIR),
  filename: (_req, file, cb) => {
    const ext = path.extname(file.originalname);
    cb(null, `${uuidv4()}${ext}`);
  },
});

const upload = multer({
  storage,
  limits: { fileSize: 100 * 1024 * 1024 }, // 100 MB
  fileFilter: (_req, file, cb) => {
    if (!isAllowedExtension(file.originalname)) {
      return cb(new Error("Unsupported file type. Allowed: .docx, .doc, .xlsx, .xls, .pdf"));
    }
    cb(null, true);
  },
});

function computeFileHash(filePath) {
  return new Promise((resolve, reject) => {
    const hash = crypto.createHash("sha256");
    const stream = fs.createReadStream(filePath);
    stream.on("data", (chunk) => hash.update(chunk));
    stream.on("end", () => resolve(hash.digest("hex")));
    stream.on("error", reject);
  });
}

function extractSearchContent(filePath, fallback = "") {
  try {
    const ext = path.extname(filePath || "").toLowerCase();
    if ([".txt", ".md", ".csv", ".json", ".xml", ".log", ".tsv"].includes(ext)) {
      const text = fs.readFileSync(filePath, "utf8").slice(0, 20000);
      return `${fallback} ${text}`.slice(0, 32000);
    }

    const raw = fs.readFileSync(filePath);
    const ascii = raw
      .toString("latin1")
      .replace(/[^\x20-\x7E]+/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 20000);
    return `${fallback} ${ascii}`.slice(0, 32000);
  } catch (_e) {
    return fallback.slice(0, 32000);
  }
}

function moveFileToObsolete(filePath) {
  try {
    if (!filePath || !fs.existsSync(filePath)) return filePath;
    const base = path.basename(filePath);
    const target = path.join(OBSOLETE_DIR, `${Date.now()}-${base}`);
    fs.renameSync(filePath, target);
    return target;
  } catch (_e) {
    return filePath;
  }
}

const ACCESS_TYPES = new Set(["owner", "view_only", "view_print", "view_print_download"]);
function normalizeAccessType(v) {
  const value = String(v || "").trim().toLowerCase();
  if (ACCESS_TYPES.has(value)) return value;
  return "view_only";
}

const ALLOWED_EXTS = new Set([".docx", ".doc", ".xlsx", ".xls", ".pdf"]);
function isAllowedExtension(filename = "") {
  const ext = path.extname(String(filename || "")).toLowerCase();
  return ALLOWED_EXTS.has(ext);
}

function sequenceToVersionLabel(seq) {
  const n = Math.max(1, Number(seq || 1));
  const major = Math.floor((n - 1) / 10) + 1;
  const minor = (n - 1) % 10;
  return `${major}.${minor}`;
}

async function getEmpContextByEmail(email) {
  const result = await spotPool
    .request()
    .input("email", sql.NVarChar, email)
    .query(`
      SELECT TOP 1
        EmpEmail,
        EmpName,
        EmpID,
        ISNULL(Dept, '') AS Department,
        ISNULL(EmpLocation, '') AS Location
      FROM dbo.EMP
      WHERE EmpEmail = @email
    `);
  return result.recordset[0] || null;
}

async function getScopeUsers(scope, groupId, actorEmail) {
  const normalizedScope = String(scope || "private").toLowerCase();
  const users = [];
  if (normalizedScope === "private") {
    const ctx = await getEmpContextByEmail(actorEmail);
    users.push({
      EmpEmail: actorEmail,
      EmpName: ctx?.EmpName || actorEmail,
      EmpID: ctx?.EmpID || null,
      Department: ctx?.Department || "",
      Location: ctx?.Location || "",
    });
    return users;
  }

  if (normalizedScope === "group") {
    const members = await dmsPool.request().input("gid", sql.Int, Number(groupId || 0))
      .query("SELECT Email FROM DmsShareGroupMembers WHERE GroupId = @gid");
    const emails = members.recordset
      .map((x) => String(x.Email || "").trim().toLowerCase())
      .filter(Boolean);
    if (!emails.length) return [];

    const users = [];
    for (const email of emails) {
      const ctx = await getEmpContextByEmail(email);
      if (!ctx) continue;
      users.push({
        EmpEmail: ctx.EmpEmail || email,
        EmpName: ctx.EmpName || email,
        EmpID: ctx.EmpID || null,
        Department: ctx.Department || "",
        Location: ctx.Location || "",
      });
    }
    return users;
  }

  if (normalizedScope === "department") {
    const actor = await getEmpContextByEmail(actorEmail);
    const dept = actor?.Department || "";
    if (!dept) return [];
    const deptUsers = await spotPool.request().input("dept", sql.NVarChar, dept)
      .query(`
        SELECT EmpEmail, EmpName, EmpID,
               ISNULL(Dept, '') AS Department,
               ISNULL(EmpLocation, '') AS Location
        FROM dbo.EMP
        WHERE ActiveFlag = 1 AND ISNULL(Dept, '') = @dept
      `);
    return deptUsers.recordset.filter((x) => x.EmpEmail);
  }

  if (normalizedScope === "company") {
    const all = await spotPool.request().query(`
      SELECT EmpEmail, EmpName, EmpID,
             ISNULL(Dept, '') AS Department,
             ISNULL(EmpLocation, '') AS Location
      FROM dbo.EMP
      WHERE ActiveFlag = 1 AND EmpEmail IS NOT NULL AND EmpEmail LIKE '%@premierenergies.com'
    `);
    return all.recordset.filter((x) => x.EmpEmail);
  }

  return [];
}

// ─── 8. AUDIT LOGGING HELPER ───────────────────────────────────
async function logAudit(action, entityType, entityId, userEmail, reason, beforeState, afterState, ip) {
  try {
    await dmsPool
      .request()
      .input("action", sql.NVarChar, action)
      .input("entityType", sql.NVarChar, entityType || null)
      .input("entityId", sql.Int, entityId || null)
      .input("userEmail", sql.NVarChar, userEmail || null)
      .input("reason", sql.NVarChar, reason || null)
      .input("beforeState", sql.NVarChar, beforeState ? JSON.stringify(beforeState) : null)
      .input("afterState", sql.NVarChar, afterState ? JSON.stringify(afterState) : null)
      .input("ip", sql.NVarChar, ip || null)
      .query(`
        INSERT INTO DmsAuditLog (Action, EntityType, EntityId, UserEmail, Reason, BeforeState, AfterState, IpAddress)
        VALUES (@action, @entityType, @entityId, @userEmail, @reason, @beforeState, @afterState, @ip)
      `);
  } catch (err) {
    console.error("[Audit] logAudit error:", err.message);
  }
}

// ─── Helper: get client IP ──────────────────────────────────────
function getIp(req) {
  return req.headers["x-forwarded-for"] || req.socket.remoteAddress || "";
}

// ─── EXPRESS APP ────────────────────────────────────────────────
const app = express();
app.use(compression());
app.use(cors({ origin: true, credentials: true }));
app.use(cookieParser());
app.use(express.json({ limit: "10mb" }));
app.use(express.urlencoded({ extended: true, limit: "10mb" }));

// ─── AUTH STUB ROUTES ───────────────────────────────────────────
// DMS does NOT issue its own tokens – auth is handled by digi.premierenergies.com.
// These stubs exist so the SPA fallback doesn't serve index.html for /auth/* paths,
// which would cause the AuthContext to think refresh succeeded (200 HTML != JSON).

app.post("/auth/refresh", (_req, res) => {
  // DMS cannot refresh tokens; only digi can. Return 401 so the client
  // knows to redirect to digi for a fresh session.
  return res.status(401).json({ error: "refresh_not_supported", message: "Authenticate via digi.premierenergies.com" });
});

app.post("/auth/logout", (_req, res) => {
  // Clear the SSO cookie on this domain so the browser stops sending it
  const COOKIE_DOMAIN = (process.env.COOKIE_DOMAIN || "").trim();
  const clearOpts = { path: "/", httpOnly: true, secure: true, sameSite: "none" };
  res.clearCookie("sso", clearOpts);
  res.clearCookie("sso", { ...clearOpts, domain: COOKIE_DOMAIN || undefined });
  res.clearCookie("sso_refresh", { ...clearOpts, path: "/auth" });
  res.clearCookie("sso_refresh", { ...clearOpts, path: "/auth", domain: COOKIE_DOMAIN || undefined });
  return res.json({ ok: true });
});

// ─── 7A. SESSION & PROFILE ─────────────────────────────────────

// GET /api/session
app.get("/api/session", requireAuth, async (req, res) => {
  try {
    const email = req.user.email;

    // Check admin status
    const adminResult = await dmsPool
      .request()
      .input("email", sql.NVarChar, email)
      .query("SELECT Email FROM DmsAdmins WHERE Email = @email AND Active = 1");
    const isAdmin = DEV_BYPASS_AUTH && DEV_BYPASS_ADMIN ? true : adminResult.recordset.length > 0;

    // Get employee details from SPOT
    // ── FIX: EMP columns are Dept, EmpLocation, ManagerID (not Department, Location, ReportingManagerID)
    const empResult = await spotPool
      .request()
      .input("email", sql.NVarChar, email)
      .query(`
        SELECT TOP 1
          e.*,
          rm.EmpEmail AS ReportingManagerEmail,
          rm.EmpName AS ReportingManagerName
        FROM dbo.EMP e
        LEFT JOIN dbo.EMP rm ON rm.EmpID = e.ManagerID
        WHERE e.EmpEmail = @email
      `);

    const emp = empResult.recordset[0] || {};

    res.json({
      email: req.user.email,
      sub: req.user.sub,
      roles: req.user.roles,
      apps: req.user.apps,
      day: req.user.day,
      isAdmin,
      empId: emp.EmpID || null,
      empName: emp.EmpName || null,
      department: emp.Dept || null,
      location: emp.EmpLocation || null,
      reportingManagerId: emp.ManagerID || null,
      reportingManagerEmail: emp.ReportingManagerEmail || null,
      reportingManagerName: emp.ReportingManagerName || null,
      employee: emp,
    });
  } catch (err) {
    console.error("[Session] error:", err.message);
    res.status(500).json({ error: "Failed to fetch session" });
  }
});

// GET /api/profile
app.get("/api/profile", requireAuth, async (req, res) => {
  try {
    const email = req.user.email;
    const empResult = await spotPool
      .request()
      .input("email", sql.NVarChar, email)
      .query("SELECT * FROM dbo.EMP WHERE EmpEmail = @email");

    if (empResult.recordset.length === 0) {
      return res.status(404).json({ error: "Employee not found" });
    }

    res.json(empResult.recordset[0]);
  } catch (err) {
    console.error("[Profile] error:", err.message);
    res.status(500).json({ error: "Failed to fetch profile" });
  }
});

// ─── 7B. EMPLOYEE SEARCH ───────────────────────────────────────

// GET /api/employees/search?q=xxx
app.get("/api/employees/search", requireAuth, async (req, res) => {
  try {
    const qRaw = String(req.query.q || "").trim();
    if (qRaw.length < 1) return res.json([]);
    const result = await spotPool
      .request()
      .input("q", sql.NVarChar, `%${qRaw}%`)
      .query(`
        SELECT TOP 30 EmpID, EmpName, EmpEmail, Dept AS Department, EmpLocation AS Location
        FROM dbo.EMP
        WHERE (
          EmpName LIKE @q
          OR EmpEmail LIKE @q
          OR EmpID LIKE @q
          OR COALESCE(Dept,'') LIKE @q
          OR COALESCE(EmpLocation,'') LIKE @q
        )
          AND ActiveFlag = 1
        ORDER BY EmpName
      `);
    res.json(result.recordset);
  } catch (err) {
    console.error("[EmployeeSearch] error:", err.message);
    res.status(500).json({ error: "Search failed" });
  }
});

// ─── 7C. DOCUMENTS ─────────────────────────────────────────────

// POST /api/documents/extract-metadata
app.post("/api/documents/extract-metadata", requireAuth, upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });
    const text = extractSearchContent(req.file.path, "").slice(0, 6000);
    fs.unlink(req.file.path, () => {});
    res.json({ extractedText: text });
  } catch (err) {
    console.error("[ExtractMetadata] error:", err.message);
    res.status(500).json({ error: "Failed to extract metadata" });
  }
});

// POST /api/documents/upload
app.post("/api/documents/upload", requireAuth, upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded" });
    }

    const { title, description, isControlled, shareScope, shareGroupId, reason, metadata, defaultAccessType, accessControl, validFrom, validTo } = req.body;
    if (!title) {
      return res.status(400).json({ error: "Title is required" });
    }
    if (!validFrom || !validTo) {
      return res.status(400).json({ error: "Validity dates (from/to) are required" });
    }
    if (new Date(validTo) < new Date(validFrom)) {
      return res.status(400).json({ error: "Validity end date must be after start date" });
    }

    const filePath = req.file.path;
    const fileHash = await computeFileHash(filePath);
    const controlled = isControlled === "true" || isControlled === true || isControlled === "1";

    // Get employee info – FIX: use correct EMP columns
    const empResult = await spotPool
      .request()
      .input("email", sql.NVarChar, req.user.email)
      .query("SELECT TOP 1 EmpID, EmpName, Dept, EmpLocation, ManagerID FROM dbo.EMP WHERE EmpEmail = @email");
    const emp = empResult.recordset[0] || {};

    const status = controlled ? "pending_approval" : "active";
    const approvalStatus = controlled ? "pending_rm" : "none";
    const metadataText = typeof metadata === "string" ? metadata : JSON.stringify(metadata || {});
    const searchContent = extractSearchContent(filePath, `${title} ${description || ""} ${req.file.originalname} ${metadataText || ""}`);

    const insertResult = await dmsPool
      .request()
      .input("title", sql.NVarChar, title)
      .input("description", sql.NVarChar, description || null)
      .input("fileName", sql.NVarChar, req.file.originalname)
      .input("filePath", sql.NVarChar, filePath)
      .input("fileSize", sql.BigInt, req.file.size)
      .input("mimeType", sql.NVarChar, req.file.mimetype)
      .input("isControlled", sql.Bit, controlled ? 1 : 0)
      .input("status", sql.NVarChar, status)
      .input("shareScope", sql.NVarChar, shareScope || "private")
      .input("shareGroupId", sql.Int, shareGroupId ? Number(shareGroupId) : null)
      .input("currentVersionLabel", sql.NVarChar, sequenceToVersionLabel(1))
      .input("creatorEmail", sql.NVarChar, req.user.email)
      .input("creatorEmpId", sql.NVarChar, emp.EmpID || null)
      .input("department", sql.NVarChar, emp.Dept || null)
      .input("location", sql.NVarChar, emp.EmpLocation || null)
      .input("fileHash", sql.NVarChar, fileHash)
      .input("approvalStatus", sql.NVarChar, approvalStatus)
      .input("searchContent", sql.NVarChar, searchContent)
      .input("metadataJson", sql.NVarChar, metadataText || null)
      .input("validFrom", sql.Date, validFrom)
      .input("validTo", sql.Date, validTo)
      .query(`
        INSERT INTO DmsDocuments
          (Title, Description, FileName, FilePath, FileSize, MimeType, IsControlled, Status, ShareScope, ShareGroupId, CurrentVersionLabel,
           CreatorEmail, CreatorEmpId, Department, Location, FileHash, ApprovalStatus, SearchContent, MetadataJson, ValidFrom, ValidTo)
        OUTPUT INSERTED.Id
        VALUES
          (@title, @description, @fileName, @filePath, @fileSize, @mimeType, @isControlled, @status, @shareScope, @shareGroupId, @currentVersionLabel,
           @creatorEmail, @creatorEmpId, @department, @location, @fileHash, @approvalStatus, @searchContent, @metadataJson, @validFrom, @validTo)
      `);

    const docId = insertResult.recordset[0].Id;

    // Insert version 1
    await dmsPool
      .request()
      .input("docId", sql.Int, docId)
      .input("version", sql.Int, 1)
      .input("fileName", sql.NVarChar, req.file.originalname)
      .input("filePath", sql.NVarChar, filePath)
      .input("fileSize", sql.BigInt, req.file.size)
      .input("fileHash", sql.NVarChar, fileHash)
      .input("uploadedBy", sql.NVarChar, req.user.email)
      .input("reason", sql.NVarChar, reason || "Initial upload")
      .query(`
        INSERT INTO DmsDocVersions (DocId, Version, FileName, FilePath, FileSize, FileHash, UploadedBy, Reason)
        VALUES (@docId, @version, @fileName, @filePath, @fileSize, @fileHash, @uploadedBy, @reason)
      `);

    // Grant creator access + scope-based recipients with per-user permissions
    await dmsPool
      .request()
      .input("docId", sql.Int, docId)
      .input("email", sql.NVarChar, req.user.email)
      .query(`
        MERGE DmsDocAccess AS t
        USING (SELECT @docId AS DocId, @email AS Email) s
        ON t.DocId=s.DocId AND t.Email=s.Email
        WHEN MATCHED THEN UPDATE SET AccessType='owner'
        WHEN NOT MATCHED THEN INSERT (DocId, Email, AccessType) VALUES (@docId, @email, 'owner');
      `);

    let perUser = [];
    try {
      if (accessControl) {
        perUser = JSON.parse(String(accessControl));
      }
    } catch (_e) {
      perUser = [];
    }
    const perUserMap = new Map(
      (Array.isArray(perUser) ? perUser : [])
        .filter((x) => x && x.email)
        .map((x) => [String(x.email).toLowerCase(), normalizeAccessType(x.accessType)])
    );
    let scopeUsers = [];
    if (String(shareScope || "").toLowerCase() === "selected_users") {
      scopeUsers = Array.from(perUserMap.keys()).map((email) => ({ EmpEmail: email }));
    } else {
      scopeUsers = await getScopeUsers(shareScope || "private", shareGroupId, req.user.email);
    }
    const defaultType = normalizeAccessType(defaultAccessType || "view_only");
    for (const member of scopeUsers) {
      const email = String(member.EmpEmail || "").trim().toLowerCase();
      if (!email || email === String(req.user.email).toLowerCase()) continue;
      const accessType = perUserMap.get(email) || defaultType;
      await dmsPool
        .request()
        .input("docId", sql.Int, docId)
        .input("email", sql.NVarChar, email)
        .input("accessType", sql.NVarChar, accessType)
        .query(`
          MERGE DmsDocAccess AS t
          USING (SELECT @docId AS DocId, @email AS Email) s
          ON t.DocId=s.DocId AND t.Email=s.Email
          WHEN MATCHED THEN UPDATE SET AccessType=@accessType
          WHEN NOT MATCHED THEN INSERT (DocId, Email, AccessType) VALUES (@docId, @email, @accessType);
        `);
    }

    const sharedWith = scopeUsers
      .map((x) => String(x.EmpEmail || "").trim().toLowerCase())
      .filter((x) => x && x !== String(req.user.email).toLowerCase());
    if (sharedWith.length > 0) {
      sendEmail(
        sharedWith,
        `[DMS] New Document Shared: ${title}`,
        `<h3>Document Shared</h3><p>A document has been shared with you.</p><p><strong>Title:</strong> ${title}</p><p><strong>Shared by:</strong> ${req.user.email}</p>`
      );
    }

    // If controlled, create approval for reporting manager
    if (controlled && emp.ManagerID) {
      const rmResult = await spotPool
        .request()
        .input("rmId", sql.NVarChar, emp.ManagerID)
        .query("SELECT TOP 1 EmpEmail, EmpName FROM dbo.EMP WHERE EmpID = @rmId");

      if (rmResult.recordset.length > 0) {
        const rmEmail = rmResult.recordset[0].EmpEmail;
        await dmsPool
          .request()
          .input("docId", sql.Int, docId)
          .input("version", sql.Int, 1)
          .input("stage", sql.NVarChar, "reporting_manager")
          .input("approverEmail", sql.NVarChar, rmEmail)
          .query(`
            INSERT INTO DmsApprovals (DocId, Version, Stage, ApproverEmail)
            VALUES (@docId, @version, @stage, @approverEmail)
          `);

        // Send notification email to RM
        sendEmail(
          rmEmail,
          `[DMS] Approval Required: ${title}`,
          `<h3>Document Approval Required</h3>
           <p><strong>${emp.EmpName || req.user.email}</strong> has uploaded a controlled document that requires your approval.</p>
           <p><strong>Title:</strong> ${title}</p>
           <p><strong>Description:</strong> ${description || "N/A"}</p>
           <p>Please log in to the DMS to review and approve/reject this document.</p>`
        );
      }
    }

    await logAudit("document_upload", "document", docId, req.user.email, reason || "Initial upload", null, { title, controlled, status }, getIp(req));

    res.json({ success: true, id: docId, status, fileHash });
  } catch (err) {
    console.error("[Upload] error:", err.message);
    res.status(500).json({ error: "Upload failed" });
  }
});

// GET /api/documents
app.get("/api/documents", requireAuth, async (req, res) => {
  try {
    const {
      search, department, location, status, isControlled,
      page = 1, pageSize = 20,
      startDate, endDate, creator, version, approver,
      revisionDateFrom, revisionDateTo, title,
      expired,
    } = req.query;

    const pageNum = Math.max(1, Number(page));
    const pageSz = Math.min(100, Math.max(1, Number(pageSize)));
    const offset = (pageNum - 1) * pageSz;

    let whereClauses = ["d.IsObsolete = 0", "d.Status != 'deleted'"];
    const request = dmsPool.request();
    request.input("userEmail", sql.NVarChar, req.user.email);

    // Access control
    const adminCheck = await dmsPool
      .request()
      .input("aEmail", sql.NVarChar, req.user.email)
      .query("SELECT Email FROM DmsAdmins WHERE Email = @aEmail AND Active = 1");
    const isAdmin = adminCheck.recordset.length > 0;

    if (!isAdmin) {
      // Get user's department and location from SPOT – FIX: use correct EMP columns
      const empCheck = await spotPool
        .request()
        .input("eEmail", sql.NVarChar, req.user.email)
        .query("SELECT TOP 1 Dept, EmpLocation FROM dbo.EMP WHERE EmpEmail = @eEmail");
      const userDept = empCheck.recordset[0]?.Dept || "";
      const userLoc = empCheck.recordset[0]?.EmpLocation || "";

      request.input("userDept", sql.NVarChar, userDept);
      request.input("userLoc", sql.NVarChar, userLoc);

      whereClauses.push(`(
        d.CreatorEmail = @userEmail
        OR EXISTS (SELECT 1 FROM DmsDocAccess da WHERE da.DocId = d.Id AND da.Email = @userEmail)
        OR d.ShareScope = 'company'
        OR (d.ShareScope = 'department' AND d.Department = @userDept)
        OR (d.ShareScope = 'group' AND d.ShareGroupId IS NOT NULL AND EXISTS (
          SELECT 1 FROM DmsShareGroupMembers sgm WHERE sgm.GroupId = d.ShareGroupId AND sgm.Email = @userEmail
        ))
      )`);
    }

    if (search) {
      request.input("search", sql.NVarChar, `%${search}%`);
      whereClauses.push("(d.Title LIKE @search OR d.Description LIKE @search OR d.FileName LIKE @search OR d.SearchContent LIKE @search OR d.MetadataJson LIKE @search)");
    }
    if (title) {
      request.input("title", sql.NVarChar, `%${title}%`);
      whereClauses.push("d.Title LIKE @title");
    }
    if (department) {
      request.input("department", sql.NVarChar, department);
      whereClauses.push("d.Department = @department");
    }
    if (location) {
      request.input("location", sql.NVarChar, location);
      whereClauses.push("d.Location = @location");
    }
    if (status) {
      request.input("status", sql.NVarChar, status);
      whereClauses.push("d.Status = @status");
    }
    if (isControlled !== undefined && isControlled !== "") {
      const ctrlVal = isControlled === "true" || isControlled === "1" ? 1 : 0;
      request.input("isControlled", sql.Bit, ctrlVal);
      whereClauses.push("d.IsControlled = @isControlled");
    }
    if (startDate) {
      request.input("startDate", sql.NVarChar, startDate);
      whereClauses.push("d.CreatedAt >= @startDate");
    }
    if (endDate) {
      request.input("endDate", sql.NVarChar, endDate);
      whereClauses.push("d.CreatedAt <= @endDate");
    }
    if (revisionDateFrom) {
      request.input("revisionDateFrom", sql.NVarChar, revisionDateFrom);
      whereClauses.push("d.UpdatedAt >= @revisionDateFrom");
    }
    if (revisionDateTo) {
      request.input("revisionDateTo", sql.NVarChar, revisionDateTo);
      whereClauses.push("d.UpdatedAt <= @revisionDateTo");
    }
    if (creator) {
      request.input("creator", sql.NVarChar, creator);
      whereClauses.push("d.CreatorEmail = @creator");
    }
    if (version) {
      request.input("version", sql.Int, Number(version));
      whereClauses.push("d.CurrentVersion = @version");
    }
    if (approver) {
      request.input("approver", sql.NVarChar, approver);
      whereClauses.push("EXISTS (SELECT 1 FROM DmsApprovals a WHERE a.DocId = d.Id AND a.ApproverEmail = @approver)");
    }
    if (expired === "true") {
      whereClauses.push("d.ValidTo IS NOT NULL AND d.ValidTo < CAST(SYSUTCDATETIME() AS DATE)");
    }
    if (expired === "false") {
      whereClauses.push("(d.ValidTo IS NULL OR d.ValidTo >= CAST(SYSUTCDATETIME() AS DATE))");
    }

    const whereStr = whereClauses.join(" AND ");

    // Count – FIX: use correct EMP columns in subqueries
    const countResult = await dmsPool
      .request()
      .input("userEmail", sql.NVarChar, req.user.email)
      .input("search", sql.NVarChar, search ? `%${search}%` : null)
      .input("title", sql.NVarChar, title ? `%${title}%` : null)
      .input("department", sql.NVarChar, department || null)
      .input("location", sql.NVarChar, location || null)
      .input("status", sql.NVarChar, status || null)
      .input("isControlled", sql.Bit, isControlled !== undefined && isControlled !== "" ? (isControlled === "true" || isControlled === "1" ? 1 : 0) : null)
      .input("startDate", sql.NVarChar, startDate || null)
      .input("endDate", sql.NVarChar, endDate || null)
      .input("revisionDateFrom", sql.NVarChar, revisionDateFrom || null)
      .input("revisionDateTo", sql.NVarChar, revisionDateTo || null)
      .input("creator", sql.NVarChar, creator || null)
      .input("version", sql.Int, version ? Number(version) : null)
      .input("approver", sql.NVarChar, approver || null)
      .input("userDept", sql.NVarChar, isAdmin ? null : ((await spotPool.request().input("eEmail2", sql.NVarChar, req.user.email).query("SELECT TOP 1 Dept FROM dbo.EMP WHERE EmpEmail = @eEmail2")).recordset[0]?.Dept || ""))
      .input("userLoc", sql.NVarChar, isAdmin ? null : ((await spotPool.request().input("eEmail3", sql.NVarChar, req.user.email).query("SELECT TOP 1 EmpLocation FROM dbo.EMP WHERE EmpEmail = @eEmail3")).recordset[0]?.EmpLocation || ""))
      .query(`SELECT COUNT(*) as total FROM DmsDocuments d WHERE ${whereStr}`);

    const total = countResult.recordset[0].total;

    // Fetch page
    request.input("offset", sql.Int, offset);
    request.input("pageSz", sql.Int, pageSz);

    const result = await request.query(`
      SELECT d.*
      FROM DmsDocuments d
      WHERE ${whereStr}
      ORDER BY d.UpdatedAt DESC
      OFFSET @offset ROWS FETCH NEXT @pageSz ROWS ONLY
    `);

    res.json({
      documents: result.recordset,
      total,
      page: pageNum,
      pageSize: pageSz,
      totalPages: Math.ceil(total / pageSz),
    });
  } catch (err) {
    console.error("[Documents] list error:", err.message);
    res.status(500).json({ error: "Failed to fetch documents" });
  }
});

// GET /api/documents/:id
app.get("/api/documents/:id", requireAuth, async (req, res) => {
  try {
    const docId = Number(req.params.id);

    const docResult = await dmsPool
      .request()
      .input("id", sql.Int, docId)
      .query("SELECT * FROM DmsDocuments WHERE Id = @id");

    if (docResult.recordset.length === 0) {
      return res.status(404).json({ error: "Document not found" });
    }

    const doc = docResult.recordset[0];

    // Check access
    const isAdmin = (await dmsPool.request().input("ae", sql.NVarChar, req.user.email)
      .query("SELECT Email FROM DmsAdmins WHERE Email = @ae AND Active = 1")).recordset.length > 0;

    if (!isAdmin && doc.CreatorEmail !== req.user.email) {
      const hasAccess = (await dmsPool.request()
        .input("did", sql.Int, docId)
        .input("ue", sql.NVarChar, req.user.email)
        .query("SELECT 1 FROM DmsDocAccess WHERE DocId = @did AND Email = @ue")).recordset.length > 0;

      if (!hasAccess && doc.ShareScope !== "company") {
        if (doc.ShareScope === "department") {
          // FIX: use correct EMP column
          const userDept = (await spotPool.request().input("ue2", sql.NVarChar, req.user.email)
            .query("SELECT TOP 1 Dept FROM dbo.EMP WHERE EmpEmail = @ue2")).recordset[0]?.Dept;
          if (userDept !== doc.Department) {
            return res.status(403).json({ error: "Access denied" });
          }
        } else if (doc.ShareScope === "group" && doc.ShareGroupId) {
          const inGroup = (await dmsPool.request()
            .input("gid", sql.Int, doc.ShareGroupId)
            .input("ue3", sql.NVarChar, req.user.email)
            .query("SELECT 1 FROM DmsShareGroupMembers WHERE GroupId = @gid AND Email = @ue3")).recordset.length > 0;
          if (!inGroup) {
            return res.status(403).json({ error: "Access denied" });
          }
        } else {
          // Also allow if user is an approver
          const isApprover = (await dmsPool.request()
            .input("did2", sql.Int, docId)
            .input("ue4", sql.NVarChar, req.user.email)
            .query("SELECT 1 FROM DmsApprovals WHERE DocId = @did2 AND ApproverEmail = @ue4")).recordset.length > 0;
          if (!isApprover) {
            return res.status(403).json({ error: "Access denied" });
          }
        }
      }
    }

    // Get versions
    const versions = await dmsPool
      .request()
      .input("docId", sql.Int, docId)
      .query("SELECT * FROM DmsDocVersions WHERE DocId = @docId ORDER BY Version DESC");

    // Get approvals
    const approvals = await dmsPool
      .request()
      .input("docId2", sql.Int, docId)
      .query("SELECT * FROM DmsApprovals WHERE DocId = @docId2 ORDER BY CreatedAt DESC");

    // Get access list
    const accessList = await dmsPool
      .request()
      .input("docId3", sql.Int, docId)
      .query("SELECT * FROM DmsDocAccess WHERE DocId = @docId3");

    // Get public links
    const publicLinks = await dmsPool
      .request()
      .input("docId4", sql.Int, docId)
      .query("SELECT * FROM DmsPublicLinks WHERE DocId = @docId4 AND Active = 1");

    await logAudit("document_view", "document", docId, req.user.email, null, null, null, getIp(req));

    res.json({
      document: doc,
      versions: versions.recordset,
      approvals: approvals.recordset,
      accessList: accessList.recordset,
      publicLinks: publicLinks.recordset,
    });
  } catch (err) {
    console.error("[Document] get error:", err.message);
    res.status(500).json({ error: "Failed to fetch document" });
  }
});

async function getEffectiveAccessType(docId, email) {
  const docResult = await dmsPool
    .request()
    .input("id", sql.Int, docId)
    .query("SELECT CreatorEmail, ShareScope, ShareGroupId, Department FROM DmsDocuments WHERE Id = @id AND Status != 'deleted'");
  if (!docResult.recordset.length) return null;
  const doc = docResult.recordset[0];
  if (String(doc.CreatorEmail || "").toLowerCase() === String(email || "").toLowerCase()) return "owner";

  const explicit = await dmsPool.request()
    .input("docId", sql.Int, docId)
    .input("email", sql.NVarChar, email)
    .query("SELECT TOP 1 AccessType FROM DmsDocAccess WHERE DocId = @docId AND Email = @email");
  if (explicit.recordset.length) return normalizeAccessType(explicit.recordset[0].AccessType);

  // Backward compatibility for legacy documents with scope-only access
  if (doc.ShareScope === "company") return "view_only";
  if (doc.ShareScope === "department") {
    const ctx = await getEmpContextByEmail(email);
    if (ctx?.Department && ctx.Department === doc.Department) return "view_only";
  }
  if (doc.ShareScope === "group" && doc.ShareGroupId) {
    const inGroup = await dmsPool.request()
      .input("gid", sql.Int, doc.ShareGroupId)
      .input("email", sql.NVarChar, email)
      .query("SELECT TOP 1 1 AS ok FROM DmsShareGroupMembers WHERE GroupId = @gid AND Email = @email");
    if (inGroup.recordset.length) return "view_only";
  }
  return null;
}

// GET /api/documents/:id/download
app.get("/api/documents/:id/download", requireAuth, async (req, res) => {
  try {
    const docId = Number(req.params.id);
    const docResult = await dmsPool
      .request()
      .input("id", sql.Int, docId)
      .query("SELECT FileName, FilePath, MimeType FROM DmsDocuments WHERE Id = @id AND Status != 'deleted'");

    if (docResult.recordset.length === 0) {
      return res.status(404).json({ error: "Document not found" });
    }

    const doc = docResult.recordset[0];

    const canAccess = await checkDocAccess(docId, req.user.email);
    if (!canAccess) {
      return res.status(403).json({ error: "Access denied" });
    }

    const accessType = await getEffectiveAccessType(docId, req.user.email);
    const canDownload = accessType === "owner" || accessType === "view_print_download";
    if (!canDownload) {
      return res.status(403).json({ error: "download_not_allowed_for_your_permission" });
    }

    if (!fs.existsSync(doc.FilePath)) {
      return res.status(404).json({ error: "File not found on disk" });
    }

    await logAudit("document_download", "document", docId, req.user.email, null, null, null, getIp(req));

    res.setHeader("Content-Disposition", `attachment; filename="${encodeURIComponent(doc.FileName)}"`);
    res.setHeader("Content-Type", doc.MimeType || "application/octet-stream");
    fs.createReadStream(doc.FilePath).pipe(res);
  } catch (err) {
    console.error("[Download] error:", err.message);
    res.status(500).json({ error: "Download failed" });
  }
});

// GET /api/documents/:id/view
app.get("/api/documents/:id/view", requireAuth, async (req, res) => {
  try {
    const docId = Number(req.params.id);
    const docResult = await dmsPool
      .request()
      .input("id", sql.Int, docId)
      .query("SELECT FileName, FilePath, MimeType FROM DmsDocuments WHERE Id = @id AND Status != 'deleted'");

    if (docResult.recordset.length === 0) {
      return res.status(404).json({ error: "Document not found" });
    }

    const doc = docResult.recordset[0];

    const canAccess = await checkDocAccess(docId, req.user.email);
    if (!canAccess) {
      return res.status(403).json({ error: "Access denied" });
    }
    const accessType = await getEffectiveAccessType(docId, req.user.email);
    const canPrint = accessType === "owner" || accessType === "view_print" || accessType === "view_print_download";
    const canDownload = accessType === "owner" || accessType === "view_print_download";
    const embedMode = String(req.query.embed || "").trim() === "1";
    const rawMode = String(req.query.raw || "").trim() === "1";

    if (!fs.existsSync(doc.FilePath)) {
      return res.status(404).json({ error: "File not found on disk" });
    }

    await logAudit("document_view_file", "document", docId, req.user.email, null, null, null, getIp(req));

    const isPdf = String(doc.MimeType || "").toLowerCase().includes("pdf");
    if (rawMode) {
      const inlineHeader = String(req.get("x-dms-inline") || "").trim();
      if (!embedMode || inlineHeader !== "1") {
        return res.status(403).json({ error: "raw_view_blocked" });
      }
      res.setHeader("Content-Disposition", `inline; filename="${encodeURIComponent(doc.FileName)}"`);
      res.setHeader("Content-Type", doc.MimeType || "application/octet-stream");
      res.setHeader("Cache-Control", "no-store");
      res.setHeader("X-Content-Type-Options", "nosniff");
      return fs.createReadStream(doc.FilePath).pipe(res);
    }

    if (isPdf) {
      const html = `<!doctype html>
<html><head><meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Secure Viewer</title>
<style>
html,body{margin:0;padding:0;background:#0b1220;color:#f8fafc;font-family:Arial,sans-serif}
.top{position:sticky;top:0;padding:10px 14px;background:#111827;border-bottom:1px solid rgba(255,255,255,.15);font-size:12px;display:flex;align-items:center;justify-content:space-between;gap:8px}
.actions{display:flex;gap:8px;align-items:center}
.btn{background:#1f2937;color:#f8fafc;border:1px solid rgba(255,255,255,.2);padding:6px 10px;border-radius:8px;cursor:pointer}
.btn:hover{background:#374151}
#pages{padding:12px;display:flex;flex-direction:column;gap:12px;align-items:center}
canvas{max-width:100%;height:auto;box-shadow:0 8px 24px rgba(0,0,0,.35);background:white}
</style></head><body>
<div class="top">
  <div>${canPrint || canDownload ? "Permission-based secure preview." : "View-only secure preview. Print and download are disabled."}</div>
  <div class="actions">
    ${canPrint ? `<button class="btn" id="printBtn" type="button">Print</button>` : ""}
    ${canDownload ? `<a class="btn" id="downloadBtn" href="/api/documents/${docId}/download" target="_blank" rel="noreferrer">Download</a>` : ""}
  </div>
</div>
<div id="pages"></div>
<script type="module">
  import * as pdfjsLib from "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.8.69/pdf.min.mjs";
  document.addEventListener("contextmenu", (e)=>e.preventDefault());
  document.addEventListener("keydown", (e)=>{
    const k = e.key.toLowerCase();
    if((e.ctrlKey||e.metaKey)&&["s","c"].includes(k)) e.preventDefault();
    if((e.ctrlKey||e.metaKey)&&k==="p" && ${canPrint ? "false" : "true"}) e.preventDefault();
  });
  pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.8.69/pdf.worker.min.mjs";
  const r = await fetch("/api/documents/${docId}/view?embed=1&raw=1", { credentials: "include", headers: { "x-dms-inline": "1" }});
  if (!r.ok) throw new Error("Unable to load document");
  const bytes = new Uint8Array(await r.arrayBuffer());
  const pdf = await pdfjsLib.getDocument({ data: bytes, disableAutoFetch: true, disableStream: true }).promise;
  const wrap = document.getElementById("pages");
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale: 1.5 });
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    await page.render({ canvasContext: ctx, viewport }).promise;
    wrap.appendChild(canvas);
  }
  const printBtn = document.getElementById("printBtn");
  if (printBtn) printBtn.addEventListener("click", () => window.print());
</script></body></html>`;
      res.setHeader("Content-Type", "text/html; charset=utf-8");
      res.setHeader("Cache-Control", "no-store");
      res.setHeader("X-Frame-Options", "SAMEORIGIN");
      res.setHeader("Content-Security-Policy", "default-src 'self'; script-src 'self' https://cdnjs.cloudflare.com; style-src 'unsafe-inline'; img-src 'self' data: blob:; connect-src 'self'; frame-ancestors 'self'");
      return res.send(html);
    }

    // Non-PDF fallback in iframe.
    const fallbackHtml = `<!doctype html><html><head><meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<style>body{margin:0;font-family:Arial,sans-serif;background:#0b1220;color:#f8fafc;padding:20px}.btn{display:inline-block;margin-top:12px;background:#1f2937;color:#fff;border:1px solid rgba(255,255,255,.2);padding:8px 12px;border-radius:8px;text-decoration:none}</style>
</head><body>
<h3>Inline preview is not supported for this file type.</h3>
<p>File: ${String(doc.FileName || "Document")}</p>
${canDownload ? `<a class="btn" href="/api/documents/${docId}/download" target="_blank" rel="noreferrer">Download File</a>` : `<p>Download is disabled for your access level.</p>`}
</body></html>`;
    res.setHeader("Content-Type", "text/html; charset=utf-8");
    res.setHeader("Cache-Control", "no-store");
    return res.send(fallbackHtml);
  } catch (err) {
    console.error("[View] error:", err.message);
    res.status(500).json({ error: "View failed" });
  }
});

// GET /api/documents/:id/permission
app.get("/api/documents/:id/permission", requireAuth, async (req, res) => {
  try {
    const docId = Number(req.params.id);
    const accessType = await getEffectiveAccessType(docId, req.user.email);
    if (!accessType) return res.status(403).json({ error: "Access denied" });
    res.json({
      accessType,
      canView: true,
      canPrint: accessType === "owner" || accessType === "view_print" || accessType === "view_print_download",
      canDownload: accessType === "owner" || accessType === "view_print_download",
    });
  } catch (err) {
    console.error("[Permission] error:", err.message);
    res.status(500).json({ error: "Failed to fetch permission" });
  }
});

// GET /api/documents/:id/audit
app.get("/api/documents/:id/audit", requireAuth, async (req, res) => {
  try {
    const docId = Number(req.params.id);
    const canAccess = await checkDocAccess(docId, req.user.email);
    if (!canAccess) {
      return res.status(403).json({ error: "Access denied" });
    }

    const logs = await dmsPool
      .request()
      .input("docId", sql.Int, docId)
      .query(`
        SELECT TOP 500 *
        FROM DmsAuditLog
        WHERE EntityType = 'document' AND EntityId = @docId
        ORDER BY CreatedAt DESC
      `);

    res.json({ logs: logs.recordset });
  } catch (err) {
    console.error("[DocAudit] error:", err.message);
    res.status(500).json({ error: "Failed to fetch document history" });
  }
});

// DELETE /api/documents/:id
app.delete("/api/documents/:id", requireAuth, async (req, res) => {
  try {
    const docId = Number(req.params.id);
    const docResult = await dmsPool
      .request()
      .input("id", sql.Int, docId)
      .query("SELECT * FROM DmsDocuments WHERE Id = @id");

    if (docResult.recordset.length === 0) {
      return res.status(404).json({ error: "Document not found" });
    }

    const doc = docResult.recordset[0];

    const isAdmin = (await dmsPool.request().input("ae", sql.NVarChar, req.user.email)
      .query("SELECT Email FROM DmsAdmins WHERE Email = @ae AND Active = 1")).recordset.length > 0;

    if (!isAdmin && doc.CreatorEmail !== req.user.email) {
      return res.status(403).json({ error: "Only creator or admin can delete" });
    }

    await dmsPool
      .request()
      .input("id", sql.Int, docId)
      .query("UPDATE DmsDocuments SET Status = 'deleted', UpdatedAt = SYSUTCDATETIME() WHERE Id = @id");

    await logAudit("document_delete", "document", docId, req.user.email, req.body.reason || null, { status: doc.Status }, { status: "deleted" }, getIp(req));

    res.json({ success: true });
  } catch (err) {
    console.error("[Delete] error:", err.message);
    res.status(500).json({ error: "Delete failed" });
  }
});

// POST /api/documents/:id/new-version
app.post("/api/documents/:id/new-version", requireAuth, upload.single("file"), async (req, res) => {
  try {
    const docId = Number(req.params.id);
    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded" });
    }

    const docResult = await dmsPool
      .request()
      .input("id", sql.Int, docId)
      .query("SELECT * FROM DmsDocuments WHERE Id = @id AND IsControlled = 1 AND Status != 'deleted'");

    if (docResult.recordset.length === 0) {
      return res.status(404).json({ error: "Controlled document not found" });
    }

    const doc = docResult.recordset[0];

    const isAdmin = (await dmsPool.request().input("ae", sql.NVarChar, req.user.email)
      .query("SELECT Email FROM DmsAdmins WHERE Email = @ae AND Active = 1")).recordset.length > 0;
    if (!isAdmin && doc.CreatorEmail !== req.user.email) {
      return res.status(403).json({ error: "Only creator or admin can upload new versions" });
    }

    const newVersion = doc.CurrentVersion + 1;
    const newVersionLabel = sequenceToVersionLabel(newVersion);
    const filePath = req.file.path;
    const fileHash = await computeFileHash(filePath);
    const { reason, metadata, validFrom, validTo } = req.body;
    if (!reason || !String(reason).trim()) {
      return res.status(400).json({ error: "Revision reason is required for controlled versioning" });
    }
    const resolvedValidFrom = validFrom || doc.ValidFrom;
    const resolvedValidTo = validTo || doc.ValidTo;
    if (!resolvedValidFrom || !resolvedValidTo) {
      return res.status(400).json({ error: "Validity dates (from/to) are required" });
    }

    // Mark old document as obsolete
    const obsoletePath = moveFileToObsolete(doc.FilePath);
    await dmsPool
      .request()
      .input("id", sql.Int, docId)
      .input("obsoletePath", sql.NVarChar, obsoletePath)
      .query("UPDATE DmsDocuments SET IsObsolete = 1, FilePath = @obsoletePath, UpdatedAt = SYSUTCDATETIME() WHERE Id = @id");

    // Get employee info – FIX: use correct EMP columns
    const empResult = await spotPool
      .request()
      .input("email", sql.NVarChar, req.user.email)
      .query("SELECT TOP 1 EmpID, EmpName, Dept, EmpLocation, ManagerID FROM dbo.EMP WHERE EmpEmail = @email");
    const emp = empResult.recordset[0] || {};

    const metadataText = typeof metadata === "string" ? metadata : JSON.stringify(metadata || {});
    const searchContent = extractSearchContent(filePath, `${doc.Title} ${doc.Description || ""} ${req.file.originalname} ${metadataText || ""}`);

    // Create new document entry for this version
    const insertResult = await dmsPool
      .request()
      .input("title", sql.NVarChar, doc.Title)
      .input("description", sql.NVarChar, doc.Description)
      .input("fileName", sql.NVarChar, req.file.originalname)
      .input("filePath", sql.NVarChar, filePath)
      .input("fileSize", sql.BigInt, req.file.size)
      .input("mimeType", sql.NVarChar, req.file.mimetype)
      .input("currentVersion", sql.Int, newVersion)
      .input("currentVersionLabel", sql.NVarChar, newVersionLabel)
      .input("isControlled", sql.Bit, 1)
      .input("status", sql.NVarChar, "pending_approval")
      .input("shareScope", sql.NVarChar, doc.ShareScope)
      .input("shareGroupId", sql.Int, doc.ShareGroupId)
      .input("creatorEmail", sql.NVarChar, req.user.email)
      .input("creatorEmpId", sql.NVarChar, emp.EmpID || doc.CreatorEmpId)
      .input("department", sql.NVarChar, emp.Dept || doc.Department)
      .input("location", sql.NVarChar, emp.EmpLocation || doc.Location)
      .input("fileHash", sql.NVarChar, fileHash)
      .input("parentDocId", sql.Int, docId)
      .input("approvalStatus", sql.NVarChar, "pending_rm")
      .input("searchContent", sql.NVarChar, searchContent)
      .input("metadataJson", sql.NVarChar, metadataText || doc.MetadataJson || null)
      .input("validFrom", sql.Date, resolvedValidFrom)
      .input("validTo", sql.Date, resolvedValidTo)
      .query(`
        INSERT INTO DmsDocuments
          (Title, Description, FileName, FilePath, FileSize, MimeType, CurrentVersion, IsControlled,
           Status, ShareScope, ShareGroupId, CreatorEmail, CreatorEmpId, Department, Location,
           FileHash, ParentDocId, ApprovalStatus, SearchContent, MetadataJson, CurrentVersionLabel, ValidFrom, ValidTo)
        OUTPUT INSERTED.Id
        VALUES
          (@title, @description, @fileName, @filePath, @fileSize, @mimeType, @currentVersion, @isControlled,
           @status, @shareScope, @shareGroupId, @creatorEmail, @creatorEmpId, @department, @location,
           @fileHash, @parentDocId, @approvalStatus, @searchContent, @metadataJson, @currentVersionLabel, @validFrom, @validTo)
      `);

    const newDocId = insertResult.recordset[0].Id;

    // Insert version record
    await dmsPool
      .request()
      .input("docId", sql.Int, newDocId)
      .input("version", sql.Int, newVersion)
      .input("fileName", sql.NVarChar, req.file.originalname)
      .input("filePath", sql.NVarChar, filePath)
      .input("fileSize", sql.BigInt, req.file.size)
      .input("fileHash", sql.NVarChar, fileHash)
      .input("uploadedBy", sql.NVarChar, req.user.email)
      .input("reason", sql.NVarChar, reason || "New version upload")
      .query(`
        INSERT INTO DmsDocVersions (DocId, Version, FileName, FilePath, FileSize, FileHash, UploadedBy, Reason)
        VALUES (@docId, @version, @fileName, @filePath, @fileSize, @fileHash, @uploadedBy, @reason)
      `);

    // Copy access from old doc
    await dmsPool
      .request()
      .input("newDocId", sql.Int, newDocId)
      .input("oldDocId", sql.Int, docId)
      .query(`
        INSERT INTO DmsDocAccess (DocId, Email, AccessType)
        SELECT @newDocId, Email, AccessType FROM DmsDocAccess WHERE DocId = @oldDocId
      `);

    // Create RM approval for new version – FIX: use ManagerID
    if (emp.ManagerID) {
      const rmResult = await spotPool
        .request()
        .input("rmId", sql.NVarChar, emp.ManagerID)
        .query("SELECT TOP 1 EmpEmail, EmpName FROM dbo.EMP WHERE EmpID = @rmId");

      if (rmResult.recordset.length > 0) {
        const rmEmail = rmResult.recordset[0].EmpEmail;
        await dmsPool
          .request()
          .input("docId", sql.Int, newDocId)
          .input("version", sql.Int, newVersion)
          .input("stage", sql.NVarChar, "reporting_manager")
          .input("approverEmail", sql.NVarChar, rmEmail)
          .query(`
            INSERT INTO DmsApprovals (DocId, Version, Stage, ApproverEmail)
            VALUES (@docId, @version, @stage, @approverEmail)
          `);

        sendEmail(
          rmEmail,
          `[DMS] Approval Required: ${doc.Title} v${newVersion}`,
          `<h3>New Version Approval Required</h3>
           <p><strong>${emp.EmpName || req.user.email}</strong> has uploaded version ${newVersion} of a controlled document.</p>
           <p><strong>Title:</strong> ${doc.Title}</p>
           <p><strong>Reason:</strong> ${reason || "N/A"}</p>
           <p>Please log in to the DMS to review and approve/reject this document.</p>`
        );
      }
    }

    // Notify all users with access to old doc
    const accessUsers = await dmsPool
      .request()
      .input("oldDocId2", sql.Int, docId)
      .input("creatorEmail", sql.NVarChar, req.user.email)
      .query("SELECT Email FROM DmsDocAccess WHERE DocId = @oldDocId2 AND Email != @creatorEmail");

    const notifyEmails = accessUsers.recordset.map((r) => r.Email).filter(Boolean);
    if (notifyEmails.length > 0) {
      sendEmail(
        notifyEmails,
        `[DMS] New Version Uploaded: ${doc.Title} v${newVersion}`,
        `<h3>Document Updated</h3>
         <p>A new version (v${newVersion}) of <strong>${doc.Title}</strong> has been uploaded and is pending approval.</p>
         <p><strong>Uploaded by:</strong> ${emp.EmpName || req.user.email}</p>
         <p><strong>Reason:</strong> ${reason || "N/A"}</p>`
      );
    }

    await logAudit("document_new_version", "document", newDocId, req.user.email, reason || null, { version: doc.CurrentVersion }, { version: newVersion }, getIp(req));

    res.json({ success: true, newDocId, version: newVersion, fileHash });
  } catch (err) {
    console.error("[NewVersion] error:", err.message);
    res.status(500).json({ error: "New version upload failed" });
  }
});

// GET /api/documents/:id/verify
app.get("/api/documents/:id/verify", requireAuth, async (req, res) => {
  try {
    const docId = Number(req.params.id);
    const docResult = await dmsPool
      .request()
      .input("id", sql.Int, docId)
      .query("SELECT FileName, FilePath, FileHash FROM DmsDocuments WHERE Id = @id");

    if (docResult.recordset.length === 0) {
      return res.status(404).json({ error: "Document not found" });
    }

    const doc = docResult.recordset[0];

    if (!fs.existsSync(doc.FilePath)) {
      return res.json({ verified: false, reason: "File not found on disk" });
    }

    const currentHash = await computeFileHash(doc.FilePath);
    const verified = currentHash === doc.FileHash;

    await logAudit("document_verify", "document", docId, req.user.email, null, null, { verified, storedHash: doc.FileHash, currentHash }, getIp(req));

    res.json({
      verified,
      storedHash: doc.FileHash,
      currentHash,
      fileName: doc.FileName,
    });
  } catch (err) {
    console.error("[Verify] error:", err.message);
    res.status(500).json({ error: "Verification failed" });
  }
});

// ─── 7D. DOCUMENT VERIFICATION (upload-based) ──────────────────

// POST /api/documents/verify-upload
app.post("/api/documents/verify-upload", requireAuth, upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded" });
    }

    const filePath = req.file.path;
    const uploadHash = await computeFileHash(filePath);

    // Clean up the uploaded file after hashing
    fs.unlink(filePath, () => {});

    // Search for matching hash in DmsDocuments
    const docMatch = await dmsPool
      .request()
      .input("hash", sql.NVarChar, uploadHash)
      .query(`
        SELECT Id, Title, FileName, CurrentVersion, Status, ApprovalStatus, CreatorEmail, Department, Location, CreatedAt
        FROM DmsDocuments
        WHERE FileHash = @hash AND IsControlled = 1 AND Status != 'deleted'
      `);

    // Also check version history
    const versionMatch = await dmsPool
      .request()
      .input("hash", sql.NVarChar, uploadHash)
      .query(`
        SELECT dv.DocId, dv.Version, dv.FileName, dv.FileHash, d.Title, d.Status, d.ApprovalStatus
        FROM DmsDocVersions dv
        JOIN DmsDocuments d ON d.Id = dv.DocId
        WHERE dv.FileHash = @hash AND d.IsControlled = 1
      `);

    const matched = docMatch.recordset.length > 0 || versionMatch.recordset.length > 0;

    await logAudit("document_verify_upload", "document", null, req.user.email, null, null, { matched, uploadHash }, getIp(req));

    res.json({
      matched,
      uploadHash,
      documents: docMatch.recordset,
      versionMatches: versionMatch.recordset,
    });
  } catch (err) {
    console.error("[VerifyUpload] error:", err.message);
    res.status(500).json({ error: "Verification failed" });
  }
});

// ─── 7E. APPROVAL WORKFLOW ──────────────────────────────────────

// GET /api/approvals/pending
app.get("/api/approvals/pending", requireAuth, async (req, res) => {
  try {
    const result = await dmsPool
      .request()
      .input("email", sql.NVarChar, req.user.email)
      .query(`
        SELECT a.*, d.Title, d.FileName, d.CreatorEmail, d.Department, d.Location, d.CurrentVersion, d.HodSkipped
        FROM DmsApprovals a
        JOIN DmsDocuments d ON d.Id = a.DocId
        WHERE a.ApproverEmail = @email AND a.Status = 'pending' AND d.Status != 'deleted'
        ORDER BY a.CreatedAt DESC
      `);

    res.json(result.recordset);
  } catch (err) {
    console.error("[Approvals] pending error:", err.message);
    res.status(500).json({ error: "Failed to fetch pending approvals" });
  }
});

// POST /api/approvals/:approvalId/approve
app.post("/api/approvals/:approvalId/approve", requireAuth, async (req, res) => {
  try {
    const approvalId = Number(req.params.approvalId);
    const { comments } = req.body;

    const apResult = await dmsPool
      .request()
      .input("id", sql.Int, approvalId)
      .query("SELECT * FROM DmsApprovals WHERE Id = @id AND Status = 'pending'");

    if (apResult.recordset.length === 0) {
      return res.status(404).json({ error: "Pending approval not found" });
    }

    const approval = apResult.recordset[0];

    if (approval.ApproverEmail !== req.user.email) {
      return res.status(403).json({ error: "You are not the designated approver" });
    }

    await dmsPool
      .request()
      .input("id", sql.Int, approvalId)
      .input("comments", sql.NVarChar, comments || null)
      .query(`
        UPDATE DmsApprovals SET Status = 'approved', Comments = @comments, ActionAt = SYSUTCDATETIME()
        WHERE Id = @id
      `);

    const docId = approval.DocId;
    const version = approval.Version;

    const docResult = await dmsPool
      .request()
      .input("docId", sql.Int, docId)
      .query("SELECT * FROM DmsDocuments WHERE Id = @docId");
    const doc = docResult.recordset[0];

    let nextStage = null;
    let docFullyApproved = false;

    if (approval.Stage === "reporting_manager") {
      const hodResult = await dmsPool
        .request()
        .input("location", sql.NVarChar, doc.Location || "")
        .input("department", sql.NVarChar, doc.Department || "")
        .query("SELECT HodEmail, HodName FROM DmsHods WHERE Location = @location AND Department = @department AND Active = 1");

      if (hodResult.recordset.length > 0) {
        const hodEmail = hodResult.recordset[0].HodEmail;
        nextStage = "hod";

        await dmsPool
          .request()
          .input("docId", sql.Int, docId)
          .input("version", sql.Int, version)
          .input("stage", sql.NVarChar, "hod")
          .input("approverEmail", sql.NVarChar, hodEmail)
          .query(`
            INSERT INTO DmsApprovals (DocId, Version, Stage, ApproverEmail)
            VALUES (@docId, @version, @stage, @approverEmail)
          `);

        await dmsPool
          .request()
          .input("docId", sql.Int, docId)
          .query("UPDATE DmsDocuments SET ApprovalStatus = 'pending_hod', UpdatedAt = SYSUTCDATETIME() WHERE Id = @docId");

        sendEmail(
          hodEmail,
          `[DMS] HOD Approval Required: ${doc.Title} v${version}`,
          `<h3>HOD Approval Required</h3>
           <p>Document <strong>${doc.Title}</strong> (v${version}) has been approved by the Reporting Manager and now requires your approval as HOD.</p>
           <p><strong>Department:</strong> ${doc.Department}</p>
           <p><strong>Location:</strong> ${doc.Location}</p>
           <p>Please log in to the DMS to review.</p>`
        );
      } else {
        await dmsPool
          .request()
          .input("docId", sql.Int, docId)
          .query("UPDATE DmsDocuments SET HodSkipped = 1, UpdatedAt = SYSUTCDATETIME() WHERE Id = @docId");

        nextStage = "document_controller";
        await createDcApproval(docId, version, doc);
      }
    } else if (approval.Stage === "hod") {
      nextStage = "document_controller";
      await createDcApproval(docId, version, doc);
    } else if (approval.Stage === "document_controller") {
      docFullyApproved = true;

      const finalHash = await computeFileHash(doc.FilePath);

      await dmsPool
        .request()
        .input("docId", sql.Int, docId)
        .input("hash", sql.NVarChar, finalHash)
        .query(`
          UPDATE DmsDocuments
          SET Status = 'approved', ApprovalStatus = 'approved', FileHash = @hash, UpdatedAt = SYSUTCDATETIME()
          WHERE Id = @docId
        `);

      await sendApprovalCompletionEmail(doc, version, req.user.email);
    }

    await logAudit("approval_approve", "approval", approvalId, req.user.email, comments || null,
      { stage: approval.Stage, status: "pending" },
      { stage: approval.Stage, status: "approved", nextStage: nextStage || "complete" },
      getIp(req));

    res.json({ success: true, nextStage, fullyApproved: docFullyApproved });
  } catch (err) {
    console.error("[Approve] error:", err.message);
    res.status(500).json({ error: "Approval failed" });
  }
});

// POST /api/approvals/:approvalId/reject
app.post("/api/approvals/:approvalId/reject", requireAuth, async (req, res) => {
  try {
    const approvalId = Number(req.params.approvalId);
    const { comments } = req.body;

    const apResult = await dmsPool
      .request()
      .input("id", sql.Int, approvalId)
      .query("SELECT * FROM DmsApprovals WHERE Id = @id AND Status = 'pending'");

    if (apResult.recordset.length === 0) {
      return res.status(404).json({ error: "Pending approval not found" });
    }

    const approval = apResult.recordset[0];

    if (approval.ApproverEmail !== req.user.email) {
      return res.status(403).json({ error: "You are not the designated approver" });
    }

    await dmsPool
      .request()
      .input("id", sql.Int, approvalId)
      .input("comments", sql.NVarChar, comments || null)
      .query(`
        UPDATE DmsApprovals SET Status = 'rejected', Comments = @comments, ActionAt = SYSUTCDATETIME()
        WHERE Id = @id
      `);

    await dmsPool
      .request()
      .input("docId", sql.Int, approval.DocId)
      .query("UPDATE DmsDocuments SET Status = 'rejected', ApprovalStatus = 'rejected', UpdatedAt = SYSUTCDATETIME() WHERE Id = @docId");

    const docResult = await dmsPool.request().input("did", sql.Int, approval.DocId)
      .query("SELECT Title, CreatorEmail FROM DmsDocuments WHERE Id = @did");
    const doc = docResult.recordset[0];

    if (doc) {
      sendEmail(
        doc.CreatorEmail,
        `[DMS] Document Rejected: ${doc.Title}`,
        `<h3>Document Rejected</h3>
         <p>Your document <strong>${doc.Title}</strong> has been rejected at the <strong>${approval.Stage.replace(/_/g, " ")}</strong> stage.</p>
         <p><strong>Rejected by:</strong> ${req.user.email}</p>
         <p><strong>Comments:</strong> ${comments || "No comments provided"}</p>
         <p>Please review the feedback and upload a revised version if needed.</p>`
      );
    }

    await logAudit("approval_reject", "approval", approvalId, req.user.email, comments || null,
      { stage: approval.Stage, status: "pending" },
      { stage: approval.Stage, status: "rejected" },
      getIp(req));

    res.json({ success: true });
  } catch (err) {
    console.error("[Reject] error:", err.message);
    res.status(500).json({ error: "Rejection failed" });
  }
});

// Helper: create document controller approval
async function createDcApproval(docId, version, doc) {
  try {
    const admins = await dmsPool.request().query("SELECT Email FROM DmsAdmins WHERE Active = 1");
    if (admins.recordset.length > 0) {
      const dcEmail = admins.recordset[0].Email;

      await dmsPool
        .request()
        .input("docId", sql.Int, docId)
        .input("version", sql.Int, version)
        .input("stage", sql.NVarChar, "document_controller")
        .input("approverEmail", sql.NVarChar, dcEmail)
        .query(`
          INSERT INTO DmsApprovals (DocId, Version, Stage, ApproverEmail)
          VALUES (@docId, @version, @stage, @approverEmail)
        `);

      await dmsPool
        .request()
        .input("docId", sql.Int, docId)
        .query("UPDATE DmsDocuments SET ApprovalStatus = 'pending_dc', UpdatedAt = SYSUTCDATETIME() WHERE Id = @docId");

      sendEmail(
        dcEmail,
        `[DMS] DC Approval Required: ${doc.Title} v${version}`,
        `<h3>Document Controller Approval Required</h3>
         <p>Document <strong>${doc.Title}</strong> (v${version}) has passed all previous approvals and now requires Document Controller approval.</p>
         <p><strong>Department:</strong> ${doc.Department}</p>
         <p><strong>Location:</strong> ${doc.Location}</p>
         <p>Please log in to the DMS to finalize.</p>`
      );
    }
  } catch (err) {
    console.error("[CreateDcApproval] error:", err.message);
  }
}

// Helper: send approval completion email
async function sendApprovalCompletionEmail(doc, version, dcEmail) {
  try {
    const approvals = await dmsPool
      .request()
      .input("docId", sql.Int, doc.Id)
      .query("SELECT DISTINCT ApproverEmail, Stage FROM DmsApprovals WHERE DocId = @docId AND Status = 'approved'");

    const approverEmails = approvals.recordset.map((a) => a.ApproverEmail);

    const accessList = await dmsPool
      .request()
      .input("docId", sql.Int, doc.Id)
      .query("SELECT Email FROM DmsDocAccess WHERE DocId = @docId");
    const ccEmails = accessList.recordset.map((a) => a.Email).filter((e) => e !== doc.CreatorEmail && !approverEmails.includes(e));

    const toEmails = [doc.CreatorEmail, ...approverEmails].filter((v, i, a) => a.indexOf(v) === i);

    const html = `
      <h2>Document Approved</h2>
      <table style="border-collapse:collapse;width:100%;max-width:600px;">
        <tr><td style="padding:8px;border:1px solid #ddd;font-weight:bold;">Title</td><td style="padding:8px;border:1px solid #ddd;">${doc.Title}</td></tr>
        <tr><td style="padding:8px;border:1px solid #ddd;font-weight:bold;">Version</td><td style="padding:8px;border:1px solid #ddd;">${version}</td></tr>
        <tr><td style="padding:8px;border:1px solid #ddd;font-weight:bold;">Department</td><td style="padding:8px;border:1px solid #ddd;">${doc.Department || "N/A"}</td></tr>
        <tr><td style="padding:8px;border:1px solid #ddd;font-weight:bold;">Location</td><td style="padding:8px;border:1px solid #ddd;">${doc.Location || "N/A"}</td></tr>
        <tr><td style="padding:8px;border:1px solid #ddd;font-weight:bold;">Creator</td><td style="padding:8px;border:1px solid #ddd;">${doc.CreatorEmail}</td></tr>
        <tr><td style="padding:8px;border:1px solid #ddd;font-weight:bold;">Status</td><td style="padding:8px;border:1px solid #ddd;color:green;font-weight:bold;">APPROVED</td></tr>
        <tr><td style="padding:8px;border:1px solid #ddd;font-weight:bold;">File Hash</td><td style="padding:8px;border:1px solid #ddd;font-family:monospace;font-size:11px;">${doc.FileHash}</td></tr>
      </table>
      <p style="margin-top:16px;">This document has completed the full approval workflow and is now an official controlled copy.</p>
    `;

    if (ccEmails.length > 0) {
      await sendEmailWithCc(toEmails, ccEmails, `[DMS] Document Approved: ${doc.Title} v${version}`, html);
    } else {
      await sendEmail(toEmails, `[DMS] Document Approved: ${doc.Title} v${version}`, html);
    }
  } catch (err) {
    console.error("[ApprovalEmail] error:", err.message);
  }
}

// Helper: check document access – FIX: use correct EMP column
async function checkDocAccess(docId, email) {
  try {
    const isAdmin = (await dmsPool.request().input("ae", sql.NVarChar, email)
      .query("SELECT Email FROM DmsAdmins WHERE Email = @ae AND Active = 1")).recordset.length > 0;
    if (isAdmin) return true;

    const docResult = await dmsPool.request().input("id", sql.Int, docId)
      .query("SELECT CreatorEmail, ShareScope, ShareGroupId, Department FROM DmsDocuments WHERE Id = @id AND Status != 'deleted'");
    if (docResult.recordset.length === 0) return false;
    const doc = docResult.recordset[0];

    if (doc.CreatorEmail === email) return true;

    const hasAccess = (await dmsPool.request().input("did", sql.Int, docId).input("ue", sql.NVarChar, email)
      .query("SELECT 1 FROM DmsDocAccess WHERE DocId = @did AND Email = @ue")).recordset.length > 0;
    if (hasAccess) return true;

    if (doc.ShareScope === "company") return true;

    if (doc.ShareScope === "department") {
      const userDept = (await spotPool.request().input("ue2", sql.NVarChar, email)
        .query("SELECT TOP 1 Dept FROM dbo.EMP WHERE EmpEmail = @ue2")).recordset[0]?.Dept;
      if (userDept === doc.Department) return true;
    }

    if (doc.ShareScope === "group" && doc.ShareGroupId) {
      const inGroup = (await dmsPool.request().input("gid", sql.Int, doc.ShareGroupId).input("ue3", sql.NVarChar, email)
        .query("SELECT 1 FROM DmsShareGroupMembers WHERE GroupId = @gid AND Email = @ue3")).recordset.length > 0;
      if (inGroup) return true;
    }

    const isApprover = (await dmsPool.request().input("did2", sql.Int, docId).input("ue4", sql.NVarChar, email)
      .query("SELECT 1 FROM DmsApprovals WHERE DocId = @did2 AND ApproverEmail = @ue4")).recordset.length > 0;
    if (isApprover) return true;

    return false;
  } catch (err) {
    console.error("[CheckAccess] error:", err.message);
    return false;
  }
}

// ─── 7F. HOD MANAGEMENT (admin only) ───────────────────────────

// GET /api/hods
app.get("/api/hods", requireAdmin, async (req, res) => {
  try {
    const result = await dmsPool.request().query("SELECT * FROM DmsHods ORDER BY Location, Department");
    res.json(result.recordset);
  } catch (err) {
    console.error("[HODs] list error:", err.message);
    res.status(500).json({ error: "Failed to fetch HODs" });
  }
});

// POST /api/hods
app.post("/api/hods", requireAdmin, async (req, res) => {
  try {
    const { location, department, hodEmail, hodName } = req.body;
    if (!location || !department || !hodEmail) {
      return res.status(400).json({ error: "location, department, and hodEmail are required" });
    }

    const existing = await dmsPool
      .request()
      .input("location", sql.NVarChar, location)
      .input("department", sql.NVarChar, department)
      .query("SELECT Id FROM DmsHods WHERE Location = @location AND Department = @department");

    if (existing.recordset.length > 0) {
      await dmsPool
        .request()
        .input("id", sql.Int, existing.recordset[0].Id)
        .input("hodEmail", sql.NVarChar, hodEmail)
        .input("hodName", sql.NVarChar, hodName || null)
        .input("updatedBy", sql.NVarChar, req.user.email)
        .query(`
          UPDATE DmsHods SET HodEmail = @hodEmail, HodName = @hodName, Active = 1, UpdatedBy = @updatedBy, UpdatedAt = SYSUTCDATETIME()
          WHERE Id = @id
        `);

      await logAudit("hod_update", "hod", existing.recordset[0].Id, req.user.email, null, null, { location, department, hodEmail }, getIp(req));
      res.json({ success: true, id: existing.recordset[0].Id, action: "updated" });
    } else {
      const insertResult = await dmsPool
        .request()
        .input("location", sql.NVarChar, location)
        .input("department", sql.NVarChar, department)
        .input("hodEmail", sql.NVarChar, hodEmail)
        .input("hodName", sql.NVarChar, hodName || null)
        .input("updatedBy", sql.NVarChar, req.user.email)
        .query(`
          INSERT INTO DmsHods (Location, Department, HodEmail, HodName, UpdatedBy)
          OUTPUT INSERTED.Id
          VALUES (@location, @department, @hodEmail, @hodName, @updatedBy)
        `);

      await logAudit("hod_create", "hod", insertResult.recordset[0].Id, req.user.email, null, null, { location, department, hodEmail }, getIp(req));
      res.json({ success: true, id: insertResult.recordset[0].Id, action: "created" });
    }
  } catch (err) {
    console.error("[HODs] create error:", err.message);
    res.status(500).json({ error: "Failed to create/update HOD" });
  }
});

// PUT /api/hods/:id
app.put("/api/hods/:id", requireAdmin, async (req, res) => {
  try {
    const id = Number(req.params.id);
    const { location, department, hodEmail, hodName } = req.body;

    const existing = await dmsPool.request().input("id", sql.Int, id).query("SELECT * FROM DmsHods WHERE Id = @id");
    if (existing.recordset.length === 0) {
      return res.status(404).json({ error: "HOD mapping not found" });
    }

    const before = existing.recordset[0];

    await dmsPool
      .request()
      .input("id", sql.Int, id)
      .input("location", sql.NVarChar, location || before.Location)
      .input("department", sql.NVarChar, department || before.Department)
      .input("hodEmail", sql.NVarChar, hodEmail || before.HodEmail)
      .input("hodName", sql.NVarChar, hodName !== undefined ? hodName : before.HodName)
      .input("updatedBy", sql.NVarChar, req.user.email)
      .query(`
        UPDATE DmsHods
        SET Location = @location, Department = @department, HodEmail = @hodEmail, HodName = @hodName,
            UpdatedBy = @updatedBy, UpdatedAt = SYSUTCDATETIME()
        WHERE Id = @id
      `);

    await logAudit("hod_update", "hod", id, req.user.email, null, before, { location, department, hodEmail, hodName }, getIp(req));
    res.json({ success: true });
  } catch (err) {
    console.error("[HODs] update error:", err.message);
    res.status(500).json({ error: "Failed to update HOD" });
  }
});

// DELETE /api/hods/:id
app.delete("/api/hods/:id", requireAdmin, async (req, res) => {
  try {
    const id = Number(req.params.id);

    const existing = await dmsPool.request().input("id", sql.Int, id).query("SELECT * FROM DmsHods WHERE Id = @id");
    if (existing.recordset.length === 0) {
      return res.status(404).json({ error: "HOD mapping not found" });
    }

    await dmsPool
      .request()
      .input("id", sql.Int, id)
      .query("UPDATE DmsHods SET Active = 0, UpdatedAt = SYSUTCDATETIME() WHERE Id = @id");

    await logAudit("hod_deactivate", "hod", id, req.user.email, null, existing.recordset[0], { active: false }, getIp(req));
    res.json({ success: true });
  } catch (err) {
    console.error("[HODs] delete error:", err.message);
    res.status(500).json({ error: "Failed to deactivate HOD" });
  }
});

// GET /api/hods/locations-departments – FIX: use correct EMP columns
app.get("/api/hods/locations-departments", requireAdmin, async (req, res) => {
  try {
    const locations = await spotPool.request().query("SELECT DISTINCT EmpLocation AS Location FROM dbo.EMP WHERE EmpLocation IS NOT NULL AND EmpLocation != '' AND ActiveFlag = 1 ORDER BY EmpLocation");
    const departments = await spotPool.request().query("SELECT DISTINCT Dept AS Department FROM dbo.EMP WHERE Dept IS NOT NULL AND Dept != '' AND ActiveFlag = 1 ORDER BY Dept");

    res.json({
      locations: locations.recordset.map((r) => r.Location),
      departments: departments.recordset.map((r) => r.Department),
    });
  } catch (err) {
    console.error("[HODs] locations-departments error:", err.message);
    res.status(500).json({ error: "Failed to fetch locations/departments" });
  }
});

// GET /api/hods/combinations?location=...&department=...
app.get("/api/hods/combinations", requireAdmin, async (req, res) => {
  try {
    const location = String(req.query.location || "").trim();
    const department = String(req.query.department || "").trim();
    const rq = spotPool.request();
    let where = "WHERE ActiveFlag = 1 AND ISNULL(EmpLocation,'') != '' AND ISNULL(Dept,'') != ''";
    if (location) {
      rq.input("location", sql.NVarChar, `%${location}%`);
      where += " AND EmpLocation LIKE @location";
    }
    if (department) {
      rq.input("department", sql.NVarChar, `%${department}%`);
      where += " AND Dept LIKE @department";
    }

    const result = await rq.query(`
      SELECT DISTINCT EmpLocation AS Location, Dept AS Department
      FROM dbo.EMP
      ${where}
      ORDER BY EmpLocation, Dept
    `);

    res.json({ combinations: result.recordset });
  } catch (err) {
    console.error("[HODs] combinations error:", err.message);
    res.status(500).json({ error: "Failed to fetch location-department combinations" });
  }
});

// ─── 7G. SHARE GROUPS ──────────────────────────────────────────

// GET /api/share-preview?scope=private|group|department|company&groupId=#
app.get("/api/share-preview", requireAuth, async (req, res) => {
  try {
    const scope = String(req.query.scope || "private");
    const groupId = Number(req.query.groupId || 0);
    const users = await getScopeUsers(scope, groupId, req.user.email);
    const dedup = Array.from(
      new Map(users.map((u) => [String(u.EmpEmail || "").toLowerCase(), u])).values()
    );
    res.json({
      scope,
      count: dedup.length,
      users: dedup.map((u) => ({
        email: u.EmpEmail,
        name: u.EmpName || u.EmpEmail,
        empId: u.EmpID || null,
        department: u.Department || "",
        location: u.Location || "",
      })),
    });
  } catch (err) {
    console.error("[SharePreview] error:", err.message);
    res.status(500).json({ error: "Failed to generate share preview" });
  }
});

// GET /api/share-groups
app.get("/api/share-groups", requireAuth, async (req, res) => {
  try {
    const groups = await dmsPool
      .request()
      .input("email", sql.NVarChar, req.user.email)
      .query(`
        SELECT sg.*,
          (SELECT COUNT(*) FROM DmsShareGroupMembers sgm WHERE sgm.GroupId = sg.Id) AS MemberCount
        FROM DmsShareGroups sg
        WHERE sg.CreatorEmail = @email
        ORDER BY sg.CreatedAt DESC
      `);

    const result = [];
    for (const group of groups.recordset) {
      const members = await dmsPool
        .request()
        .input("groupId", sql.Int, group.Id)
        .query("SELECT Email FROM DmsShareGroupMembers WHERE GroupId = @groupId");
      result.push({
        ...group,
        members: members.recordset.map((m) => m.Email),
      });
    }

    res.json(result);
  } catch (err) {
    console.error("[ShareGroups] list error:", err.message);
    res.status(500).json({ error: "Failed to fetch share groups" });
  }
});

// POST /api/share-groups
app.post("/api/share-groups", requireAuth, async (req, res) => {
  try {
    const name = String(req.body?.name || "").trim();
    const members = Array.isArray(req.body?.members) ? req.body.members : [];
    if (!name || name.length < 2) {
      return res.status(400).json({ error: "Group name is required" });
    }
    const normalizedMembers = Array.from(
      new Set(
        members
          .map((m) => String(m || "").trim().toLowerCase())
          .filter(Boolean)
      )
    );
    if (!normalizedMembers.length) {
      return res.status(400).json({ error: "At least one member is required" });
    }

    const insertResult = await dmsPool
      .request()
      .input("name", sql.NVarChar, name)
      .input("creatorEmail", sql.NVarChar, req.user.email)
      .query(`
        INSERT INTO DmsShareGroups (Name, CreatorEmail)
        OUTPUT INSERTED.Id
        VALUES (@name, @creatorEmail)
      `);

    const groupId = insertResult.recordset[0].Id;

    for (const email of normalizedMembers) {
      await dmsPool
        .request()
        .input("groupId", sql.Int, groupId)
        .input("email", sql.NVarChar, email)
        .query(`
          MERGE DmsShareGroupMembers AS t
          USING (SELECT @groupId AS GroupId, @email AS Email) s
          ON t.GroupId=s.GroupId AND t.Email=s.Email
          WHEN NOT MATCHED THEN INSERT (GroupId, Email) VALUES (@groupId, @email);
        `);
    }

    await logAudit("share_group_create", "share_group", groupId, req.user.email, null, null, { name, members: normalizedMembers }, getIp(req));
    res.json({ success: true, id: groupId, name, members: normalizedMembers });
  } catch (err) {
    console.error("[ShareGroups] create error:", err.message);
    res.status(500).json({ error: "Failed to create share group" });
  }
});

// PUT /api/share-groups/:id
app.put("/api/share-groups/:id", requireAuth, async (req, res) => {
  try {
    const groupId = Number(req.params.id);
    const { name, members } = req.body;

    const group = await dmsPool.request().input("id", sql.Int, groupId).input("email", sql.NVarChar, req.user.email)
      .query("SELECT * FROM DmsShareGroups WHERE Id = @id AND CreatorEmail = @email");
    if (group.recordset.length === 0) {
      return res.status(404).json({ error: "Share group not found or not owned by you" });
    }

    if (name) {
      await dmsPool.request().input("id", sql.Int, groupId).input("name", sql.NVarChar, name)
        .query("UPDATE DmsShareGroups SET Name = @name WHERE Id = @id");
    }

    if (members && Array.isArray(members)) {
      await dmsPool.request().input("groupId", sql.Int, groupId)
        .query("DELETE FROM DmsShareGroupMembers WHERE GroupId = @groupId");

      for (const email of members) {
        if (email && email.trim()) {
          await dmsPool
            .request()
            .input("groupId", sql.Int, groupId)
            .input("email", sql.NVarChar, email.trim())
            .query("INSERT INTO DmsShareGroupMembers (GroupId, Email) VALUES (@groupId, @email)");
        }
      }
    }

    await logAudit("share_group_update", "share_group", groupId, req.user.email, null, null, { name, members }, getIp(req));
    res.json({ success: true });
  } catch (err) {
    console.error("[ShareGroups] update error:", err.message);
    res.status(500).json({ error: "Failed to update share group" });
  }
});

// DELETE /api/share-groups/:id
app.delete("/api/share-groups/:id", requireAuth, async (req, res) => {
  try {
    const groupId = Number(req.params.id);

    const group = await dmsPool.request().input("id", sql.Int, groupId).input("email", sql.NVarChar, req.user.email)
      .query("SELECT * FROM DmsShareGroups WHERE Id = @id AND CreatorEmail = @email");
    if (group.recordset.length === 0) {
      return res.status(404).json({ error: "Share group not found or not owned by you" });
    }

    await dmsPool.request().input("groupId", sql.Int, groupId)
      .query("DELETE FROM DmsShareGroupMembers WHERE GroupId = @groupId");

    await dmsPool.request().input("id", sql.Int, groupId)
      .query("DELETE FROM DmsShareGroups WHERE Id = @id");

    await logAudit("share_group_delete", "share_group", groupId, req.user.email, null, group.recordset[0], null, getIp(req));
    res.json({ success: true });
  } catch (err) {
    console.error("[ShareGroups] delete error:", err.message);
    res.status(500).json({ error: "Failed to delete share group" });
  }
});

// ─── 7H. PUBLIC LINKS ──────────────────────────────────────────

// POST /api/public-links
app.post("/api/public-links", requireAuth, async (req, res) => {
  try {
    const { docId, expiresAt } = req.body;
    if (!docId) {
      return res.status(400).json({ error: "docId is required" });
    }

    const canAccess = await checkDocAccess(Number(docId), req.user.email);
    if (!canAccess) {
      return res.status(403).json({ error: "Access denied" });
    }

    const token = uuidv4() + "-" + crypto.randomBytes(16).toString("hex");

    const result = await dmsPool
      .request()
      .input("docId", sql.Int, Number(docId))
      .input("token", sql.NVarChar, token)
      .input("createdBy", sql.NVarChar, req.user.email)
      .input("expiresAt", sql.NVarChar, expiresAt || null)
      .query(`
        INSERT INTO DmsPublicLinks (DocId, LinkToken, CreatedBy, ExpiresAt)
        OUTPUT INSERTED.Id
        VALUES (@docId, @token, @createdBy, ${expiresAt ? "@expiresAt" : "NULL"})
      `);

    await logAudit("public_link_create", "public_link", result.recordset[0].Id, req.user.email, null, null, { docId, token }, getIp(req));

    res.json({ success: true, id: result.recordset[0].Id, token });
  } catch (err) {
    console.error("[PublicLinks] create error:", err.message);
    res.status(500).json({ error: "Failed to create public link" });
  }
});

// GET /api/documents/:id/public-links
app.get("/api/documents/:id/public-links", requireAuth, async (req, res) => {
  try {
    const docId = Number(req.params.id);
    if (!docId) return res.status(400).json({ error: "invalid_doc_id" });

    const canAccess = await checkDocAccess(docId, req.user.email);
    if (!canAccess) return res.status(403).json({ error: "Access denied" });

    const links = await dmsPool
      .request()
      .input("docId", sql.Int, docId)
      .query("SELECT * FROM DmsPublicLinks WHERE DocId = @docId ORDER BY CreatedAt DESC");

    res.json({ links: links.recordset });
  } catch (err) {
    console.error("[PublicLinks] list error:", err.message);
    res.status(500).json({ error: "Failed to fetch public links" });
  }
});

// GET /api/public/meta/:token
app.get("/api/public/meta/:token", async (req, res) => {
  try {
    const token = req.params.token;
    const linkResult = await dmsPool
      .request()
      .input("token", sql.NVarChar, token)
      .query(`
        SELECT pl.Id, pl.DocId, pl.ExpiresAt, pl.Active, pl.CreatedAt,
               d.Title, d.FileName, d.MimeType, d.CurrentVersion, d.IsControlled, d.Status
        FROM DmsPublicLinks pl
        JOIN DmsDocuments d ON d.Id = pl.DocId
        WHERE pl.LinkToken = @token
      `);

    if (linkResult.recordset.length === 0) {
      return res.status(404).json({ error: "Link not found" });
    }
    const link = linkResult.recordset[0];
    if (!link.Active) return res.status(410).json({ error: "Link is inactive" });
    if (link.ExpiresAt && new Date(link.ExpiresAt) < new Date()) {
      return res.status(410).json({ error: "Link has expired" });
    }

    res.json(link);
  } catch (err) {
    console.error("[PublicMeta] error:", err.message);
    res.status(500).json({ error: "Failed to fetch public link metadata" });
  }
});

// GET /api/public/view/:token  (NO AUTH REQUIRED)
app.get("/api/public/view/:token", async (req, res) => {
  try {
    const token = req.params.token;

    const linkResult = await dmsPool
      .request()
      .input("token", sql.NVarChar, token)
      .query(`
        SELECT pl.*, d.FileName, d.FilePath, d.MimeType, d.Title
        FROM DmsPublicLinks pl
        JOIN DmsDocuments d ON d.Id = pl.DocId
        WHERE pl.LinkToken = @token AND pl.Active = 1 AND d.Status != 'deleted'
      `);

    if (linkResult.recordset.length === 0) {
      return res.status(404).json({ error: "Link not found or expired" });
    }

    const link = linkResult.recordset[0];

    if (link.ExpiresAt && new Date(link.ExpiresAt) < new Date()) {
      return res.status(410).json({ error: "Link has expired" });
    }

    if (!fs.existsSync(link.FilePath)) {
      return res.status(404).json({ error: "File not found" });
    }

    await logAudit("public_link_view", "public_link", link.Id, null, null, null, { token }, getIp(req));

    res.setHeader("Content-Disposition", `inline; filename="${encodeURIComponent(link.FileName)}"`);
    res.setHeader("Content-Type", link.MimeType || "application/octet-stream");
    res.setHeader("X-Content-Type-Options", "nosniff");
    res.setHeader("Content-Security-Policy", "sandbox");
    fs.createReadStream(link.FilePath).pipe(res);
  } catch (err) {
    console.error("[PublicView] error:", err.message);
    res.status(500).json({ error: "Failed to serve public document" });
  }
});

// DELETE /api/public-links/:id
app.delete("/api/public-links/:id", requireAuth, async (req, res) => {
  try {
    const id = Number(req.params.id);

    const existing = await dmsPool.request().input("id", sql.Int, id)
      .query("SELECT * FROM DmsPublicLinks WHERE Id = @id");
    if (existing.recordset.length === 0) {
      return res.status(404).json({ error: "Public link not found" });
    }

    const isAdmin = (await dmsPool.request().input("ae", sql.NVarChar, req.user.email)
      .query("SELECT Email FROM DmsAdmins WHERE Email = @ae AND Active = 1")).recordset.length > 0;

    if (!isAdmin && existing.recordset[0].CreatedBy !== req.user.email) {
      return res.status(403).json({ error: "Access denied" });
    }

    await dmsPool.request().input("id", sql.Int, id)
      .query("UPDATE DmsPublicLinks SET Active = 0 WHERE Id = @id");

    await logAudit("public_link_revoke", "public_link", id, req.user.email, null, existing.recordset[0], { active: false }, getIp(req));
    res.json({ success: true });
  } catch (err) {
    console.error("[PublicLinks] delete error:", err.message);
    res.status(500).json({ error: "Failed to revoke public link" });
  }
});

// ─── 7I. USER MANAGEMENT (admin only) ──────────────────────────

// GET /api/admin/users – FIX: use correct EMP columns with aliases
app.get("/api/admin/users", requireAdmin, async (req, res) => {
  try {
    const { search, page = 1, pageSize = 50 } = req.query;
    const pageNum = Math.max(1, Number(page));
    const pageSz = Math.min(200, Math.max(1, Number(pageSize)));
    const offset = (pageNum - 1) * pageSz;

    let whereClause = "ActiveFlag = 1";
    const request = spotPool.request();

    if (search) {
      request.input("search", sql.NVarChar, `%${search}%`);
      whereClause += " AND (EmpName LIKE @search OR EmpEmail LIKE @search OR EmpID LIKE @search OR Dept LIKE @search)";
    }

    const countResult = await spotPool.request()
      .input("search", sql.NVarChar, search ? `%${search}%` : null)
      .query(`SELECT COUNT(*) as total FROM dbo.EMP WHERE ${whereClause}`);

    request.input("offset", sql.Int, offset);
    request.input("pageSz", sql.Int, pageSz);

    const result = await request.query(`
      SELECT EmpID, EmpName, EmpEmail, Dept AS Department, EmpLocation AS Location, ManagerID AS ReportingManagerID
      FROM dbo.EMP
      WHERE ${whereClause}
      ORDER BY EmpName
      OFFSET @offset ROWS FETCH NEXT @pageSz ROWS ONLY
    `);

    res.json({
      users: result.recordset,
      total: countResult.recordset[0].total,
      page: pageNum,
      pageSize: pageSz,
    });
  } catch (err) {
    console.error("[AdminUsers] list error:", err.message);
    res.status(500).json({ error: "Failed to fetch users" });
  }
});

// PATCH /api/admin/users/:email – FIX: map front-end field names to actual EMP columns
app.patch("/api/admin/users/:email", requireAdmin, async (req, res) => {
  try {
    const email = req.params.email;
    const updates = req.body;

    // Map front-end names to actual EMP column names
    const fieldMap = { Department: "Dept", Location: "EmpLocation" };
    const allowedFields = ["Department", "Location"];
    const setClauses = [];
    const request = spotPool.request();
    request.input("email", sql.NVarChar, email);

    for (const field of allowedFields) {
      if (updates[field] !== undefined) {
        const realCol = fieldMap[field];
        request.input(realCol, sql.NVarChar, updates[field]);
        setClauses.push(`${realCol} = @${realCol}`);
      }
    }

    if (setClauses.length === 0) {
      return res.status(400).json({ error: "No valid fields to update" });
    }

    await request.query(`UPDATE dbo.EMP SET ${setClauses.join(", ")} WHERE EmpEmail = @email`);

    await logAudit("admin_user_update", "user", null, req.user.email, null, { email }, updates, getIp(req));
    res.json({ success: true });
  } catch (err) {
    console.error("[AdminUsers] update error:", err.message);
    res.status(500).json({ error: "Failed to update user" });
  }
});

// POST /api/admin/users
app.post("/api/admin/users", requireAdmin, async (req, res) => {
  try {
    const payload = req.body || {};
    const EmpEmail = String(payload.EmpEmail || payload.email || "").trim();
    const EmpID = String(payload.EmpID || payload.empId || "").trim();
    const EmpName = String(payload.EmpName || payload.empName || "").trim();
    const Dept = String(payload.Dept || payload.Department || payload.department || "").trim();
    const EmpLocation = String(payload.EmpLocation || payload.Location || payload.location || "").trim();
    const ManagerID = String(payload.ManagerID || payload.ReportingManagerID || payload.reportingManagerId || "").trim();
    const ActiveFlag = payload.ActiveFlag === undefined ? 1 : payload.ActiveFlag ? 1 : 0;

    if (!EmpEmail) return res.status(400).json({ error: "EmpEmail is required" });

    const exists = await spotPool.request().input("email", sql.NVarChar, EmpEmail)
      .query("SELECT TOP 1 EmpEmail FROM dbo.EMP WHERE EmpEmail = @email");
    if (exists.recordset.length) {
      return res.status(409).json({ error: "User already exists" });
    }

    const values = {
      EmpEmail,
      ActiveFlag,
      ...(EmpID ? { EmpID } : {}),
      ...(EmpName ? { EmpName } : {}),
      ...(Dept ? { Dept } : {}),
      ...(EmpLocation ? { EmpLocation } : {}),
      ...(ManagerID ? { ManagerID } : {}),
    };

    const cols = Object.keys(values);
    const params = cols.map((k) => `@${k}`);
    const rq = spotPool.request();
    for (const [k, v] of Object.entries(values)) {
      if (k === "ActiveFlag") rq.input(k, sql.Bit, Number(v));
      else rq.input(k, sql.NVarChar(sql.MAX), String(v));
    }

    await rq.query(`INSERT INTO dbo.EMP (${cols.join(",")}) VALUES (${params.join(",")})`);
    await logAudit("admin_user_create", "user", null, req.user.email, null, null, values, getIp(req));
    res.json({ success: true });
  } catch (err) {
    console.error("[AdminUsers] create error:", err.message);
    res.status(500).json({ error: "Failed to create user", detail: String(err.message || err) });
  }
});

// DELETE /api/admin/users/:email (soft delete)
app.delete("/api/admin/users/:email", requireAdmin, async (req, res) => {
  try {
    const email = String(req.params.email || "").trim();
    if (!email) return res.status(400).json({ error: "email_required" });

    const existing = await spotPool.request().input("email", sql.NVarChar, email)
      .query("SELECT TOP 1 * FROM dbo.EMP WHERE EmpEmail = @email");
    if (!existing.recordset.length) return res.status(404).json({ error: "User not found" });

    await spotPool.request().input("email", sql.NVarChar, email)
      .query("UPDATE dbo.EMP SET ActiveFlag = 0 WHERE EmpEmail = @email");

    await logAudit(
      "admin_user_delete",
      "user",
      null,
      req.user.email,
      null,
      existing.recordset[0],
      { EmpEmail: email, ActiveFlag: 0 },
      getIp(req)
    );
    res.json({ success: true });
  } catch (err) {
    console.error("[AdminUsers] delete error:", err.message);
    res.status(500).json({ error: "Failed to delete user" });
  }
});

// GET /api/admin/admins
app.get("/api/admin/admins", requireAdmin, async (req, res) => {
  try {
    const result = await dmsPool.request().query("SELECT * FROM DmsAdmins ORDER BY CreatedAt");
    res.json(result.recordset);
  } catch (err) {
    console.error("[AdminAdmins] list error:", err.message);
    res.status(500).json({ error: "Failed to fetch admins" });
  }
});

// POST /api/admin/admins
app.post("/api/admin/admins", requireAdmin, async (req, res) => {
  try {
    const { email } = req.body;
    if (!email) {
      return res.status(400).json({ error: "Email is required" });
    }

    const existing = await dmsPool.request().input("email", sql.NVarChar, email)
      .query("SELECT Email, Active FROM DmsAdmins WHERE Email = @email");

    if (existing.recordset.length > 0) {
      if (existing.recordset[0].Active) {
        return res.status(409).json({ error: "Admin already exists" });
      }
      await dmsPool.request().input("email", sql.NVarChar, email).input("updatedBy", sql.NVarChar, req.user.email)
        .query("UPDATE DmsAdmins SET Active = 1, UpdatedBy = @updatedBy WHERE Email = @email");
    } else {
      await dmsPool.request()
        .input("email", sql.NVarChar, email)
        .input("updatedBy", sql.NVarChar, req.user.email)
        .query("INSERT INTO DmsAdmins (Email, UpdatedBy) VALUES (@email, @updatedBy)");
    }

    await logAudit("admin_add", "admin", null, req.user.email, null, null, { email }, getIp(req));
    res.json({ success: true });
  } catch (err) {
    console.error("[AdminAdmins] add error:", err.message);
    res.status(500).json({ error: "Failed to add admin" });
  }
});

// DELETE /api/admin/admins/:email
app.delete("/api/admin/admins/:email", requireAdmin, async (req, res) => {
  try {
    const email = decodeURIComponent(req.params.email);

    const existing = await dmsPool.request().input("email", sql.NVarChar, email)
      .query("SELECT Email FROM DmsAdmins WHERE Email = @email AND Active = 1");
    if (existing.recordset.length === 0) {
      return res.status(404).json({ error: "Admin not found" });
    }

    await dmsPool.request().input("email", sql.NVarChar, email).input("updatedBy", sql.NVarChar, req.user.email)
      .query("UPDATE DmsAdmins SET Active = 0, UpdatedBy = @updatedBy WHERE Email = @email");

    await logAudit("admin_remove", "admin", null, req.user.email, null, { email }, { active: false }, getIp(req));
    res.json({ success: true });
  } catch (err) {
    console.error("[AdminAdmins] remove error:", err.message);
    res.status(500).json({ error: "Failed to remove admin" });
  }
});

// ─── 7J. ANALYTICS (admin only) ────────────────────────────────

// GET /api/analytics/overview
app.get("/api/analytics/overview", requireAdmin, async (req, res) => {
  try {
    const totalDocs = (await dmsPool.request().query("SELECT COUNT(*) as c FROM DmsDocuments WHERE Status != 'deleted'")).recordset[0].c;
    const controlledDocs = (await dmsPool.request().query("SELECT COUNT(*) as c FROM DmsDocuments WHERE IsControlled = 1 AND Status != 'deleted'")).recordset[0].c;
    const pendingApprovals = (await dmsPool.request().query("SELECT COUNT(*) as c FROM DmsApprovals WHERE Status = 'pending'")).recordset[0].c;
    const approvedDocs = (await dmsPool.request().query("SELECT COUNT(*) as c FROM DmsDocuments WHERE Status = 'approved'")).recordset[0].c;
    const totalUsers = (await spotPool.request().query("SELECT COUNT(DISTINCT EmpEmail) as c FROM dbo.EMP WHERE ActiveFlag = 1")).recordset[0].c;
    const uniqueUploaders = (await dmsPool.request().query("SELECT COUNT(DISTINCT CreatorEmail) as c FROM DmsDocuments WHERE Status != 'deleted'")).recordset[0].c;

    res.json({
      totalDocuments: totalDocs,
      controlledDocuments: controlledDocs,
      pendingApprovals,
      approvedDocuments: approvedDocs,
      totalUsers,
      uniqueUploaders,
    });
  } catch (err) {
    console.error("[Analytics] overview error:", err.message);
    res.status(500).json({ error: "Failed to fetch analytics" });
  }
});

// GET /api/analytics/by-department
app.get("/api/analytics/by-department", requireAdmin, async (req, res) => {
  try {
    const result = await dmsPool.request().query(`
      SELECT Department, COUNT(*) as DocumentCount
      FROM DmsDocuments
      WHERE Status != 'deleted' AND Department IS NOT NULL
      GROUP BY Department
      ORDER BY DocumentCount DESC
    `);
    res.json(result.recordset);
  } catch (err) {
    console.error("[Analytics] by-department error:", err.message);
    res.status(500).json({ error: "Failed to fetch department analytics" });
  }
});

// GET /api/analytics/by-location
app.get("/api/analytics/by-location", requireAdmin, async (req, res) => {
  try {
    const result = await dmsPool.request().query(`
      SELECT Location, COUNT(*) as DocumentCount
      FROM DmsDocuments
      WHERE Status != 'deleted' AND Location IS NOT NULL
      GROUP BY Location
      ORDER BY DocumentCount DESC
    `);
    res.json(result.recordset);
  } catch (err) {
    console.error("[Analytics] by-location error:", err.message);
    res.status(500).json({ error: "Failed to fetch location analytics" });
  }
});

// GET /api/analytics/by-user
app.get("/api/analytics/by-user", requireAdmin, async (req, res) => {
  try {
    const result = await dmsPool.request().query(`
      SELECT TOP 20 CreatorEmail, COUNT(*) as DocumentCount
      FROM DmsDocuments
      WHERE Status != 'deleted'
      GROUP BY CreatorEmail
      ORDER BY DocumentCount DESC
    `);
    res.json(result.recordset);
  } catch (err) {
    console.error("[Analytics] by-user error:", err.message);
    res.status(500).json({ error: "Failed to fetch user analytics" });
  }
});

// GET /api/analytics/timeline
app.get("/api/analytics/timeline", requireAdmin, async (req, res) => {
  try {
    const result = await dmsPool.request().query(`
      SELECT
        FORMAT(CreatedAt, 'yyyy-MM') AS Month,
        COUNT(*) AS DocumentCount
      FROM DmsDocuments
      WHERE Status != 'deleted'
      GROUP BY FORMAT(CreatedAt, 'yyyy-MM')
      ORDER BY Month DESC
    `);
    res.json(result.recordset);
  } catch (err) {
    console.error("[Analytics] timeline error:", err.message);
    res.status(500).json({ error: "Failed to fetch timeline analytics" });
  }
});

// GET /api/analytics/approval-stats
app.get("/api/analytics/approval-stats", requireAdmin, async (req, res) => {
  try {
    const totalApprovals = (await dmsPool.request().query("SELECT COUNT(*) as c FROM DmsApprovals")).recordset[0].c;
    const approved = (await dmsPool.request().query("SELECT COUNT(*) as c FROM DmsApprovals WHERE Status = 'approved'")).recordset[0].c;
    const rejected = (await dmsPool.request().query("SELECT COUNT(*) as c FROM DmsApprovals WHERE Status = 'rejected'")).recordset[0].c;
    const pending = (await dmsPool.request().query("SELECT COUNT(*) as c FROM DmsApprovals WHERE Status = 'pending'")).recordset[0].c;

    const avgTime = await dmsPool.request().query(`
      SELECT AVG(DATEDIFF(HOUR, CreatedAt, ActionAt)) AS AvgHours
      FROM DmsApprovals
      WHERE Status IN ('approved', 'rejected') AND ActionAt IS NOT NULL
    `);

    const byStage = await dmsPool.request().query(`
      SELECT Stage, Status, COUNT(*) as Count
      FROM DmsApprovals
      GROUP BY Stage, Status
      ORDER BY Stage, Status
    `);

    res.json({
      total: totalApprovals,
      approved,
      rejected,
      pending,
      approvalRate: totalApprovals > 0 ? ((approved / (approved + rejected)) * 100).toFixed(1) : 0,
      averageApprovalHours: avgTime.recordset[0]?.AvgHours || 0,
      byStage: byStage.recordset,
    });
  } catch (err) {
    console.error("[Analytics] approval-stats error:", err.message);
    res.status(500).json({ error: "Failed to fetch approval stats" });
  }
});

// ─── 7K. AUDIT LOG ─────────────────────────────────────────────

// GET /api/audit-log
app.get("/api/audit-log", requireAdmin, async (req, res) => {
  try {
    const { action, user, entityType, entityId, startDate, endDate, page = 1, pageSize = 50 } = req.query;
    const pageNum = Math.max(1, Number(page));
    const pageSz = Math.min(200, Math.max(1, Number(pageSize)));
    const offset = (pageNum - 1) * pageSz;

    let whereClauses = ["1=1"];
    const request = dmsPool.request();

    if (action) {
      request.input("action", sql.NVarChar, action);
      whereClauses.push("Action = @action");
    }
    if (user) {
      request.input("user", sql.NVarChar, user);
      whereClauses.push("UserEmail = @user");
    }
    if (entityType) {
      request.input("entityType", sql.NVarChar, entityType);
      whereClauses.push("EntityType = @entityType");
    }
    if (entityId) {
      request.input("entityId", sql.Int, Number(entityId));
      whereClauses.push("EntityId = @entityId");
    }
    if (startDate) {
      request.input("startDate", sql.NVarChar, startDate);
      whereClauses.push("CreatedAt >= @startDate");
    }
    if (endDate) {
      request.input("endDate", sql.NVarChar, endDate);
      whereClauses.push("CreatedAt <= @endDate");
    }

    const whereStr = whereClauses.join(" AND ");

    const countReq = dmsPool.request();
    if (action) countReq.input("action", sql.NVarChar, action);
    if (user) countReq.input("user", sql.NVarChar, user);
    if (entityType) countReq.input("entityType", sql.NVarChar, entityType);
    if (entityId) countReq.input("entityId", sql.Int, Number(entityId));
    if (startDate) countReq.input("startDate", sql.NVarChar, startDate);
    if (endDate) countReq.input("endDate", sql.NVarChar, endDate);

    const countResult = await countReq.query(`SELECT COUNT(*) as total FROM DmsAuditLog WHERE ${whereStr}`);

    request.input("offset", sql.Int, offset);
    request.input("pageSz", sql.Int, pageSz);

    const result = await request.query(`
      SELECT * FROM DmsAuditLog
      WHERE ${whereStr}
      ORDER BY CreatedAt DESC
      OFFSET @offset ROWS FETCH NEXT @pageSz ROWS ONLY
    `);

    res.json({
      logs: result.recordset,
      total: countResult.recordset[0].total,
      page: pageNum,
      pageSize: pageSz,
      totalPages: Math.ceil(countResult.recordset[0].total / pageSz),
    });
  } catch (err) {
    console.error("[AuditLog] error:", err.message);
    res.status(500).json({ error: "Failed to fetch audit log" });
  }
});

// ─── 10. STATIC FILE SERVING + SPA FALLBACK ────────────────────

const distDir = path.join(__dirname, "../dist");
app.use(express.static(distDir));

// SPA fallback
app.use((req, res, next) => {
  if (req.path.startsWith("/api/") || req.path.startsWith("/auth/")) return next();
  res.sendFile(path.join(distDir, "index.html"));
});

// ─── Global error handler ───────────────────────────────────────
app.use((err, req, res, _next) => {
  console.error("[Server] Unhandled error:", err);
  res.status(500).json({ error: "Internal server error" });
});

async function runValidityReminderJob() {
  try {
    const docs = await dmsPool.request().query(`
      SELECT Id, Title, CreatorEmail, ValidTo
      FROM DmsDocuments
      WHERE Status != 'deleted'
        AND ValidTo IS NOT NULL
        AND ValidTo >= CAST(SYSUTCDATETIME() AS DATE)
        AND ValidTo <= DATEADD(DAY, 30, CAST(SYSUTCDATETIME() AS DATE))
        AND ValidityReminderSentAt IS NULL
    `);
    for (const d of docs.recordset) {
      const access = await dmsPool.request().input("docId", sql.Int, d.Id)
        .query("SELECT Email FROM DmsDocAccess WHERE DocId = @docId");
      const recipients = Array.from(new Set([d.CreatorEmail, ...access.recordset.map((x) => x.Email)].filter(Boolean)));
      if (recipients.length) {
        await sendEmail(
          recipients,
          `[DMS] Document Validity Expiring: ${d.Title}`,
          `<h3>Validity Reminder</h3><p><strong>${d.Title}</strong> is expiring on <strong>${new Date(d.ValidTo).toLocaleDateString("en-GB")}</strong>.</p><p>Please review and renew if required.</p>`
        );
      }
      await dmsPool.request().input("docId", sql.Int, d.Id)
        .query("UPDATE DmsDocuments SET ValidityReminderSentAt = SYSUTCDATETIME() WHERE Id = @docId");
    }
  } catch (err) {
    console.error("[ValidityReminderJob] error:", err.message);
  }
}

app.post("/api/admin/validity-reminders/run", requireAdmin, async (_req, res) => {
  await runValidityReminderJob();
  res.json({ ok: true });
});

// ─── 11. HTTPS SERVER BOOT ─────────────────────────────────────
async function boot() {
  await initPools();
  await ensureTables();
  await runValidityReminderJob();
  setInterval(() => {
    runValidityReminderJob().catch(() => {});
  }, 6 * 60 * 60 * 1000);

  try {
    const httpsOptions = {
      key: fs.readFileSync(path.resolve(__dirname, "../", process.env.TLS_KEY_FILE)),
      cert: fs.readFileSync(path.resolve(__dirname, "../", process.env.TLS_CERT_FILE)),
      ca: fs.readFileSync(path.resolve(__dirname, "../", process.env.TLS_CA_FILE)),
    };

    https.createServer(httpsOptions, app).listen(PORT, HOST, () => {
      console.log(`DMS HTTPS → https://${HOST === "0.0.0.0" ? "localhost" : HOST}:${PORT}`);
    });
  } catch (err) {
    console.error("[Server] TLS certificate error, falling back to HTTP:", err.message);
    app.listen(PORT, HOST, () => {
      console.log(`DMS HTTP → http://${HOST === "0.0.0.0" ? "localhost" : HOST}:${PORT}`);
    });
  }
}

boot().catch((err) => {
  console.error("[Server] Fatal boot error:", err);
  process.exit(1);
});
