import { getEnv } from "./env-context";

const SESSION_COOKIE = "billflow_session";
const SESSION_SECONDS = 60 * 60 * 24 * 7;
// Cloudflare Workers currently caps PBKDF2 at 100,000 iterations.
const PASSWORD_ITERATIONS = 100_000;

export type SessionUser = {
  id: number;
  name: string;
  user_id: string;
  email?: string | null;
  role: string;
  enabled: number;
  phone?: string | null;
  designation?: string | null;
  company?: string | null;
  timezone?: string | null;
  last_login_at?: string | null;
  created_at?: string | null;
};

const encoder = new TextEncoder();

const toHex = (bytes: Uint8Array) =>
  [...bytes].map((byte) => byte.toString(16).padStart(2, "0")).join("");

const fromHex = (value: string) => {
  const bytes = new Uint8Array(Math.floor(value.length / 2));
  for (let index = 0; index < bytes.length; index += 1)
    bytes[index] = Number.parseInt(value.slice(index * 2, index * 2 + 2), 16);
  return bytes;
};

export const normalizeUserId = (value: unknown) =>
  String(value || "")
    .trim()
    .toLowerCase();

export function validateUserId(value: unknown) {
  const userId = normalizeUserId(value);
  if (!/^[a-z0-9][a-z0-9._-]{2,31}$/.test(userId))
    throw new Error(
      "User ID must be 3–32 characters using letters, numbers, dot, dash or underscore",
    );
  return userId;
}

export function validatePassword(value: unknown) {
  const password = String(value || "");
  if (
    password.length < 8 ||
    !/[A-Za-z]/.test(password) ||
    !/[0-9]/.test(password)
  )
    throw new Error(
      "Password must contain at least 8 characters, including a letter and a number",
    );
  return password;
}

export async function createPassword(passwordValue: unknown) {
  const password = validatePassword(passwordValue),
    salt = crypto.getRandomValues(new Uint8Array(16)),
    key = await crypto.subtle.importKey(
      "raw",
      encoder.encode(password),
      "PBKDF2",
      false,
      ["deriveBits"],
    ),
    bits = await crypto.subtle.deriveBits(
      {
        name: "PBKDF2",
        salt,
        iterations: PASSWORD_ITERATIONS,
        hash: "SHA-256",
      },
      key,
      256,
    );
  return { hash: toHex(new Uint8Array(bits)), salt: toHex(salt) };
}

export async function verifyPassword(
  password: unknown,
  expectedHash: string,
  saltHex: string,
) {
  if (!expectedHash || !saltHex) return false;
  const key = await crypto.subtle.importKey(
      "raw",
      encoder.encode(String(password || "")),
      "PBKDF2",
      false,
      ["deriveBits"],
    ),
    bits = new Uint8Array(
      await crypto.subtle.deriveBits(
        {
          name: "PBKDF2",
          salt: fromHex(saltHex),
          iterations: PASSWORD_ITERATIONS,
          hash: "SHA-256",
        },
        key,
        256,
      ),
    ),
    actual = toHex(bits);
  if (actual.length !== expectedHash.length) return false;
  let difference = 0;
  for (let index = 0; index < actual.length; index += 1)
    difference |= actual.charCodeAt(index) ^ expectedHash.charCodeAt(index);
  return difference === 0;
}

const digest = async (value: string) =>
  toHex(
    new Uint8Array(
      await crypto.subtle.digest("SHA-256", encoder.encode(value)),
    ),
  );

const randomToken = () => toHex(crypto.getRandomValues(new Uint8Array(32)));

const cookieValue = (request: Request, value: string, maxAge: number) => {
  const secure = new URL(request.url).protocol === "https:" ? "; Secure" : "";
  return `${SESSION_COOKIE}=${value}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${maxAge}${secure}`;
};

export const clearSessionCookie = (request: Request) =>
  cookieValue(request, "", 0);

export const getSessionToken = (request: Request) => {
  const cookie = request.headers.get("cookie") || "";
  return (
    cookie
      .split(";")
      .map((part) => part.trim())
      .find((part) => part.startsWith(`${SESSION_COOKIE}=`))
      ?.slice(SESSION_COOKIE.length + 1) || ""
  );
};

export async function createSession(request: Request, userId: number) {
  const token = randomToken(),
    tokenHash = await digest(token),
    expiresAt = new Date(Date.now() + SESSION_SECONDS * 1000).toISOString();
  await getEnv()
    .DB.prepare(
      "INSERT INTO auth_sessions(user_id,token_hash,expires_at,created_at) VALUES(?,?,?,?)",
    )
    .bind(userId, tokenHash, expiresAt, new Date().toISOString())
    .run();
  return { cookie: cookieValue(request, token, SESSION_SECONDS), expiresAt };
}

export async function deleteSession(request: Request) {
  const token = getSessionToken(request);
  if (token)
    await getEnv()
      .DB.prepare("DELETE FROM auth_sessions WHERE token_hash = ?")
      .bind(await digest(token))
      .run();
}

export async function getSessionUser(request: Request) {
  const token = getSessionToken(request);
  if (!token) return null;
  const tokenHash = await digest(token),
    now = new Date().toISOString(),
    user = await getEnv()
      .DB.prepare(
        `SELECT u.id,u.name,u.user_id,u.email,u.role,u.enabled,u.phone,
                u.designation,u.company,u.timezone,u.last_login_at,u.created_at
         FROM auth_sessions s
         JOIN users u ON u.id = s.user_id
         WHERE s.token_hash = ? AND s.expires_at > ? AND u.enabled = 1
         LIMIT 1`,
      )
      .bind(tokenHash, now)
      .first<SessionUser>();
  if (!user)
    await getEnv()
      .DB.prepare("DELETE FROM auth_sessions WHERE token_hash = ?")
      .bind(tokenHash)
      .run();
  return user || null;
}

export async function credentialsConfigured() {
  const row = await getEnv()
    .DB.prepare(
      "SELECT COUNT(*) AS count FROM users WHERE user_id IS NOT NULL AND password_hash IS NOT NULL",
    )
    .first<{ count: number }>();
  return Number(row?.count || 0) > 0;
}
