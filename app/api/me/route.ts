import { getEnv } from "../../../lib/env-context";
import {
  createPassword,
  validateUserId,
  verifyPassword,
} from "../../../lib/auth";
import { ensureSchema, getCurrentUser } from "../data/route";

const profileFields = ["name", "phone", "designation", "company", "timezone"];
const safeUserSql = `SELECT id,name,user_id,role,enabled,phone,designation,company,timezone,last_login_at,created_at
                     FROM users WHERE id=?`;

export async function GET(request: Request) {
  try {
    await ensureSchema();
    const user = await getCurrentUser(request);
    if (!user)
      return Response.json(
        { error: "Your BillFlow session has expired" },
        { status: 401 },
      );
    return Response.json({ user });
  } catch (error) {
    return Response.json(
      {
        error:
          error instanceof Error ? error.message : "Unable to load profile",
      },
      { status: 500 },
    );
  }
}

export async function PATCH(request: Request) {
  try {
    await ensureSchema();
    const user = await getCurrentUser(request);
    if (!user) return Response.json({ error: "Unauthorized" }, { status: 401 });
    const body = (await request.json()) as Record<string, unknown>,
      requestedUserId = validateUserId(body.user_id || user.user_id),
      changingUserId = requestedUserId !== user.user_id,
      newPassword = String(body.new_password || ""),
      changingPassword = Boolean(newPassword),
      values: Record<string, unknown> = Object.fromEntries(
        profileFields
          .filter((field) => body[field] !== undefined)
          .map((field) => [field, String(body[field] || "").trim()]),
      );
    if (changingUserId || changingPassword) {
      const credential = await getEnv()
        .DB.prepare(
          "SELECT password_hash,password_salt FROM users WHERE id=? LIMIT 1",
        )
        .bind(user.id)
        .first<{ password_hash: string; password_salt: string }>();
      if (
        !credential ||
        !(await verifyPassword(
          body.current_password,
          credential.password_hash,
          credential.password_salt,
        ))
      )
        return Response.json(
          { error: "Current password is incorrect" },
          { status: 400 },
        );
      values.user_id = requestedUserId;
      if (changingPassword) {
        const password = await createPassword(newPassword);
        values.password_hash = password.hash;
        values.password_salt = password.salt;
        values.failed_login_count = 0;
        values.locked_until = null;
      }
    }
    const keys = Object.keys(values);
    if (!keys.length)
      return Response.json(
        { error: "No profile fields to update" },
        { status: 400 },
      );
    if (!String(values.name || user.name).trim())
      return Response.json({ error: "Name is required" }, { status: 400 });
    await getEnv()
      .DB.prepare(
        `UPDATE users SET ${keys.map((key) => `${key}=?`).join(",")} WHERE id=?`,
      )
      .bind(...keys.map((key) => values[key]), user.id)
      .run();
    const updated = await getEnv()
      .DB.prepare(safeUserSql)
      .bind(user.id)
      .first();
    return Response.json({ user: updated });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Unable to update profile";
    return Response.json(
      {
        error: /unique constraint/i.test(message)
          ? "That User ID is already in use"
          : message,
      },
      { status: /unique constraint/i.test(message) ? 409 : 500 },
    );
  }
}
