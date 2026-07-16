import {
  createPassword,
  createSession,
  credentialsConfigured,
  validateUserId,
} from "../../../../lib/auth";
import { getEnv } from "../../../../lib/env-context";
import { ensureSchema } from "../../data/route";

export async function POST(request: Request) {
  try {
    await ensureSchema();
    if (await credentialsConfigured())
      return Response.json(
        { error: "Administrator setup is already complete" },
        { status: 409 },
      );
    const body = (await request.json()) as Record<string, unknown>,
      name = String(body.name || "").trim(),
      userId = validateUserId(body.user_id),
      password = await createPassword(body.password);
    if (!name)
      return Response.json({ error: "Full name is required" }, { status: 400 });
    const db = getEnv().DB,
      existing = await db
        .prepare(
          "SELECT id,email FROM users ORDER BY CASE WHEN lower(role)='admin' THEN 0 ELSE 1 END,id LIMIT 1",
        )
        .first<{ id: number; email: string }>();
    let id = Number(existing?.id || 0);
    if (id) {
      await db
        .prepare(
          `UPDATE users
           SET name=?,user_id=?,password_hash=?,password_salt=?,role='Admin',enabled=1,
               failed_login_count=0,locked_until=NULL,created_at=COALESCE(created_at,?)
           WHERE id=?`,
        )
        .bind(
          name,
          userId,
          password.hash,
          password.salt,
          new Date().toISOString(),
          id,
        )
        .run();
    } else {
      const created = await db
        .prepare(
          `INSERT INTO users(name,email,user_id,password_hash,password_salt,role,enabled,failed_login_count,created_at)
           VALUES(?,?,?,?,?,'Admin',1,0,?) RETURNING id`,
        )
        .bind(
          name,
          `${userId}@billflow.local`,
          userId,
          password.hash,
          password.salt,
          new Date().toISOString(),
        )
        .first<{ id: number }>();
      id = Number(created?.id || 0);
    }
    if (!id) throw new Error("Unable to create the administrator account");
    const session = await createSession(request, id);
    return Response.json(
      { ok: true },
      { headers: { "set-cookie": session.cookie } },
    );
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Unable to complete setup";
    return Response.json(
      {
        error: /unique constraint/i.test(message)
          ? "That User ID is already in use"
          : message,
      },
      { status: /unique constraint/i.test(message) ? 409 : 400 },
    );
  }
}
