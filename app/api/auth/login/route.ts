import {
  createSession,
  normalizeUserId,
  verifyPassword,
} from "../../../../lib/auth";
import { getEnv } from "../../../../lib/env-context";
import { ensureSchema } from "../../data/route";

type LoginUser = {
  id: number;
  password_hash: string | null;
  password_salt: string | null;
  enabled: number;
  failed_login_count: number;
  locked_until: string | null;
};

export async function POST(request: Request) {
  try {
    await ensureSchema();
    const body = (await request.json()) as Record<string, unknown>,
      userId = normalizeUserId(body.user_id),
      password = String(body.password || ""),
      db = getEnv().DB,
      user = await db
        .prepare(
          `SELECT id,password_hash,password_salt,enabled,failed_login_count,locked_until
           FROM users WHERE lower(user_id)=lower(?) LIMIT 1`,
        )
        .bind(userId)
        .first<LoginUser>(),
      genericError = "Invalid User ID or password";
    if (!user || !user.password_hash || !user.password_salt) {
      await new Promise((resolve) => setTimeout(resolve, 250));
      return Response.json({ error: genericError }, { status: 401 });
    }
    if (!user.enabled)
      return Response.json(
        { error: "This user account has been disabled" },
        { status: 403 },
      );
    if (user.locked_until && new Date(user.locked_until).getTime() > Date.now())
      return Response.json(
        { error: "Too many attempts. Try again after 15 minutes." },
        { status: 429 },
      );
    const valid = await verifyPassword(
      password,
      user.password_hash,
      user.password_salt,
    );
    if (!valid) {
      const attempts = Number(user.failed_login_count || 0) + 1,
        lockedUntil =
          attempts >= 5
            ? new Date(Date.now() + 15 * 60 * 1000).toISOString()
            : null;
      await db
        .prepare(
          "UPDATE users SET failed_login_count=?,locked_until=? WHERE id=?",
        )
        .bind(lockedUntil ? 0 : attempts, lockedUntil, user.id)
        .run();
      return Response.json(
        {
          error: lockedUntil
            ? "Too many attempts. Login is locked for 15 minutes."
            : genericError,
        },
        { status: lockedUntil ? 429 : 401 },
      );
    }
    await db
      .prepare(
        "UPDATE users SET failed_login_count=0,locked_until=NULL,last_login_at=? WHERE id=?",
      )
      .bind(new Date().toISOString(), user.id)
      .run();
    await db
      .prepare("DELETE FROM auth_sessions WHERE expires_at <= ?")
      .bind(new Date().toISOString())
      .run();
    const session = await createSession(request, user.id);
    return Response.json(
      { ok: true },
      { headers: { "set-cookie": session.cookie } },
    );
  } catch (error) {
    return Response.json(
      { error: error instanceof Error ? error.message : "Unable to sign in" },
      { status: 500 },
    );
  }
}
