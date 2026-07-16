import { credentialsConfigured, getSessionUser } from "../../../../lib/auth";
import { ensureSchema } from "../../data/route";

function sessionError(error: unknown) {
  const message =
    error instanceof Error ? error.message : "Unable to check login";
  if (/HTTP error! status: (400|401|403)/i.test(message))
    return {
      code: "DATABASE_CONNECTION_REJECTED",
      error:
        "Turso rejected the database connection. Check that TURSO_DATABASE_URL and TURSO_AUTH_TOKEN are copied from the same database and are available to Netlify Functions.",
    };
  if (/fetch failed|network|timeout|timed out/i.test(message))
    return {
      code: "DATABASE_UNREACHABLE",
      error:
        "BillFlow cannot reach Turso. Check the database URL and try again.",
    };
  return { code: "SESSION_CHECK_FAILED", error: message };
}

export async function GET(request: Request) {
  try {
    await ensureSchema();
    const [user, configured] = await Promise.all([
      getSessionUser(request),
      credentialsConfigured(),
    ]);
    return Response.json({
      authenticated: Boolean(user),
      needsSetup: !configured,
      user: user || undefined,
    });
  } catch (error) {
    const failure = sessionError(error);
    return Response.json(
      failure,
      { status: 503 },
    );
  }
}
