import { credentialsConfigured, getSessionUser } from "../../../../lib/auth";
import { ensureSchema } from "../../data/route";

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
    return Response.json(
      {
        error: error instanceof Error ? error.message : "Unable to check login",
      },
      { status: 500 },
    );
  }
}
