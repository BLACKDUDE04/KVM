import { getEnv } from "../../../lib/env-context";
import { reconcileLinkedData } from "../../../lib/reconcile";
import { ensureSchema, getCurrentUser } from "../data/route";

export async function POST(request: Request) {
  try {
    await ensureSchema();
    const user = await getCurrentUser(request);
    if (!user || user.role.toLowerCase() === "viewer")
      return Response.json(
        { error: "Unauthorized or disabled" },
        { status: 401 },
      );
    return Response.json(await reconcileLinkedData(getEnv().DB));
  } catch (error) {
    return Response.json(
      {
        error:
          error instanceof Error
            ? error.message
            : "Unable to link imported data",
      },
      { status: 500 },
    );
  }
}
