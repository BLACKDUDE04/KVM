import { getEnv } from "../../../lib/env-context";
import { reconcileLinkedData } from "../../../lib/reconcile";
import { authorize, ensureSchema } from "../data/route";

export async function POST(request: Request) {
  try {
    await ensureSchema();
    if (!(await authorize(request)))
      return Response.json(
        { error: "Unauthorized or disabled" },
        { status: 403 },
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
