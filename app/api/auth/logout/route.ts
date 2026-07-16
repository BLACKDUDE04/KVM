import { clearSessionCookie, deleteSession } from "../../../../lib/auth";
import { ensureSchema } from "../../data/route";

export async function POST(request: Request) {
  try {
    await ensureSchema();
    await deleteSession(request);
  } catch {
    // The browser cookie must still be cleared if the server session expired.
  }
  return Response.json(
    { ok: true },
    { headers: { "set-cookie": clearSessionCookie(request) } },
  );
}
