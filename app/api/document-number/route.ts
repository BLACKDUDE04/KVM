import {
  ensureSchema,
  getCurrentUser,
  nextDocumentNumber,
} from "../data/route";

export async function GET(request: Request) {
  try {
    await ensureSchema();
    const user = await getCurrentUser(request);
    if (!user) return Response.json({ error: "Unauthorized" }, { status: 401 });
    const requested = new URL(request.url).searchParams.get("kind"),
      kind = requested === "return" ? "return" : "invoice";
    return Response.json({ number: await nextDocumentNumber(kind) });
  } catch (error) {
    return Response.json(
      {
        error:
          error instanceof Error
            ? error.message
            : "Unable to generate the next document number",
      },
      { status: 500 },
    );
  }
}
