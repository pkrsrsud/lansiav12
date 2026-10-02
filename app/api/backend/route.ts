import { NextResponse } from "next/server";

export const dynamic = "force-dynamic";

export async function POST(request: Request) {
  try {
    const body = await request.json();
    const url = process.env.GOOGLE_APPS_SCRIPT_URL;
    const secret = process.env.GOOGLE_APPS_SCRIPT_SECRET;

    if (!url || !secret) {
      return NextResponse.json({ error: true, message: "Backend Google Apps Script belum dikonfigurasi." }, { status: 500 });
    }

    const gasResponse = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        secret,
        action: body.action,
        payload: body.payload || {},
        adminToken: body.adminToken || ""
      }),
      cache: "no-store"
    });

    const text = await gasResponse.text();
    let json: unknown;
    try { json = JSON.parse(text); } catch { json = { error: true, message: text || "Respons backend tidak valid." }; }

    return NextResponse.json(json, { status: gasResponse.ok ? 200 : 502 });
  } catch (error) {
    return NextResponse.json({ error: true, message: error instanceof Error ? error.message : "Gagal menghubungi backend." }, { status: 500 });
  }
}
