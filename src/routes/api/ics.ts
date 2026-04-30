import { createFileRoute } from "@tanstack/react-router";

const ALLOWED_HOSTS = [
  "outlook.office365.com",
  "outlook.office.com",
  "outlook.live.com",
];

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

export const Route = createFileRoute("/api/ics")({
  server: {
    handlers: {
      OPTIONS: async () => new Response(null, { status: 204, headers: corsHeaders }),
      GET: async ({ request }: { request: Request }) => {
        const url = new URL(request.url);
        const target = url.searchParams.get("url");
        if (!target) {
          return new Response(JSON.stringify({ error: "Missing url" }), {
            status: 400,
            headers: { "Content-Type": "application/json", ...corsHeaders },
          });
        }
        let parsed: URL;
        try {
          parsed = new URL(target);
        } catch {
          return new Response(JSON.stringify({ error: "Invalid url" }), {
            status: 400,
            headers: { "Content-Type": "application/json", ...corsHeaders },
          });
        }
        if (parsed.protocol !== "https:" || !ALLOWED_HOSTS.includes(parsed.hostname)) {
          return new Response(JSON.stringify({ error: "Host not allowed" }), {
            status: 400,
            headers: { "Content-Type": "application/json", ...corsHeaders },
          });
        }
        try {
          const upstream = await fetch(parsed.toString(), {
            headers: { Accept: "text/calendar, text/plain, */*" },
          });
          const body = await upstream.text();
          return new Response(body, {
            status: upstream.status,
            headers: {
              "Content-Type": "text/calendar; charset=utf-8",
              ...corsHeaders,
            },
          });
        } catch (e) {
          return new Response(JSON.stringify({ error: "Upstream fetch failed" }), {
            status: 502,
            headers: { "Content-Type": "application/json", ...corsHeaders },
          });
        }
      },
    },
  },
} as any);
