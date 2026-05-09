import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type"
};

Deno.serve(async (request) => {
  if (request.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  try {
    const { email, full_name } = await request.json();

    const supabaseUrl = Deno.env.get("SUPABASE_URL");
    const serviceRoleKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
    const resendApiKey = Deno.env.get("RESEND_API_KEY");
    const fromEmail = Deno.env.get("ACCESS_REQUEST_FROM_EMAIL");

    if (!supabaseUrl || !serviceRoleKey || !resendApiKey || !fromEmail) {
      return new Response(JSON.stringify({ error: "Missing function environment variables." }), {
        status: 500,
        headers: { ...corsHeaders, "Content-Type": "application/json" }
      });
    }

    const supabase = createClient(supabaseUrl, serviceRoleKey);
    const { data: admins, error } = await supabase
      .from("profiles")
      .select("email,full_name")
      .eq("role", "admin")
      .eq("access_status", "approved");

    if (error) {
      throw error;
    }

    const recipients = (admins || [])
      .map((item) => String(item.email || "").trim())
      .filter(Boolean);

    if (!recipients.length) {
      return new Response(JSON.stringify({ delivered: false, reason: "No approved admin recipients." }), {
        status: 200,
        headers: { ...corsHeaders, "Content-Type": "application/json" }
      });
    }

    const payload = {
      from: fromEmail,
      to: recipients,
      subject: "Nova solicitacao de acesso ao CRM Pax",
      html: `
        <div style="font-family:Arial,sans-serif;color:#102419">
          <h2>Nova solicitacao de acesso</h2>
          <p><strong>Nome:</strong> ${String(full_name || "").trim() || "-"}</p>
          <p><strong>E-mail:</strong> ${String(email || "").trim() || "-"}</p>
          <p>Acesse a aba <strong>Solicitacoes</strong> no CRM para aprovar ou recusar o cadastro.</p>
        </div>
      `
    };

    const response = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${resendApiKey}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      const message = await response.text();
      return new Response(JSON.stringify({ error: message }), {
        status: 502,
        headers: { ...corsHeaders, "Content-Type": "application/json" }
      });
    }

    return new Response(JSON.stringify({ delivered: true }), {
      status: 200,
      headers: { ...corsHeaders, "Content-Type": "application/json" }
    });
  } catch (error) {
    return new Response(JSON.stringify({ error: error instanceof Error ? error.message : "Unknown error" }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" }
    });
  }
});
