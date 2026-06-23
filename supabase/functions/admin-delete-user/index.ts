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
    const supabaseUrl = Deno.env.get("SUPABASE_URL");
    const serviceRoleKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
    const anonKey = Deno.env.get("SUPABASE_ANON_KEY");
    const authHeader = request.headers.get("Authorization");

    if (!supabaseUrl || !serviceRoleKey || !anonKey || !authHeader) {
      return new Response(JSON.stringify({ error: "Missing function environment variables or auth header." }), {
        status: 500,
        headers: { ...corsHeaders, "Content-Type": "application/json" }
      });
    }

    const authClient = createClient(supabaseUrl, anonKey, {
      global: {
        headers: {
          Authorization: authHeader
        }
      }
    });

    const { data: authData, error: authError } = await authClient.auth.getUser();
    if (authError || !authData.user) {
      return new Response(JSON.stringify({ error: "Unauthorized." }), {
        status: 401,
        headers: { ...corsHeaders, "Content-Type": "application/json" }
      });
    }

    const adminClient = createClient(supabaseUrl, serviceRoleKey);
    const { data: requesterProfile, error: requesterError } = await adminClient
      .from("profiles")
      .select("id, role, access_status")
      .eq("id", authData.user.id)
      .maybeSingle();

    if (
      requesterError ||
      !requesterProfile ||
      String(requesterProfile.role || "").toLowerCase() !== "admin" ||
      String(requesterProfile.access_status || "").toLowerCase() !== "approved"
    ) {
      return new Response(JSON.stringify({ error: "Forbidden." }), {
        status: 403,
        headers: { ...corsHeaders, "Content-Type": "application/json" }
      });
    }

    const { profileId } = await request.json();
    const targetId = String(profileId || "").trim();
    if (!targetId) {
      return new Response(JSON.stringify({ error: "profileId is required." }), {
        status: 400,
        headers: { ...corsHeaders, "Content-Type": "application/json" }
      });
    }

    if (targetId === authData.user.id) {
      return new Response(JSON.stringify({ error: "You cannot delete your own account." }), {
        status: 400,
        headers: { ...corsHeaders, "Content-Type": "application/json" }
      });
    }

    const { data: targetProfile, error: targetError } = await adminClient
      .from("profiles")
      .select("id, email, full_name, role, access_status")
      .eq("id", targetId)
      .maybeSingle();

    if (targetError || !targetProfile) {
      return new Response(JSON.stringify({ error: "User not found." }), {
        status: 404,
        headers: { ...corsHeaders, "Content-Type": "application/json" }
      });
    }

    if (
      String(targetProfile.role || "").toLowerCase() === "admin" &&
      String(targetProfile.access_status || "").toLowerCase() === "approved"
    ) {
      const { count, error: countError } = await adminClient
        .from("profiles")
        .select("id", { count: "exact", head: true })
        .eq("role", "admin")
        .eq("access_status", "approved");

      if (countError) {
        throw countError;
      }

      if ((count || 0) <= 1) {
        return new Response(JSON.stringify({ error: "Nao e permitido excluir o ultimo administrador aprovado." }), {
          status: 400,
          headers: { ...corsHeaders, "Content-Type": "application/json" }
        });
      }
    }

    const cleanupOps = await Promise.all([
      adminClient.from("leads").update({ assigned_to: null }).eq("assigned_to", targetId),
      adminClient.from("leads").update({ created_by: null }).eq("created_by", targetId),
      adminClient.from("stages").update({ created_by: null }).eq("created_by", targetId),
      adminClient.from("profiles").update({ approved_by: null }).eq("approved_by", targetId),
      adminClient.from("access_requests").update({ reviewed_by: null }).eq("reviewed_by", targetId),
      adminClient.from("admin_requests").update({ reviewed_by: null }).eq("reviewed_by", targetId),
      adminClient.from("admin_requests").delete().eq("requested_by_id", targetId),
      adminClient.from("access_requests").delete().eq("auth_user_id", targetId),
      targetProfile.email
        ? adminClient.from("access_requests").delete().eq("email", targetProfile.email)
        : Promise.resolve({ error: null }),
      adminClient.from("change_history").delete().eq("user_id", targetId)
    ]);

    const cleanupError = cleanupOps.find((result) => result.error)?.error;
    if (cleanupError) {
      throw cleanupError;
    }

    const { error: deleteAuthError } = await adminClient.auth.admin.deleteUser(targetId);
    if (deleteAuthError) {
      throw deleteAuthError;
    }

    const { error: deleteProfileError } = await adminClient
      .from("profiles")
      .delete()
      .eq("id", targetId);

    if (deleteProfileError) {
      throw deleteProfileError;
    }

    return new Response(JSON.stringify({ deleted: true }), {
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
