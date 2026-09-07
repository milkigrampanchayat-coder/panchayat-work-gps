const GAS_URL = 'https://script.google.com/macros/s/AKfycbzdSDqaoWLvJuYLy6PlGZNam1kzLTA_7BgTgAuBi3reHC7-opnqT0kdLVtFYQ7USz20/exec';

function corsHeaders(extra={}) {
  return {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    ...extra
  };
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === "/health") {
      return new Response(JSON.stringify({ok:true,service:"panchayat-work-gps"}), {
        status: 200,
        headers: {...corsHeaders(), "Content-Type":"application/json"}
      });
    }

    if (url.pathname === "/api") {
      if (request.method === "OPTIONS") {
        return new Response(null, {status:204, headers:corsHeaders()});
      }

      const target = new URL(GAS_URL);
      url.searchParams.forEach((value,key)=>target.searchParams.set(key,value));

      const headers = new Headers();
      headers.set("Accept","application/json,text/plain,*/*");
      if (request.method !== "GET" && request.method !== "HEAD") {
        headers.set("Content-Type","application/json;charset=UTF-8");
      }

      // Read the incoming body once, then send a fresh request body.
      // This avoids forwarding the browser's Content-Length/stream metadata
      // through Google's Apps Script redirect, which can produce HTTP 500.
      let body;
      if (request.method !== "GET" && request.method !== "HEAD") {
        body = await request.arrayBuffer();
      }

      let upstream;
      try {
        upstream = await fetch(target.toString(), {
          method: request.method,
          headers,
          body,
          redirect: "follow"
        });
      } catch (e) {
        return new Response(JSON.stringify({
          success:false,
          message:"GAS connection failed: "+String(e.message||e)
        }), {
          status:502,
          headers:{...corsHeaders(), "Content-Type":"application/json"}
        });
      }

      const outHeaders = new Headers(upstream.headers);
      Object.entries(corsHeaders()).forEach(([k,v])=>outHeaders.set(k,v));

      return new Response(upstream.body, {
        status: upstream.status,
        statusText: upstream.statusText,
        headers: outHeaders
      });
    }

    return env.ASSETS.fetch(request);
  }
};
