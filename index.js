export default {
  async fetch(request, env) {
    const response = await env.ASSETS.fetch(request);
    const newHeaders = new Headers(response.headers);
    newHeaders.set('Cache-Control', 'no-cache, no-store, must-revalidate');
    newHeaders.set('Pragma', 'no-cache');
    newHeaders.set('Expires', '0');
    return new Response(response.body, {
      status: response.status,
      statusText: response.statusText,
      headers: newHeaders
    });
  }
};
