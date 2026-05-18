module.exports = function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  const appUrl = process.env.APP_URL || `https://${req.headers.host}`;
  res.json({
    configured: !!(process.env.GOOGLE_CLIENT_ID && process.env.GOOGLE_CLIENT_SECRET),
    authenticated: !!(process.env.GOOGLE_CLIENT_ID && process.env.GOOGLE_CLIENT_SECRET && process.env.GOOGLE_REFRESH_TOKEN),
    hasRefreshToken: !!process.env.GOOGLE_REFRESH_TOKEN,
    folderId: process.env.GOOGLE_DRIVE_FOLDER_ID || null,
    appUrl,
    callbackUrl: `${appUrl}/api/google-auth/callback`,
    authUrl: `${appUrl}/api/google-auth`,
  });
};
