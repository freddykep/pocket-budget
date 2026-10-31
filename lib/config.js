// Konfiguration ausschliesslich aus Umgebungsvariablen. Werte werden bei jedem Zugriff gelesen.
const env = (name, fallback) => process.env[name] || fallback;

export const config = {
  get clientId() { return env('MS_CLIENT_ID'); },
  get clientSecret() { return env('MS_CLIENT_SECRET'); },
  get authority() { return env('MS_AUTHORITY', 'consumers'); },
  get redirectUri() { return env('MS_REDIRECT_URI', 'http://localhost:3000/api/auth/callback'); },
  get encKey() { return env('TOKEN_ENC_KEY'); },
  get graphBase() { return env('GRAPH_BASE', 'https://graph.microsoft.com/v1.0'); },
  get loginBase() { return env('LOGIN_BASE', 'https://login.microsoftonline.com'); },
  get databaseUrl() { return env('DATABASE_URL') || env('POSTGRES_URL'); },
  get allowedAccounts() { return (process.env.ALLOWED_ACCOUNTS || '').split(',').map((s) => s.trim().toLowerCase()).filter(Boolean); },
  get dataDir() { return env('DATA_DIR', '.data'); },
  get secureCookies() { return config.redirectUri.startsWith('https://'); },
};

export const SCOPES = 'openid profile offline_access User.Read Files.ReadWrite';

export function assertConfigured() {
  const missing = ['MS_CLIENT_ID', 'MS_CLIENT_SECRET', 'TOKEN_ENC_KEY'].filter((n) => !process.env[n]);
  if (missing.length) {
    throw new Error(`Fehlende Umgebungsvariablen: ${missing.join(', ')} (siehe .env.example)`);
  }
}
