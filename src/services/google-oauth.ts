import { type OAuthProvider, readStandardTokenResponse } from '../oauth/provider';

/** What the server may do in the user's calendars. Users grant both at sign-in. */
export const CALENDAR_SCOPES = [
  'https://www.googleapis.com/auth/calendar.events', // create, change, move and delete events; respond
  'https://www.googleapis.com/auth/calendar.readonly', // list calendars, search events, free/busy
] as const;

/**
 * Google as the proxy's provider. The client ID and secret are the server's own Google OAuth
 * client (secrets `PROVIDER_CLIENT_ID` and `PROVIDER_CLIENT_SECRET`), whose authorized
 * redirect URI is `<server origin>/oauth/callback`.
 */
export const googleOAuth: OAuthProvider = {
  name: 'Google',
  authorizationUrl: 'https://accounts.google.com/o/oauth2/v2/auth',
  tokenUrl: 'https://oauth2.googleapis.com/token',
  scopes: CALENDAR_SCOPES,
  // Without offline access Google issues no refresh token. Without consent it issues one only
  // the first time a user signs in, and a second sign-in would leave the server without one.
  authorizationParams: { access_type: 'offline', prompt: 'consent' },
  // Google answers RFC 6749 token responses. A refresh response omits `refresh_token`, and
  // the proxy keeps the one it has.
  readTokens: readStandardTokenResponse,
};
