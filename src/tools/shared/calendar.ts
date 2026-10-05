import type { ServerContext } from '@modelcontextprotocol/server';
import { providerToken } from '../../oauth/verifier';
import { type ToolErrorResult, toolError } from '../../platform/primitives';
import type { Deps } from '../../server';
import { CalendarApiError, type CalendarClient } from '../../services/google-calendar';

/**
 * A Calendar client acting as the signed-in user, with the Google access token the OAuth proxy
 * resolved for this request. `undefined` without one, for example when authentication is off.
 */
export function calendarFor(ctx: ServerContext, deps: Deps): CalendarClient | undefined {
  const token = providerToken(ctx.http?.authInfo);
  return token ? deps.calendar(token, ctx.mcpReq.signal) : undefined;
}

export function signInRequired(): ToolErrorResult {
  return toolError('Authentication required. Please authenticate with Google Calendar.');
}

/** A failed call, told to the model with Google's own message. */
export function calendarFailure(action: string, error: unknown): ToolErrorResult {
  return toolError(`Failed to ${action}: ${(error as Error).message}`);
}

/**
 * Run a change against `calendarId`, and once more against `primary` if Google answers 404.
 * Google lists events under the owner's email address but may expect `primary` for changes,
 * or the other way round.
 */
export async function withPrimaryFallback<T>(
  calendarId: string,
  attempt: (calendarId: string) => Promise<T>,
): Promise<T> {
  try {
    return await attempt(calendarId);
  } catch (error) {
    if (error instanceof CalendarApiError && error.isNotFound && calendarId !== 'primary') {
      return attempt('primary');
    }
    throw error;
  }
}
