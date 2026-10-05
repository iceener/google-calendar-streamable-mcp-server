import * as z from 'zod/v4';

/**
 * Schemas shared by several tools. The published tool schemas are part of the server's
 * contract (`tests/contract.test.ts`): change them only on purpose.
 */

/** Google needs timestamps with a zone. */
export const rfc3339 = z.string().refine((value) => /Z$|[+-]\d{2}:\d{2}$/.test(value), {
  message: 'Timestamp must include timezone: append "Z" for UTC or an offset like "+01:00".',
});

/**
 * The email pattern of Zod 4.4, which published these schemas first. Zod 4.6 changed the
 * default `.email()` pattern; attendee lists keep this one, so the published input schemas
 * and their validation stay as they were.
 */
const ATTENDEE_EMAIL_SOURCE = String.raw`^(?!\.)(?!.*\.\.)([A-Za-z0-9_'+\-\.]*)[A-Za-z0-9_+-]@([A-Za-z0-9][A-Za-z0-9\-]*\.)+[A-Za-z]{2,}$`;
// A string, not a regex literal: the published pattern is its source, character for character.
const ATTENDEE_EMAIL = new RegExp(ATTENDEE_EMAIL_SOURCE);

export const attendeeEmail = () => z.string().email({ pattern: ATTENDEE_EMAIL });

export const RemindersSchema = z.object({
  useDefault: z.boolean(),
  overrides: z
    .array(
      z.object({
        method: z.enum(['popup', 'email']),
        minutes: z.number().int().min(0).max(40320),
      }),
    )
    .optional(),
});

const EventDateTimeSchema = z.object({
  dateTime: z.string().optional(),
  date: z.string().optional(),
  timeZone: z.string().optional(),
});

const AttendeeSchema = z.object({
  email: z.string(),
  displayName: z.string().optional(),
  responseStatus: z.enum(['needsAction', 'declined', 'tentative', 'accepted']).optional(),
  optional: z.boolean().optional(),
  organizer: z.boolean().optional(),
  self: z.boolean().optional(),
});

/** A Google Calendar event, as create_event, update_event and respond_to_event return it. */
export const CalendarEventOutputSchema = z.object({
  id: z.string(),
  summary: z.string().optional(),
  description: z.string().optional(),
  start: EventDateTimeSchema.optional(),
  end: EventDateTimeSchema.optional(),
  location: z.string().optional(),
  status: z.enum(['confirmed', 'tentative', 'cancelled']).optional(),
  htmlLink: z.string().optional(),
  hangoutLink: z.string().optional(),
  conferenceData: z.record(z.string(), z.unknown()).optional(),
  attendees: z.array(AttendeeSchema).optional(),
  organizer: z
    .object({
      email: z.string(),
      displayName: z.string().optional(),
      self: z.boolean().optional(),
    })
    .optional(),
  creator: z.object({ email: z.string(), displayName: z.string().optional() }).optional(),
  eventType: z
    .enum(['default', 'birthday', 'focusTime', 'fromGmail', 'outOfOffice', 'workingLocation'])
    .optional(),
  visibility: z.enum(['default', 'public', 'private', 'confidential']).optional(),
  colorId: z.string().optional(),
  recurringEventId: z.string().optional(),
  recurrence: z.array(z.string()).optional(),
  reminders: z
    .object({
      useDefault: z.boolean(),
      overrides: z
        .array(z.object({ method: z.enum(['popup', 'email']), minutes: z.number() }))
        .optional(),
    })
    .optional(),
  created: z.string().optional(),
  updated: z.string().optional(),
});

/** `YYYY-MM-DD`: an all-day date rather than a date and time. */
export function isAllDayDate(value: string): boolean {
  return /^\d{4}-\d{2}-\d{2}$/.test(value);
}
