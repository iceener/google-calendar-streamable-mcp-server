import { z } from 'zod/v4';

const EventDateTimeSchema = z.object({
  dateTime: z.string().optional(),
  date: z.string().optional(),
  timeZone: z.string().optional(),
});

const AttendeeSchema = z.object({
  email: z.string(),
  displayName: z.string().optional(),
  responseStatus: z
    .enum(['needsAction', 'declined', 'tentative', 'accepted'])
    .optional(),
  optional: z.boolean().optional(),
  organizer: z.boolean().optional(),
  self: z.boolean().optional(),
});

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
  creator: z
    .object({ email: z.string(), displayName: z.string().optional() })
    .optional(),
  eventType: z
    .enum([
      'default',
      'birthday',
      'focusTime',
      'fromGmail',
      'outOfOffice',
      'workingLocation',
    ])
    .optional(),
  visibility: z.enum(['default', 'public', 'private', 'confidential']).optional(),
  colorId: z.string().optional(),
  recurringEventId: z.string().optional(),
  recurrence: z.array(z.string()).optional(),
  reminders: z
    .object({
      useDefault: z.boolean(),
      overrides: z
        .array(
          z.object({
            method: z.enum(['popup', 'email']),
            minutes: z.number(),
          }),
        )
        .optional(),
    })
    .optional(),
  created: z.string().optional(),
  updated: z.string().optional(),
});

export const ListCalendarsOutputSchema = z.object({
  items: z.array(
    z.object({
      id: z.string(),
      summary: z.string(),
      description: z.string().optional(),
      primary: z.boolean().optional(),
      backgroundColor: z.string().optional(),
      foregroundColor: z.string().optional(),
      accessRole: z.enum(['owner', 'writer', 'reader', 'freeBusyReader']),
      timeZone: z.string().optional(),
    }),
  ),
});

const SearchEventOutputSchema = CalendarEventOutputSchema.partial().extend({
  calendarId: z.string().optional(),
  calendarName: z.string().optional(),
});

export const SearchEventsOutputSchema = z.object({
  items: z.array(SearchEventOutputSchema),
  calendarsSearched: z.array(z.string()),
  nextPageToken: z.string().optional(),
});

export const AvailabilityOutputSchema = z.object({
  timeMin: z.string(),
  timeMax: z.string(),
  calendars: z.record(
    z.string(),
    z.object({
      busy: z.array(z.object({ start: z.string(), end: z.string() })),
      errors: z.array(z.object({ domain: z.string(), reason: z.string() })).optional(),
    }),
  ),
});

export const DeleteEventOutputSchema = z.object({
  success: z.literal(true),
  eventId: z.string(),
  calendarId: z.string(),
});

export const RespondToEventOutputSchema = z.object({
  ok: z.literal(true),
  action: z.literal('respond_to_event'),
  response: z.enum(['accepted', 'declined', 'tentative']),
  event: CalendarEventOutputSchema,
});
