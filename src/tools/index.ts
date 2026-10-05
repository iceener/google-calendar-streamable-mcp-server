import { checkAvailability } from './check-availability';
import { createEvent } from './create-event';
import { deleteEvent } from './delete-event';
import { listCalendars } from './list-calendars';
import { respondToEvent } from './respond-to-event';
import { searchEvents } from './search-events';
import { updateEvent } from './update-event';

/** Every tool, in the order clients list them. The order is part of the published contract. */
export const tools = [
  listCalendars,
  searchEvents,
  checkAvailability,
  createEvent,
  updateEvent,
  deleteEvent,
  respondToEvent,
];
