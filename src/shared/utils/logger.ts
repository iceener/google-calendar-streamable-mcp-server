export type LogLevel = 'debug' | 'info' | 'warning' | 'error';
interface LogData {
  message: string;
  [key: string]: unknown;
}
const LEVELS: Record<LogLevel, number> = { debug: 0, info: 1, warning: 2, error: 3 };
const SENSITIVE = [
  'password',
  'token',
  'secret',
  'key',
  'authorization',
  'access_token',
  'refresh_token',
];
let currentLevel: LogLevel = 'info';
function output(level: LogLevel, name: string, data: LogData): void {
  if (LEVELS[level] < LEVELS[currentLevel]) return;
  const sanitized = Object.fromEntries(
    Object.entries(data).map(([key, value]) => [
      key,
      SENSITIVE.some((sensitive) => key.toLowerCase().includes(sensitive))
        ? '[REDACTED]'
        : value,
    ]),
  );
  const line = JSON.stringify({
    timestamp: new Date().toISOString(),
    level,
    logger: name,
    ...sanitized,
  });
  if (level === 'error') console.error(line);
  else if (level === 'warning') console.warn(line);
  else if (level === 'debug') console.debug(line);
  else console.info(line);
}
export const sharedLogger = {
  setLevel(level: LogLevel): void {
    currentLevel = level;
  },
  debug(name: string, data: LogData): void {
    output('debug', name, data);
  },
  info(name: string, data: LogData): void {
    output('info', name, data);
  },
  warning(name: string, data: LogData): void {
    output('warning', name, data);
  },
  error(name: string, data: LogData): void {
    output('error', name, data);
  },
};
export const logger = sharedLogger;
