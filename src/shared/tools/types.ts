import type { CallToolResult } from '@modelcontextprotocol/server';
import { type ZodObject, type ZodType, z } from 'zod/v4';

export interface ToolContext {
  signal: AbortSignal;
  providerAccessToken?: string;
}

const RFC3339_TZ_REGEX = /Z$|[+-]\d{2}:\d{2}$/;
export const rfc3339 = z.string().refine((value) => RFC3339_TZ_REGEX.test(value), {
  message:
    'Timestamp must include timezone: append "Z" for UTC or an offset like "+01:00".',
});

export type ToolResult = CallToolResult;

export interface SharedToolDefinition<
  TInput extends ZodObject = ZodObject,
  TOutput extends ZodType | undefined = ZodType | undefined,
> {
  name: string;
  title?: string;
  description: string;
  inputSchema: TInput;
  outputSchema?: TOutput;
  handler: (args: z.infer<TInput>, context: ToolContext) => Promise<ToolResult>;
  annotations?: {
    title?: string;
    readOnlyHint?: boolean;
    destructiveHint?: boolean;
    idempotentHint?: boolean;
    openWorldHint?: boolean;
  };
}

export function defineTool<
  TInput extends ZodObject,
  TOutput extends ZodType | undefined = undefined,
>(
  definition: SharedToolDefinition<TInput, TOutput>,
): SharedToolDefinition<TInput, TOutput> {
  return definition;
}
