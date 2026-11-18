import type { enUS } from "./en-US.js";

/** English message IDs stay stable when a translation changes. */
export type MessageKey = keyof typeof enUS.messages;
export interface LocaleDefinition {
  readonly name: string;
  readonly messages: Readonly<Partial<Record<MessageKey, string>>>;
}
export type Locale = "zh-CN" | "en-US" | LocaleDefinition;
export type Translator = (key: MessageKey, ...values: readonly (string | number | undefined)[]) => string;
