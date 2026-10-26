/*
 * User-facing messages in English and Arabic. The ribbon labels themselves are
 * localized in manifest.xml; these are the messages shown by the commands.
 */

/* global Office */

export type CommandName = "matchSize" | "matchWidth" | "matchHeight" | "align";

interface Strings {
  commandNames: Record<CommandName, string>;
  needsTwo(command: string, selected: number): string;
  swapNeedsExactlyTwo(selected: number): string;
  needsApi(api: string): string;
  needsText: string;
  needsSelection: string;
  unexpected(detail: string): string;
}

const en: Strings = {
  commandNames: {
    matchSize: "Match Size",
    matchWidth: "Match Width",
    matchHeight: "Match Height",
    align: "Align",
  },
  needsTwo: (command, selected) =>
    `${command} needs at least 2 selected shapes (${selected} selected).`,
  swapNeedsExactlyTwo: (selected) => `Swap needs exactly 2 selected shapes (${selected} selected).`,
  needsApi: (api) => `Nasaq requires ${api}. Please update PowerPoint.`,
  needsText: "Select some text, or shapes that contain text, to align text.",
  needsSelection: "Select at least one shape first.",
  unexpected: (detail) => `Something went wrong: ${detail}`,
};

const ar: Strings = {
  commandNames: {
    matchSize: "مطابقة الحجم",
    matchWidth: "مطابقة العرض",
    matchHeight: "مطابقة الارتفاع",
    align: "المحاذاة",
  },
  needsTwo: (command, selected) =>
    `يتطلب أمر «${command}» تحديد شكلين على الأقل (عدد الأشكال المحددة: ${selected}).`,
  swapNeedsExactlyTwo: (selected) =>
    `يتطلب التبديل تحديد شكلين بالضبط (عدد الأشكال المحددة: ${selected}).`,
  needsApi: (api) => `يتطلب نَسَق ${api}. يُرجى تحديث PowerPoint.`,
  needsText: "حدد نصًا، أو أشكالًا تحتوي على نص، لمحاذاة النص.",
  needsSelection: "حدد شكلًا واحدًا على الأقل أولًا.",
  unexpected: (detail) => `حدث خطأ: ${detail}`,
};

function isArabic(): boolean {
  return (Office.context.displayLanguage ?? "").toLowerCase().startsWith("ar");
}

/** Messages in the current Office display language (Arabic or English). */
export function strings(): Strings {
  return isArabic() ? ar : en;
}
