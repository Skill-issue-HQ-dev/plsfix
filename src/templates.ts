// The template catalogue the Excel adapter reads: every ready block as a plain
// Template, and the one lookup `insertTemplate` uses. The list, its order, the
// groups and the ids the Insert buttons carry live in the registry
// (src/templates/index.ts); every grid lives in `template-blocks.ts` or in a
// file of its own under src/templates/, and the grammar in `template-cells.ts`,
// both of which this module re-exports as one entry point.

import { TEMPLATES as ENTRIES } from "./templates/index";
import { type Template } from "./template-cells";

export {
  placeholderRefs,
  resolveFormula,
  type CellRef,
  type Template,
  type TemplateCell,
  type TemplateCellKind,
  type TemplateFormat,
} from "./template-cells";

/** Every block, in the order the picker lists them. */
export const TEMPLATES: Template[] = ENTRIES.map((entry) => entry.block);

export function templateById(id: string): Template | null {
  return TEMPLATES.find((template) => template.id === id) ?? null;
}
