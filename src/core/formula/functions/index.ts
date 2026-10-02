/**
 * The function registry: every built-in, keyed by uppercase name.
 *
 * Later entries with a duplicate name are ignored, so shared families can be
 * listed without worrying about the ordering between modules.
 */

import type { FnDef, FnRegistry } from '../fntypes';
import { mathFns } from './math';
import { statsFns } from './stats';
import { logicalFns } from './logical';
import { arrayFns } from './array';
import { textFns } from './text';
import { dateFns } from './datetime';
import { financialFns } from './financial';
import { lookupFns, databaseFns } from './lookup';
import { infoFns } from './info';

const ALL: FnDef[] = [
  ...mathFns,
  ...statsFns,
  ...logicalFns,
  ...arrayFns,
  ...textFns,
  ...dateFns,
  ...financialFns,
  ...lookupFns,
  ...databaseFns,
  ...infoFns,
];

export function buildRegistry(): FnRegistry {
  const map: FnRegistry = new Map();
  for (const def of ALL) {
    if (map.has(def.name)) continue;
    map.set(def.name, def);
  }
  return map;
}

/** Names of every registered function, sorted, for the function picker. */
export function functionNames(): string[] {
  return [...buildRegistry().keys()].sort();
}

export type { FnDef, FnRegistry };
export { ALL as ALL_FUNCTIONS };
