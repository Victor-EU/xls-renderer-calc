import type { Node } from '../ast.js';
import { Unsupported } from '../errors.js';
import type { EvalValue } from '../values.js';
import type { FnContext } from '../interpreter.js';

export type FnImpl = (args: EvalValue[], ctx: FnContext) => EvalValue;
export type LazyImpl = (
  nodes: Node[],
  ctx: FnContext,
  evaluate: (n: Node, ctx: FnContext) => EvalValue,
) => EvalValue;

export interface FnDef {
  readonly name: string;
  readonly minArgs: number;
  /** -1 means unbounded. */
  readonly maxArgs: number;
  /** NOW/TODAY/RAND — computed, but marked so the cached-value diff excludes them. */
  readonly volatile?: boolean;
  /** Evaluate arguments in array mode (SUMPRODUCT and friends). */
  readonly arrayArgs?: boolean;
  readonly call: FnImpl;
  /** Set for functions that must not evaluate every argument (IF, CHOOSE, IFS). */
  readonly lazy?: LazyImpl;
  /** Registered only to give a better refusal message — never returns a value. */
  readonly refused?: boolean;
}

const REGISTRY = new Map<string, FnDef>();

export function register(def: FnDef): void {
  if (REGISTRY.has(def.name)) throw new Error(`duplicate function registration: ${def.name}`);
  REGISTRY.set(def.name, def);
}

export function fn(
  name: string,
  minArgs: number,
  maxArgs: number,
  call: FnImpl,
  extra: Partial<Pick<FnDef, 'volatile' | 'arrayArgs'>> = {},
): void {
  register({ name, minArgs, maxArgs, call, ...extra });
}

export function lazyFn(name: string, minArgs: number, maxArgs: number, lazy: LazyImpl): void {
  register({
    name,
    minArgs,
    maxArgs,
    lazy,
    call: () => {
      throw new Error(`${name} is lazy`);
    },
  });
}

/**
 * Declare a function we recognise but deliberately do not evaluate.
 *
 * The distinction matters for the coverage backlog: `FN_UNKNOWN` might be a typo
 * in the generated model, while a refusal is our gap with a stated reason. Both
 * render ⚠ — neither ever renders a number.
 */
export function refuse(name: string, reason: string): void {
  register({
    name,
    minArgs: 0,
    maxArgs: -1,
    refused: true,
    call: () => {
      throw new Unsupported('FN_UNIMPLEMENTED', reason, name);
    },
  });
}

export function getFunction(name: string): FnDef | undefined {
  return REGISTRY.get(name);
}

export function registeredFunctions(): string[] {
  return [...REGISTRY.keys()].sort();
}

/** The published capability floor — what a generator prompt may safely emit. */
export function implementedFunctions(): string[] {
  return [...REGISTRY.values()]
    .filter((d) => !d.refused)
    .map((d) => d.name)
    .sort();
}

/** Functions we know of and deliberately refuse, with the stated reason. */
export function refusedFunctions(): string[] {
  return [...REGISTRY.values()]
    .filter((d) => d.refused)
    .map((d) => d.name)
    .sort();
}
