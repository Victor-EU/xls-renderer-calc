/**
 * The number-format seam.
 *
 * `TEXT(A1,"0.0%")` needs a full Excel number-format implementation, which is a
 * project of its own — and the renderer already carries one (`numfmt`, MIT) for
 * painting cells. Rather than duplicate it here and diverge, or take a
 * dependency and lose the engine's zero-dependency property, the engine asks for
 * a formatter and refuses honestly when none is installed.
 *
 * The renderer and the oracle harness both install `numfmt`; a bare `node`
 * consumer that never calls `setNumberFormatter` gets a ⚠ on TEXT() rather than
 * a plausible-looking string.
 */

export type NumberFormatter = (code: string, value: number | Date) => string;

let formatter: NumberFormatter | undefined;

export function setNumberFormatter(f: NumberFormatter | undefined): void {
  formatter = f;
}

export function getNumberFormatter(): NumberFormatter | undefined {
  return formatter;
}
