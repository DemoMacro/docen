/**
 * Type guard utilities for DOCX processing
 */

/**
 * Type guard factory function
 * Creates a type guard function that checks if a value is one of the valid values
 *
 * @param validValues - Readonly array of valid string values
 * @returns Type guard function
 *
 * @example
 * const isValidAlign = createStringValidator(["left", "right", "center"] as const);
 * if (isValidAlign(value)) {
 *   // value is typed as "left" | "right" | "center"
 * }
 */
export function createStringValidator<T extends string>(validValues: readonly T[]) {
  return (value: string): value is T => {
    return validValues.includes(value as T);
  };
}
