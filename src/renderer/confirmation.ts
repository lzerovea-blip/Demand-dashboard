export interface ConfirmationFocusTarget {
  readonly isConnected: boolean;
  focus(options?: FocusOptions): void;
}

export function restoreFocusAfterConfirmation(
  preferred: ConfirmationFocusTarget | null,
  fallback: ConfirmationFocusTarget | null,
): ConfirmationFocusTarget | null {
  const target = preferred?.isConnected ? preferred : fallback?.isConnected ? fallback : null;
  target?.focus({ preventScroll: true });
  return target;
}
