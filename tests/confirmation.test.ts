import { describe, expect, it, vi } from "vitest";
import { restoreFocusAfterConfirmation, type ConfirmationFocusTarget } from "../src/renderer/confirmation";

function target(isConnected: boolean): ConfirmationFocusTarget & { focus: ReturnType<typeof vi.fn> } {
  return { isConnected, focus: vi.fn() };
}

describe("confirmation focus restoration", () => {
  it("returns focus to the delete trigger after cancellation", () => {
    const trigger = target(true);
    const fallback = target(true);

    expect(restoreFocusAfterConfirmation(trigger, fallback)).toBe(trigger);
    expect(trigger.focus).toHaveBeenCalledWith({ preventScroll: true });
    expect(fallback.focus).not.toHaveBeenCalled();
  });

  it("focuses the dictionary input after the deleted trigger leaves the document", () => {
    const deletedTrigger = target(false);
    const dictionaryInput = target(true);

    expect(restoreFocusAfterConfirmation(deletedTrigger, dictionaryInput)).toBe(dictionaryInput);
    expect(deletedTrigger.focus).not.toHaveBeenCalled();
    expect(dictionaryInput.focus).toHaveBeenCalledWith({ preventScroll: true });
  });

  it("does nothing when neither target remains connected", () => {
    const trigger = target(false);
    const fallback = target(false);

    expect(restoreFocusAfterConfirmation(trigger, fallback)).toBeNull();
    expect(trigger.focus).not.toHaveBeenCalled();
    expect(fallback.focus).not.toHaveBeenCalled();
  });
});
