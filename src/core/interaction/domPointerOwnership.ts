import type { SurfaceId } from "../scene";
import {
  nativeInputTarget,
  surfaceInputTarget,
  type InputTargetResolution
} from "./InputRouter";

export type DomPointerOwnership = "native" | "surface";

const DOM_POINTER_INPUT_TARGETS = new WeakMap<EventTarget, InputTargetResolution>();

/** Binds a DOM renderer's declared terminal pointer owner to its mounted host. */
export function declareDomPointerOwnership(
  host: EventTarget,
  ownership: DomPointerOwnership,
  surface: SurfaceId
): void {
  if (ownership !== "native" && ownership !== "surface") {
    throw new RangeError(`Unsupported DOM pointer ownership: ${String(ownership)}`);
  }
  DOM_POINTER_INPUT_TARGETS.set(
    host,
    ownership === "native" ? nativeInputTarget() : surfaceInputTarget(surface)
  );
}

/** Returns only explicit renderer-host declarations; ordinary DOM keeps the caller's fallback policy. */
export function getDeclaredDomPointerInputTarget(
  entry: EventTarget
): InputTargetResolution | null {
  return DOM_POINTER_INPUT_TARGETS.get(entry) ?? null;
}
