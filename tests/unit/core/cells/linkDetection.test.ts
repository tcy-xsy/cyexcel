import { describe, expect, it } from "vitest";

import { detectLinkUrl } from "../../../../src/core/cells/linkDetection";

describe("detectLinkUrl", () => {
  it("recognizes a bare https URL", () => {
    expect(detectLinkUrl("https://example.com/path")).toBe("https://example.com/path");
  });

  it("recognizes a bare http URL", () => {
    expect(detectLinkUrl("http://example.com")).toBe("http://example.com");
  });

  it("prefixes a www.-only address with https://", () => {
    expect(detectLinkUrl("www.example.com")).toBe("https://www.example.com");
  });

  it("prefixes an email address with mailto:", () => {
    expect(detectLinkUrl("someone@example.com")).toBe("mailto:someone@example.com");
  });

  it("ignores surrounding whitespace", () => {
    expect(detectLinkUrl("  https://example.com  ")).toBe("https://example.com");
  });

  it("does not match a URL embedded in a longer sentence — the whole value must be the link", () => {
    expect(detectLinkUrl("see https://example.com for details")).toBeUndefined();
  });

  it("does not match ordinary text", () => {
    expect(detectLinkUrl("工业温度传感器")).toBeUndefined();
  });

  it("does not match a plain number-looking string", () => {
    expect(detectLinkUrl("123.456")).toBeUndefined();
  });

  it("does not match the empty string", () => {
    expect(detectLinkUrl("")).toBeUndefined();
    expect(detectLinkUrl("   ")).toBeUndefined();
  });
});
