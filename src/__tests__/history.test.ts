import { describe, expect, it } from "vitest";
import { buildHistory, type HistoryMessage } from "../utils/history";

const msg = (role: "user" | "assistant", text: string, promptText?: string): HistoryMessage => ({ role, text, promptText, status: "done" });

describe("buildHistory", () => {
  it("pairs user/assistant turns and appends the new prompt", () => {
    const h = buildHistory([msg("user", "hi", "hi + ctx"), msg("assistant", "hello")], "next");
    expect(h).toEqual([
      { role: "user", text: "hi + ctx" },
      { role: "model", text: "hello" },
      { role: "user", text: "next" },
    ]);
  });

  it("skips failed answers", () => {
    const h = buildHistory([msg("user", "q"), { role: "assistant", text: "boom", status: "error" }], "again");
    expect(h).toEqual([{ role: "user", text: "again" }]);
  });

  it("drops bulky document context from older turns only", () => {
    const list: HistoryMessage[] = [];
    for (let i = 0; i < 4; i++) list.push(msg("user", `q${i}`, `q${i}\n\n[Selected text]\n"""\nlong\n"""`), msg("assistant", `a${i}`));
    const h = buildHistory(list, "now");
    expect(h[0].text).toBe("q0\n\n(document context omitted)");
    expect(h[6].text).toContain("[Selected text]");
  });
});
