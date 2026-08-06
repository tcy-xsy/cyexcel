import { defineConfig, devices } from "@playwright/test";

export default defineConfig({
  testDir: "./tests",
  // Playwright's own default testMatch also picks up "*.test.ts" — harmless
  // everywhere else in the repo (Playwright is only ever pointed at
  // tests/browser or tests/e2e, which have no .test.ts files), but
  // tests/build/ mixes real Vitest *.test.ts files alongside Playwright
  // *.spec.ts fixtures in the same directory, so the collision needs an
  // explicit fix rather than continuing to rely on directory separation.
  testMatch: "**/*.spec.ts",
  fullyParallel: true,
  forbidOnly: Boolean(process.env.CI),
  retries: process.env.CI ? 1 : 0,
  reporter: process.env.CI ? "github" : "list",
  use: {
    baseURL: "http://127.0.0.1:4173",
    trace: "retain-on-failure"
  },
  projects: [
    {
      name: "chrome",
      use: { ...devices["Desktop Chrome"], channel: "chrome" }
    },
    { name: "chromium", use: { ...devices["Desktop Chrome"] } },
    { name: "firefox", use: { ...devices["Desktop Firefox"] } },
    { name: "webkit", use: { ...devices["Desktop Safari"] } }
  ],
  webServer: {
    command: "pnpm exec vite --host 127.0.0.1 --port 4173",
    url: "http://127.0.0.1:4173/tests/browser/fixture.html",
    reuseExistingServer: !process.env.CI
  }
});
