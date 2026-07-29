// Flat config for the repo-root code that neither add-in's lint reaches.
//
// `npm run lint` in word/ and excel/ runs office-addin-lint, which is an
// ESLint 9 wrapper: flat config only, and ESLint 9 refuses to lint files
// outside the config's own directory. So shared/ (the code BOTH task panes
// import) and scripts/ were never linted by anything. This config covers them.

const browserGlobals = {
  fetch: "readonly",
  AbortController: "readonly",
  AbortSignal: "readonly",
  setTimeout: "readonly",
  clearTimeout: "readonly",
  document: "readonly",
  window: "readonly",
  console: "readonly",
};

const nodeGlobals = {
  process: "readonly",
  console: "readonly",
  Buffer: "readonly",
  URL: "readonly",
  setTimeout: "readonly",
  clearTimeout: "readonly",
};

export default [
  {
    // shared/ is imported into the Office WebView bundles, so it is browser
    // code even though its tests run in node.
    files: ["shared/**/*.js"],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: "module",
      globals: browserGlobals,
    },
    linterOptions: { reportUnusedDisableDirectives: true },
    rules: {
      "no-unused-vars": ["error", { argsIgnorePattern: "^_" }],
      "no-undef": "error",
      "no-console": "off",
      eqeqeq: ["warn", "smart"],
      "prefer-const": "warn",
    },
  },
  {
    // Test files additionally get node's test globals via imports, but they
    // also touch `global` to stub fetch.
    files: ["shared/**/*.test.js"],
    languageOptions: {
      globals: { ...browserGlobals, ...nodeGlobals, global: "writable" },
    },
  },
  {
    files: ["scripts/**/*.mjs"],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: "module",
      globals: nodeGlobals,
    },
    linterOptions: { reportUnusedDisableDirectives: true },
    rules: {
      "no-unused-vars": ["error", { argsIgnorePattern: "^_" }],
      "no-undef": "error",
      eqeqeq: ["warn", "smart"],
      "prefer-const": "warn",
    },
  },
];
