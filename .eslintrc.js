module.exports = {
  root: true,
  env: {
    node: true,
  },
  extends: [
    "eslint:recommended",
    "standard",
    "plugin:import/errors",
    "plugin:react/recommended",
    "plugin:jsx-a11y/recommended",
    "plugin:office-addins/recommended",
    "plugin:office-addins/react",
    "prettier",
    "prettier/react",
    "prettier/standard",
  ],
  plugins: ["react", "import", "jsx-a11y", "react-hooks", "@typescript-eslint", "office-addins"],
  parserOptions: {
    sourceType: "script",
  },
  rules: {
    "no-console": process.env.NODE_ENV === "production" ? "warn" : "off",
    "no-debugger": process.env.NODE_ENV === "production" ? "error" : "off",
    "react/prop-types": 0,
    "react-hooks/rules-of-hooks": 2,
    "react-hooks/exhaustive-deps": 1,
    "no-use-before-define": 0,
    "@typescript-eslint/no-use-before-define": 1,
  },
  overrides: [
    {
      files: ["src/**/*"],
      parser: "@typescript-eslint/parser",
      parserOptions: {
        ecmaVersion: 6,
        sourceType: "module",
        ecmaFeatures: {
          jsx: true,
        },
        project: "./tsconfig.json",
      },
      env: {
        browser: true,
      },
    },
  ],
  settings: {
    react: {
      version: "detect",
    },
  },
};
