// office-addin-react - Koppeling van Mozard met Microsoft Office
// Copyright (C) 2021-2022  Mozard BV
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.

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
    //"plugin:office-addins/recommended",
    //"plugin:office-addins/react",
    "prettier",
  ],
  plugins: ["react", "import", "jsx-a11y", "react-hooks", "@typescript-eslint"], // "office-addins tijdelijk uit
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
