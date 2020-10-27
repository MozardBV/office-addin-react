// office-addin-react - Koppeling van Mozard met Microsoft Office
// Copyright (C) 2020  Mozard BV
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
  extends: [
    // Gebruik de Standard config als base
    // https://github.com/stylelint/stylelint-config-standard
    "stylelint-config-standard",
    // Standaardvolgorde van CSS properties
    // https://github.com/stormwarning/stylelint-config-recess-order
    "stylelint-config-recess-order",
    // Override rules die conflicteren met Prettier
    // https://github.com/shannonmoeller/stylelint-config-prettier
    "stylelint-config-prettier",
    // Override rules voor linten van CSS modules
    // https://github.com/pascalduez/stylelint-config-css-modules
    "stylelint-config-css-modules",
  ],
  // Rule lists:
  // - https://stylelint.io/user-guide/rules/
  rules: {
    /**
     * Possible errors
     */

    /* Font family */
    "font-family-no-missing-generic-family-keyword": null,

    /* String */
    "string-no-newline": null,

    /**
     * Limit language features
     */
    /* Alpha-value */
    "alpha-value-notation": "number",

    /* Hue */
    "hue-degree-notation": "number",

    /**
     * Stylistic issues
     */
    /* Color */
    "color-function-notation": "legacy",
    "color-named": "never",
    "color-no-hex": true,

    /* Font weight */
    "font-weight-notation": "numeric",

    /* Function */
    "function-url-scheme-disallowed-list": ["ftp", "http"],

    /* Number */
    "number-max-precision": 8,

    /* Time */
    "time-min-milliseconds": 100,

    /* Unit */
    "unit-disallowed-list": ["cm", "in", "mm", "pc", "pt"],

    /* Shorthand property */
    "shorthand-property-no-redundant-values": true,

    /* Value */
    "value-no-vendor-prefix": true,

    /* Property */
    "property-no-vendor-prefix": true,

    /* Declaration */
    "declaration-empty-line-before": "never",
    "declaration-no-important": true,

    /* Selector */
    "selector-max-attribute": 1,
    "selector-max-id": 1,
    "selector-max-universal": 1,
    "selector-no-vendor-prefix": true,

    /* Function */
    "function-url-quotes": "always",

    /* String */
    "string-quotes": "double",

    /* Declaration block */
    "declaration-block-semicolon-newline-after": "always",

    /* Selector */
    "selector-attribute-quotes": "always",

    /* Media query list */
    "media-query-list-comma-newline-after": "never-multi-line",
    "media-query-list-comma-newline-before": "never-multi-line",
    "media-query-list-comma-space-after": "always",

    /* At-rule */
    "at-rule-name-space-after": "always",
    "at-rule-semicolon-space-before": "never",

    /* General / Sheet */
    linebreaks: "unix",
    "unicode-bom": "never",
  },
};
