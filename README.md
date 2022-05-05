# mozardbv/office-addin-react

Office web add-in voor Mozard

- [Meld een bevinding](https://intranet.mozard.nl/mozard/!suite09.scherm1089?mWfr=367)
- [Verzoek nieuwe functionaliteit](https://intranet.mozard.nl/mozard/!suite09.scherm1089?mWfr=604&mDdv=990842)

<!-- INHOUDSOPGAVE -->

## Inhoudsopgave

- [mozardbv/office-addin-react](#mozardbvoffice-addin-react)
  - [Inhoudsopgave](#inhoudsopgave)
  - [Over dit project](#over-dit-project)
  - [Aan de slag](#aan-de-slag)
    - [Afhankelijkheden](#afhankelijkheden)
    - [Installatie](#installatie)
  - [Tests draaien](#tests-draaien)
    - [Coding style tests](#coding-style-tests)
    - [Manifestvalidatie](#manifestvalidatie)
  - [Deployment](#deployment)
    - [Bouwen voor productie](#bouwen-voor-productie)
  - [Gebouwd met](#gebouwd-met)
  - [Bijdragen](#bijdragen)
  - [Versioning](#versioning)
  - [Auteurs](#auteurs)
  - [Licentie](#licentie)

## Over dit project

Zie ook: [Office add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)

## Aan de slag

### Afhankelijkheden

- `nodejs` >= 14.2
- Microsoft Office

### Installatie

Clone dit git repository:

```sh
git clone git@gitlab.com:MozardBV/office-addin-react.git
```

Haal lokale node modules binnen. Let op, dit werkt niet met `yarn`, aangezien deze niet goed omgaat met CRLF line endings. Zie [GitHub Issue](https://github.com/OfficeDev/Office-Addin-Scripts/issues/168)

```sh
npm install
```

Start de add-in:

```sh
npm run start
# of
npm run start:desktop
# of
npm run start:web
```

Het kan zijn dat door een [bug](https://github.com/OfficeDev/Office-Addin-Scripts/issues/330) de devserver niet automatisch start, is dat het geval, start deze dan met:

```sh
npm run dev-server
```

Om te stoppen:

```sh
npm run stop
```

Let op dat bij Windows op dat de current directory path exact hetzelfde is qua hoofdletters als het path van de modules. Anders geeft webpack een foutmelding. Bijvoorbeeld:

```zsh
C:\Users\user\Documents\office-addin-react\
# in plaats van
C:\users\user\documents\office-addin-react\
```

Je kan op meerdere manieren lokaal ontwikkelen:

- [Met de browser developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Met een debugger in de task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
  - [Inspector aanzetten](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/debug-office-add-ins-on-ipad-and-mac)
- [Met de F12 Developer Tools op Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

Omdat HTTPS een vereiste is, worden de React DevTools niet automatisch gestart. Dit moet je handmatig doen, en daarbij een bestaand x.509 keypair opgeven. Bijvoorbeeld:

```sh
KEY="/Users/patrick/.office-addin-dev-certs/localhost.key" CERT="/Users/patrick/.office-addin-dev-certs/localhost.crt" node node_modules/.bin/react-devtools
```

De dev server verzorgt _hot module reload_, maar wil je handmatig compileren:

```sh
npm run build:dev
```

## Tests draaien

### Coding style tests

- `npm run lint` voor alle linters.
- `npm run lint:fix` voor alle linters en automatisch herstellen (waar mogelijk).
- `npm run prettier` voor code formatting

### Manifestvalidatie

- `npm run validate`

## Deployment

### Bouwen voor productie

```sh
npm run build
```

## Gebouwd met

- [Axios](https://github.com/axios/axios) - HTTP client
- [Babel](https://babeljs.io) - ES6 transpiler
- [ESLint](https://eslint.org) - JavaScript linter
- [NodeJS](https://nodejs.org/en/) - JavaScript runtime
- [Office Addin Scripts](https://github.com/OfficeDev/Office-Addin-Scripts)
- [React](https://reactjs.org) - JavaScript framework
- [React Router](https://reactrouter.com) - Hash router
- [PostCSS](https://postcss.org) - CSS transformaties
- [Prettier](https://prettier.io) - Code formatter
- [Styled Components](https://styled-components.com)
- [StyleLint](https://stylelint.io) - CSS linter
- [Webpack](https://webpack.js.org) - bundler
- [UUID](https://github.com/uuidjs/uuid)

## Bijdragen

Zie [CONTRIBUTING.md](https://gitlab.com/mozardbv/office-addin-react/-/blob/main/CONTRIBUTING.md) voor de inhoudelijke procesafspraken.

## Versioning

Gebruikt [SemVer](https://semver.org/).

## Auteurs

- **Patrick Godschalk (Mozard)** - _Ontwikkelaar_ - [pgodschalk](https://gitlab.com/pgodschalk)

Zie ook de lijst van [contributors](https://gitlab.com/mozardbv/office-addin-react/-/graphs/main) die hebben bijgedragen aan dit project.

## Licentie

[SPDX](https://spdx.org/licenses/) license: `GPL-3.0-or-later`

Copyright (c) 2006-2021 Mozard B.V.

[Leveringsvoorwaarden](https://www.mozard.nl/mozard/!suite86.scherm0325?mPag=204&mLok=1)
