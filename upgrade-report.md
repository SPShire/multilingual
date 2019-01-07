# Upgrade project C:\code\Shire\Multilingual.SPFxExt to v1.7.0

Date: 2018-11-12

## Findings

Following is the list of steps required to upgrade your project to SharePoint Framework version 1.7.0. [Summary](#Summary) of the modifications is included at the end of the report.

### FN001001 @microsoft/sp-core-library | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-core-library

Execute the following command:

```sh
npm i @microsoft/sp-core-library@1.7.0 -SE
```

File: [./package.json](./package.json)

### FN001008 react | Required

Upgrade SharePoint Framework dependency package react

Execute the following command:

```sh
npm i react@16.3.2 -SE
```

File: [./package.json](./package.json)

### FN001009 react-dom | Required

Upgrade SharePoint Framework dependency package react-dom

Execute the following command:

```sh
npm i react-dom@16.3.2 -SE
```

File: [./package.json](./package.json)

### FN001011 @microsoft/sp-dialog | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-dialog

Execute the following command:

```sh
npm i @microsoft/sp-dialog@1.7.0 -SE
```

File: [./package.json](./package.json)

### FN001012 @microsoft/sp-application-base | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-application-base

Execute the following command:

```sh
npm i @microsoft/sp-application-base@1.7.0 -SE
```

File: [./package.json](./package.json)

### FN001013 @microsoft/decorators | Required

Upgrade SharePoint Framework dependency package @microsoft/decorators

Execute the following command:

```sh
npm i @microsoft/decorators@1.7.0 -SE
```

File: [./package.json](./package.json)

### FN002001 @microsoft/sp-build-web | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-build-web

Execute the following command:

```sh
npm i @microsoft/sp-build-web@1.7.0 -DE
```

File: [./package.json](./package.json)

### FN002002 @microsoft/sp-module-interfaces | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-module-interfaces

Execute the following command:

```sh
npm i @microsoft/sp-module-interfaces@1.7.0 -DE
```

File: [./package.json](./package.json)

### FN002003 @microsoft/sp-webpart-workbench | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-webpart-workbench

Execute the following command:

```sh
npm i @microsoft/sp-webpart-workbench@1.7.0 -DE
```

File: [./package.json](./package.json)

### FN002008 tslint-microsoft-contrib | Required

Remove SharePoint Framework dev dependency package tslint-microsoft-contrib

Execute the following command:

```sh
npm un tslint-microsoft-contrib -D
```

File: [./package.json](./package.json)

### FN006003 package-solution.json isDomainIsolated | Required

Update package-solution.json isDomainIsolated

In file [./config/package-solution.json](./config/package-solution.json) update the code as follows:

```json
{
  "solution": {
    "isDomainIsolated": false
  }
}
```

File: [./config/package-solution.json](./config/package-solution.json)

### FN010001 .yo-rc.json version | Recommended

Update version in .yo-rc.json

In file [./.yo-rc.json](./.yo-rc.json) update the code as follows:

```json
{
  "@microsoft/generator-sharepoint": {
    "version": "1.7.0"
  }
}
```

File: [./.yo-rc.json](./.yo-rc.json)

### FN010007 .yo-rc.json isDomainIsolated | Recommended

Update isDomainIsolated in .yo-rc.json

In file [./.yo-rc.json](./.yo-rc.json) update the code as follows:

```json
{
  "@microsoft/generator-sharepoint": {
    "isDomainIsolated": false
  }
}
```

File: [./.yo-rc.json](./.yo-rc.json)

### FN019001 tslint.json rulesDirectory | Required

Remove rulesDirectory from tslint.json

In file [./tslint.json](./tslint.json) update the code as follows:

```json
{
  "rulesDirectory": []
}
```

File: [./tslint.json](./tslint.json)

### FN019002 tslint.json extends | Required

Update tslint.json extends property

In file [./tslint.json](./tslint.json) update the code as follows:

```json
{
  "extends": "@microsoft/sp-tslint-rules/base-tslint.json"
}
```

File: [./tslint.json](./tslint.json)

### FN017001 Run npm dedupe | Optional

If, after upgrading npm packages, when building the project you have errors similar to: "error TS2345: Argument of type 'SPHttpClientConfiguration' is not assignable to parameter of type 'SPHttpClientConfiguration'", try running 'npm dedupe' to cleanup npm packages.

Execute the following command:

```sh
npm dedupe
```

File: [./package.json](./package.json)

## Summary

### Execute script

```sh
npm i @microsoft/sp-core-library@1.7.0 react@16.3.2 react-dom@16.3.2 @microsoft/sp-dialog@1.7.0 @microsoft/sp-application-base@1.7.0 @microsoft/decorators@1.7.0 -SE
npm i @microsoft/sp-build-web@1.7.0 @microsoft/sp-module-interfaces@1.7.0 @microsoft/sp-webpart-workbench@1.7.0 -DE
npm un tslint-microsoft-contrib -D
npm dedupe
```

### Modify files

#### [./config/package-solution.json](./config/package-solution.json)

Update package-solution.json isDomainIsolated:

```json
{
  "solution": {
    "isDomainIsolated": false
  }
}
```

#### [./.yo-rc.json](./.yo-rc.json)

Update version in .yo-rc.json:

```json
{
  "@microsoft/generator-sharepoint": {
    "version": "1.7.0"
  }
}
```

Update isDomainIsolated in .yo-rc.json:

```json
{
  "@microsoft/generator-sharepoint": {
    "isDomainIsolated": false
  }
}
```

#### [./tslint.json](./tslint.json)

Remove rulesDirectory from tslint.json:

```json
{
  "rulesDirectory": []
}
```

Update tslint.json extends property:

```json
{
  "extends": "@microsoft/sp-tslint-rules/base-tslint.json"
}
```
