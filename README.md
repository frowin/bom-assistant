# BOM-Assistant (Excel Add-In)
## An Excel Add-In with a descriptive name due to lack of imagination.

This is a simple Excel add-in that allows you to extract a bill of materials (BOM) from an Excel sheet. It is highly specialized for PCB manufacturing.


## Stack

- React
- TypeScript
- Office JS API
- Webpack
- Babel
- ESLint

## Features

- Extract a list of components from an Excel sheet in JSON format
- Import component attributes from JSON to add/or update existing components
- Taylored to a tool called `pcbbuddy` which can help ordering PCBs and consumes JSON and fetches component data as stock, price, availability.

## Installation

```bash
npm install
npm run dev
```

## Build

```bash
npm run build
```

## Deploy

- create GitHub Pages and Actions for deployment.
