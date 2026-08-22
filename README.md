# React + Vite

This template provides a minimal setup to get React working in Vite with HMR and some Oxlint rules.

Currently, two official plugins are available:

- [@vitejs/plugin-react](https://github.com/vitejs/vite-plugin-react/blob/main/packages/plugin-react) uses [Oxc](https://oxc.rs)
- [@vitejs/plugin-react-swc](https://github.com/vitejs/vite-plugin-react/blob/main/packages/plugin-react-swc) uses [SWC](https://swc.rs/)

## React Compiler

The React Compiler is not enabled on this template because of its impact on dev & build performances. To add it, see [this documentation](https://react.dev/learn/react-compiler/installation).

## Expanding the Oxlint configuration

If you are developing a production application, we recommend using TypeScript with type-aware lint rules enabled. Check out the [TS template](https://github.com/vitejs/vite/tree/main/packages/create-vite/template-react-ts) for information on how to integrate TypeScript and Oxlint's TypeScript related rules in your project.

## Login and access roles

The dashboard uses signed, eight-hour login sessions and enforces access roles on both the UI and API.

- `viewer`: dashboard and report export only
- `staff`: viewer access plus IT Ticket, equipment request, and return request
- `admin`: full data management, approval, close-work, import, and database synchronization

Configure production accounts with the secret environment variable `AUTH_USERS_JSON` and set a strong `AUTH_TOKEN_SECRET`. Each user entry supports `username`, `name`, `role`, and either `password` or `passwordHash`.
