# Project: NFA (Note for Approval)

This is a SharePoint Framework (SPFx) project for a "Note for Approval" (NFA) system.

## Tech Stack
- **Framework:** SPFx 1.22.2
- **UI Library:** React 17, Fluent UI React (^8.106.4), PrimeReact (^10.9.7)
- **Styling:** Tailwind CSS (^3.4.19), Sass (.scss)
- **State Management:** Functional components with Hooks
- **Data Access:** @pnp/sp (^4.19.0)
- **Build Tool:** Heft

## Architecture & Conventions
- **Components:** Functional components are preferred. Use Fluent UI for standard SharePoint look and feel, and PrimeReact for specialized components (like Tables if applicable, though `CommonTable` and `MasterTable` are present in `src/Share/Components`).
- **Styling:** Use Tailwind CSS for utility-first styling. Custom styles should go into `.scss` files. Note that Tailwind is integrated via a custom build step (`tw:build`).
- **Services:** API and data logic should be encapsulated in service classes located in `src/Share/Services`.
- **Naming:** Follow standard SPFx/React naming conventions (PascalCase for components, camelCase for variables/functions).

## Common Commands
- `npm run start`: Start the local workbench (via Heft).
- `npm run build`: Production build.
- `npm run tw:dev`: Watch Tailwind CSS changes.
- `npm run tw:build`: Build Tailwind CSS.

## Directory Structure
- `src/Share/Components`: Reusable UI components.
- `src/Share/Services`: Data services (NFAService, DropdownService, etc.).
- `src/webparts/nfa`: Main web part logic.
- `src/webparts/nfa/components`: Web part specific components (Home, Inbox, NewNFA, etc.). See [src/webparts/nfa/components/GEMINI.md](./src/webparts/nfa/components/GEMINI.md) for more details.
