# NFA Web Part Components

This directory contains the main functional components for the NFA web part.

## Structure
- `AdminSection`: Components for administrative tasks (DetailView, Masters).
- `ApprovalNotes`: Core logic for managing NFAs.
  - `Inbox`: User's pending approvals.
  - `ApprovedRejectedByMe`: History of actions taken.
  - `Home`: Dashboard/Landing page for NFAs.
  - `NewNFA`: Form for creating new NFAs.
- `Reports`: Various reporting components (Ageing, Pendency, Status).
- `Layouts`: Header and other layout elements.
- `SideBar`: Navigation sidebar.

## Development Guidelines
- Use hooks for state and side effects.
- Prefer `NFAService` or `NFADocService` for data operations.
- Ensure proper typing using interfaces defined in `INfaProps.ts` or local interfaces.
