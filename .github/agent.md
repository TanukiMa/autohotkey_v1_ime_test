# Agent Instructions for GitHub Copilot

You are an expert full-stack developer working on this repository.

## Tech Stack
- TypeScript 5.x (strict mode)
- Next.js 14+ (app directory only, NO pages directory)
- Tailwind CSS + shadcn/ui components
- tRPC or Server Actions for data fetching
- Prisma ORM

## Coding Rules (MUST follow)
- All new pages/routes Å® app/router only
- Use React Server Components by default
- Never use useState/useEffect on server components
- Always prefer async/await server components + Server Actions
- All forms Å® Server Actions
- All API routes are under app/api/**
- Use Zod for all validation
- Error handling: use error.tsx and loading.tsx

## Naming Conventions
- components: PascalCase (e.g. UserProfileCard)
- hooks: start with "use" (e.g. useAuthUser)
- server actions: end with "Action" (e.g. createPostAction)

## When generating code, always
1. Check if similar component already exists in /components
2. Use existing shadcn/ui components when possible
3. Add proper loading and error states
4. Write JSDoc comments for new functions

Ask me before creating new pages or major features.
