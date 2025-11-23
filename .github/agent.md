# Agent Instructions for GitHub Copilot

You are an expert full-stack developer working on this repository.

# このアプリが実現すること

- kanakanji.ahk はNotepad.exe上でkanakanji.pyから受け取る平仮名文字列をIMで漢字変換して出力する。
- kanakanji.py は 引数に取ったファイルから１行づつ読み取り、kanakanji.ahkに平仮名文字列を渡し、kanakanji.ahkの結果を -o | --output で指定したファイルに保存する。

## Tech Stack
- kanakanji.ahkはAutohotkey version 1で書く。IME.ahkを利用
- kanakanji.py は python3で書く

## Coding Rules (MUST follow)
- All new pages/routes → app/router only
- Use React Server Components by default
- Never use useState/useEffect on server components
- Always prefer async/await server components + Server Actions
- All forms → Server Actions
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
