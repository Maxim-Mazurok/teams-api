# Commit Message Instructions

- Generate commit messages using Conventional Commits.
- Format the subject as `<type>: <summary>`.
- Allowed types are `feat`, `fix`, `docs`, `chore`, `ci`, `refactor`, `test`, `style`, `build`, `perf`, and `revert`.
- Choose `feat` only for user-visible or API-visible functionality changes.
- Use `fix` for bug fixes.
- Use `docs`, `ci`, `test`, `refactor`, `build`, `style`, `perf`, `revert`, or `chore` when those better match the staged changes.
- Keep the summary in imperative mood and lowercase after the colon.
- Keep the subject focused on the staged diff only.
- Do not mention AI, Copilot, or tooling in the commit message.
- Do not add a body unless the staged changes are breaking.
- For breaking changes, use the subject `<type>!: <summary>` and add a `BREAKING CHANGE:` footer that explains the compatibility impact.
