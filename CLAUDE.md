# excel-recon-addin

## Auto-commit and push

After every code change, without asking and without waiting to be reminded:

1. **Commit** — `git add -A && git commit` with a clear message you choose.
2. **Push** — `git push`.

If no remote repository exists, ask the user whether to create one.

Commit messages must **not** include any `Co-Authored-By: Claude` / Anthropic
trailer or any other attribution to Claude or Anthropic.

A change that is not committed and pushed is **incomplete**. After a successful
commit and push, respond with only: **commited, pushed.**
