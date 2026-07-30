# AI4 Publication Development Research

The application is in
[`publication-research-automation`](publication-research-automation/README.md).
It is a standalone Windows and Node.js project; Codex is not required.

## Start here

- [Installation and operation](publication-research-automation/STEP_BY_STEP_GUIDE.md)
- [Guide for any AI coding assistant](publication-research-automation/AI_ASSISTANT_GUIDE.md)
- [Application README](publication-research-automation/README.md)

To use ChatGPT, Claude, Gemini, Copilot, or another coding assistant, give it the
repository and ask it to read `publication-research-automation/AI_ASSISTANT_GUIDE.md`
before making changes. That guide contains a ready-to-copy prompt, architecture,
behavioral requirements, validation commands, and runtime-provider guidance.

## New-computer setup

```powershell
git clone https://github.com/vijayvani0690/AI4-review-development.git
cd .\AI4-review-development\publication-research-automation
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

Node.js 20 or newer and Microsoft Edge are required. Do not copy
`node_modules`; `setup.ps1` installs the locked dependencies on each computer.
