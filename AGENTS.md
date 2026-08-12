# Guide for AI Agents

## 0. Project Overview

```txt
root: .
├── benchmarks/
├── build/
│   └── VBA-HTTP.xlsm
├── dist/
├── docs/
│   ├── adr/
│   ├── specs/
│   └── design.md
├── examples/
├── fixtures/
├── src/
│   ├── classes/
│   ├── forms/
│   ├── modules/
│   └── workbook/
├── tasks/
│   ├── lessons.md
│   └── todo.md
├── tools/
├── AGENTS.md
├── CHANGELOG.md
├── LICENSE
├── README.md
└── xlflow.toml
```

## 1. Workflow Design

### 1. Basic Approach: Work in Plan Mode First

- For tasks involving three or more steps or those affecting the overall architecture, always begin in Plan mode.
- If progress stalls at any point, do not force continuation - stop and replan instead.
- Use the Plan mode not only for implementation but also for designing verification procedures.
- As early as possible, refine specifications to reduce ambiguity.

### 2. Multi-Agent Strategy

- Make active use of subagents to avoid contaminating the main context.
- Delegate tasks such as research, verification, and parallel analysis to subagents.
- For complex problems, utilize subagents even when they require significant computational resources.
- Assign each subagent only one task to maintain focused execution.
- Use an explorer for codebase exploration (primarily reading activities).
- Use a worker for implementation and modifications.
- Use a reviewer for code reviews.

### 3. Self-Improvement Loop

- When receiving correction instructions from users, document these patterns in `tasks/lessons.md`.
- Formulate clear rules for yourself to prevent repeating the same mistakes.
- Continuously refine these rules until error rates decrease significantly.
- At the beginning of each session, review relevant lessons related to the project.

### 4. Always Verify Before Finalizing

- Do not mark tasks as complete until you can demonstrate their functionality.
- When necessary, compare your changes against the main branch for verification.
- Always ask yourself: "Would a staff engineer approve this?"
- Complete the process by running tests, reviewing logs, and demonstrating proper operation.

### 5. Maintain Balance While Pursuing More Elegant Solutions

- Before implementing major changes, pause to first consider: "Is there a more elegant way to do this?"
- If your fix feels ad hoc, reframe it as: "How can I implement this in a more refined manner based on what I know now?"
- However, do not overthink simple and obvious fixes - avoid excessive design.
- Before delivering any deliverable, thoroughly review your own implementation with a critical eye.

### 6. Handle Bug Fixes Autonomously

- When receiving bug reports, investigate them independently without waiting for instructions, then proceed directly to resolution.
- Use logs, errors, and failing tests to autonomously identify and resolve the issue.
- Avoid forcing users into unnecessary context switches.

## - Even without explicit instructions, if the CI pipeline is down, take initiative to resolve it.

## 2. Required Workflow Procedures

Before generating or modifying code, perform the following steps according to the scale of your work:

1. Understand the requirements: Review relevant specification documents, ADR documentation, and existing implementations.
2. Consider the design implications: Assess impact scope, compatibility with current designs, and alternative approaches.
3. If necessary, create working notes:

- For recurrence prevention: `tasks/lessons.md`

4. Add or update tests as needed.
5. Implement changes.
6. Verify functionality.
7. Run tests.
8. Conduct self-review.
9. Update documentation, ADR documents, specifications, and the CHANGELOG as appropriate.

- Any updates to ADR documents or specifications must be recorded in the respective directories:
- For ADR documents: `docs/adr/`
- For specification documents: `docs/specs/`
- If changes affect public APIs, they may require recording in the following documentation:
- Specification documents within `docs/specs/`
- Overview descriptions in the `README.md` file
- When making changes that impact users, append updates to the `CHANGELOG.md` file.

## 3. Documentation Retention Policy

### Role Separation Guidelines

- The `tasks/lessons.md` should exclusively serve as a repository for recording recurrence prevention rules - it must not be used for storing design decisions or actual specifications themselves.
- Design judgments and trade-offs should be recorded in the `docs/adr/` directory, while current internal specifications and constraints should be moved to the `docs/specs/` directory.

### Distinction Between ADR Documents and Specification Documentation

- ADR documentation should capture the reasoning behind decisions and document the rationale for choosing one approach over others—information that will be valuable for future implementers facing similar challenges.
- When editing ADR documents, use the `adr-manager` skill.
- Specification documents should record: permanent rules established through review processes, continuous integration testing, and failure resolution; as well as CLI specifications, validation requirements, and compatibility agreements.
- If additional regression tests were added due to specific design considerations where forgetting the rationale could lead to recurrence, document these in the specification documentation.

## - Simple procedure lists without accompanying reasoning justifications

## 4. Core Principles

- **Keep it simple first**: All changes should maintain maximum simplicity with minimal scope impact.
- **Do no harm**: Identify the root cause. Do not resort to quick fixes. Maintain professional engineering standards.
- **Minimize impact**: Only modify necessary components without introducing new bugs.

<!-- headroom:rtk-instructions -->

# RTK (Rust Token Killer) - Token-Optimized Commands

When running shell commands, **always prefix with `rtk`**. This reduces context
usage with zero behavior change. If rtk has no filter for a command, it passes
through unchanged — so it is always safe to use.

This project is developed on **Windows**, so prefer PowerShell-compatible
commands and paths.

## Key Commands

```powershell
# Git
rtk git status
rtk git diff
rtk git log --oneline -20

# Files & Search
rtk dir
rtk dir .\src
rtk read .\path\to\file.txt
rtk rg "pattern"
rtk rg "pattern" .\src
rtk find "pattern"
rtk diff .\path\to\file.txt

# Analysis
rtk err <command>
rtk log .\path\to\log.txt
rtk json .\path\to\file.json
rtk summary <command>
rtk deps
rtk env

# GitHub
rtk gh pr view <number>
rtk gh run list
rtk gh issue list
```

<!-- /headroom:rtk-instructions -->
