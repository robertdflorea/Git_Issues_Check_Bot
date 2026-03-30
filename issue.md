### 1.1.1 Issue Check rule.
Before claiming an issue:

- Always check to see if the codebase is valid. Here are some requirements:
    - **Python, Javascript or Typescript** and hosted in a Git repository
    - **≤ 200 MB** at the selected commit
    - **Understandable** enough for you to review PRs confidently
    - **Executable**, with clear build and run instructions
    - **Dependency-complete**, using standard package managers (pip/conda/npm)
    - **Tests** present
    - Written entirely in **English**
    - Other quality signals (recommended, not mandatory):
        - Linters or formatters configured
        - Actively used open-source projects (especially for applications)
- Get familiar with the repository and feel free to use AI to learn more about its concepts and how it's structured.
- Validate that the issue is **non-trivial** and will require multiple iterations.

<aside>
⚠️

If the model can fully solve the issue in 1–2 interactions, the task is too easy. Choose a more complex issue.

</aside>

### 1.1.2 What a Good Issue Looks Like

✅ **Well-scoped examples**

- Add async/await support to existing callback-based APIs while preserving backward compatibility
- Introduce a `-dry-run` flag that reports changes without executing them
- Refactor validation logic to use a schema-based approach

❌ **Poorly-scoped examples**

- Build a full admin dashboard (too broad)
- Rename a variable as a standalone task (too trivial)
- "Add some validation" (vague)
- Issues described only by a link