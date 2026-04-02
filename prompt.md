I need to find valid issues(closed) that satisfies several requirements.

Issues requirements are as follows:
"### 1.1.1 Issue Check rule.
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
- Issues described only by a link"

First give me your plan.
And once I agree then proceed with it.


6a1a45de5b1fd46aa40bd0b7629085bf4b1e34c1



The requirement for valid issues are as follows : 
"REPOSITORY & ISSUE REQUIREMENTS (UPDATED)

==========================================





REPOSITORY REQUIREMENTS

------------------------



Required:

- Language: Python, JavaScript, or TypeScript

- Hosted in Git

- Size: 200 MB or less at selected commit

- Must be understandable enough for confident review

- Must have clear build and run instructions

- Must use standard dependency tooling

- Must include tests

- Must be written in English



Optional (helpful):

- Linters configured

- Formatters configured

- Evidence of active usage or maintenance





ISSUE REQUIREMENTS

------------------



General Rules:

- Must be non-trivial

- Must require multiple iterations

- Must be one clear problem

- Must be medium or high complexity

- Must fit in a single engineering ticket

- Must naturally require new tests as part of the fix (avoid pure refactors or doc-only changes)



Complexity Gate (CRITICAL):

- Fix must require MORE than 20 meaningful lines of code changes

- Issue must naturally need at least 3 significant coding interactions

- Significant means actual code or test logic changes; not commits, cleanup, formatting, or moving files

- If the original PR that fixed the issue changed fewer than 20 lines of real logic (excluding comments and test-only lines), the issue is too simple

- If the model delivers a complete fix in 1 or 2 interactions, the issue is too simple; stop and pick a new one

- Do not force extra interactions by asking only for tests or only for a commit; every interaction must involve real code changes

- Asking only for tests in one round does not count as a significant interaction, even if bundled with a commit request

- Moving tests from one file to another does not count as a significant interaction



Good Issue Examples:

- Add async/await support while preserving backward compatibility

- Add a dry-run mode that reports actions without applying them

- Refactor validation logic to a schema-based approach

- Fix a rule logic bug that requires handling multiple edge cases and adding targeted tests



Bad Issue Examples:

- Build a full admin dashboard

- Rename one variable as a standalone task

- Add some validation

- Issue described only by a link

- Issues whose original PR fix is under 20 lines of logic

- Issues that are pure documentation or CSS-only changes

- Issues that are internal refactors with no need for new tests

- Issues resolved with a tiny patch (e.g., 4 to 8 lines) even if tests are added afterward"

Based on this requirement, update issue check rules.
And then make all cells content in the table editable and selectable.
And for file changed cell, now it shows file routes. I want to make it show changed files number, meaningful lines number, and the files' name in one line with out line change (Enter command). 
And give each cell min width so that it's whol e content can be shown. 
