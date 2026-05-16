---
description: Generate optimized Linear.app tasks
argument-hint: [task outline or feature description]
---

# Linear Task Generator

You will create well-structured Linear tasks from the user's input. Analyze the input to determine task type and complexity, then generate an appropriate task.
Use templates from .claude/templates/task-creation.md to generate the task. Please follow the same structure and format as in the provided template.
In the folder .claude/themes/dawn_theme there is old theme structure. In the folder .claude/themes/horizon_theme there is new theme structure. Old theme structure is currently supported in the project. All tasks will be related to add support to some elements of the new theme.

## Task Analysis
1. **Determine task type**: Feature, Bug Fix, Tech Debt, Research, or Spike
2. **Assess complexity**: Simple (< 1 hours), Medium (2-5 hours), Complex (> 5 hours)
3. **Identify scope boundaries**: What's included vs. excluded


## Guidelines
- **Be specific**: Use short, concrete and precise requirements, not vague descriptions
- **Include examples**: Show expected data formats or API responses
- **Define boundaries**: Explicitly state what's NOT included
- **Make it testable**: Include verification steps
- **Consider the developer**: Assume they're competent but need clear requirements

**Language**: Use American English, intermediate level
**Tone**: Professional and direct

Generate the task now based on: $ARGUMENTS
If arguments contains a target path, write result to target path `linear-links-output_$TOPIC_$TIMESTAMP.md` otherwise to `.claude/tmp`.