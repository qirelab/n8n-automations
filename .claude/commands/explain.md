# Explain

This command breaks down complex topics into simple, understandable explanations.
It can explain code, architecture decisions, commits, technical concepts,
/or any development-related topic using mentor agents when available.

## Usage

```
/explain [code|concept|file|commit]
```


## Implementation Instructions

When this command is used:
1. **Identify the subject matter** - determine if it's code, architecture, concept, or commit
2. **Select appropriate agent** - use specialized mentor agent if suitable exists 
3. **Launch the agent** with a prompt to explain the topic in simple, educational terms
4. **Focus on clarity** - break complex topics into digestible parts with examples
5. **Tailor explanation** to user's apparent skill level and background

## How it works

1. **Analyzes** the topic you want explained
2. **Selects** appropriate mentor agent based on the subject matter
3. **Breaks down** complex concepts into simple, digestible parts
4. **Provides examples** and analogies to aid understanding
5. **Suggests** next steps for deeper learning

## Examples

```
/explain this async function
/explain why we use components here
/explain commit abc123f
```