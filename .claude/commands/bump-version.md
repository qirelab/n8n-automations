# Bump version

This command bump version, changes CHANGELOG.md if exists, commit and push changes.

## Usage

```
/bump-version
```

## Process

1. Use one commone CHANGELOG.md file in the root for both Frontend and Backend folders/projects
1. Find current app version in package.json file (Frontend or Backend folder/project) in the "version" field or in CHANGELOG.md
2. Bump the app version number - new features should increase the minor version number, bug fixes should increase the patch version number. Breaking changes should increase the major version number. Follow [Semantic Versioning](https://semver.org/).
3. Update package.json files in Backend and Frontend folders/projects in the "version" field if exists
4. Update CHANGELOG.md if exists, ask user which changes that in not released section should be included in new version.
5. Commit the changes with a short, but descriptive message
6. Ask user approval to push the changes to the remote repository