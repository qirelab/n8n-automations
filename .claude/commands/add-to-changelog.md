# Update Changelog

This command adds a new entry to the project's CHANGELOG.md file. If no params provided - determine changes and version automatically.

## Usage

```
/add-to-changelog <version> <change_type> <message>
```

Where:
- `<version>` (*optional*) is the version number (e.g., "1.1.0")
- `<change_type>` (*optional*) is one of: "added", "changed", "deprecated", "removed", "fixed", "security"
- `<message>` (*optional*) is the description of the change

## Examples

```
/add-to-changelog 1.1.0 added "New end point for pickups count"
```

```
/add-to-changelog 1.0.2 fixed "Bug in the response structure of the products list endpoint"
```

```
/add-to-changelog
```

## Description

This command will:

1. Check if a CHANGELOG.md file exists and create one if needed
2. Look for an existing section for the specified version
    - If found, add the new entry under the appropriate change type section
    - If not found, create a new version section with today's date
3. Format the entry according to Keep a Changelog conventions
4. If no version provided ask user add changelog without version(as not released) or determine version according to [Semantic Versioning](https://semver.org/spec/v2.0.0.html)
5. If no info about changes provided determine changes by looking at changes(commited and not) after last changelog entry.
6. Version in the CHANGELOG.md file should be the same as in the package.json file
7. Commit the changes with a short, but descriptive message

The CHANGELOG.md follows the [Keep a Changelog](https://keepachangelog.com/) format and [Semantic Versioning](https://semver.org/).

## Implementation

The command should:

1. Parse the arguments to extract version, change type, and message
2. Read the existing CHANGELOG.md file if it exists
3. If the file doesn't exist, create a new one with standard header
4. Check if the version section already exists
5. Add the new entry in the appropriate section
6. Write the updated content back to the file
7. Synchronize version in the package.json file with the version in the CHANGELOG.md file
8. Commit the changes with a short, but descriptive message
9. Ask user approval to push the changes to the remote repository
