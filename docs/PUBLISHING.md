# Marketplace publishing runbook

This repository is prepared for trusted publishing to the Visual Studio Marketplace. It does not publish merely because the workflow file exists.

## One-time setup

1. Create or confirm the Marketplace publisher ID `steph-tools`.
2. Confirm that extension name `excel-ai-vba-studio` and display name `Excel AI & VBA Studio` are available.
3. Create the public source repository.
4. Add its `repository`, `homepage`, and `bugs` URLs to `package.json`.
5. Configure a private security contact and public support destination.
6. In the Marketplace publisher management portal, create a trusted-publishing policy for the exact GitHub owner, repository, and workflow `.github/workflows/publish.yml`.
7. Protect release tags and require review for workflow changes.

The workflow uses GitHub Actions OIDC. Do not create a `VSCE_PAT` secret.

## Release

1. Start from a clean, reviewed default branch.
2. Run `npm ci`.
3. Run `npm run release:prepare -- 0.1.1` with the intended semantic version.
4. Update `CHANGELOG.md` and review privacy/notices.
5. Run `npm run validate`.
6. Commit the version and lockfile.
7. Create an annotated tag whose value exactly matches `v` plus `package.json` version:

   ```text
   git tag -a v0.1.1 -m "Excel AI & VBA Studio 0.1.1"
   git push origin main v0.1.1
   ```

8. GitHub Actions validates, builds, packages for `win32-x64`, obtains a short-lived Marketplace credential through OIDC, and publishes.

If the tag and manifest version differ, publication stops before `vsce publish`.

## Preview channel

The initial Marketplace listing uses `"preview": true`. It is a Preview-labelled regular Marketplace release, not a separate VS Code pre-release channel. If a pre-release channel is introduced later, use distinct `major.minor.patch` versions and the `--pre-release` flag according to VS Code Marketplace rules.
