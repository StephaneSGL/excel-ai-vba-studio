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

The current workflow builds verified artifacts only. Marketplace OIDC publication is not wired up yet. Do not create a `VSCE_PAT` secret as a workaround.

## Release

1. Start from a clean, reviewed default branch.
2. Run `npm ci`.
3. Run `npm run release:prepare -- 0.1.1` with the intended semantic version.
4. Update `CHANGELOG.md` and review privacy/notices.
5. Confirm that `LICENSE`, `LICENSING.md`, the `LICENSES` directory, and all
   third-party notices match the files included in the VSIX.
6. Run `npm run validate`, then `npm run test:extension-host -- 1.95.0` and `npm run test:extension-host -- stable`.
   Before calling a release fully validated, run `npm run validate:native-release` on a dedicated Windows/Excel test machine with the required user-approved VBA access. This strict gate fails when native tests would otherwise skip. Never change Office security settings automatically to make a test pass.
7. Commit the version and lockfile.
8. Create an annotated tag whose value exactly matches `v` plus `package.json` version:

   ```text
   git tag -a v0.1.1 -m "Excel AI & VBA Studio 0.1.1"
   git push origin main v0.1.1
   ```

9. GitHub Actions validates, rebuilds the native helper, packages for `win32-x64`, and uploads the verified VSIX as a 30-day workflow artifact.
10. After the tagged workflow succeeds, download that exact artifact, verify its SHA-256, and attach the VSIX to a GitHub Release for the same tag.

If the tag and manifest version differ, packaging stops before the VSIX is created.

Automatic Visual Studio Marketplace publication is intentionally disabled until a Microsoft Entra workload identity and trusted-publishing policy are configured for the publisher. A successful tagged workflow proves the package was built and validated; it does not prove Marketplace publication.

A candidate VSIX may be prepared for review without native acceptance, but must be labelled as such. Do not promote it as fully validated or remove the Preview flag while the native gate remains blocked. The 0.6.1 audit and remaining acceptance checklist are in `docs/RELEASE_AUDIT_0.6.1.md`.

## Preview channel

The initial Marketplace listing uses `"preview": true`. It is a Preview-labelled regular Marketplace release, not a separate VS Code pre-release channel. If a pre-release channel is introduced later, use distinct `major.minor.patch` versions and the `--pre-release` flag according to VS Code Marketplace rules.
