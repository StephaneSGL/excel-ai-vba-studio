# Dépôt canonique et prototypes

## Source unique

Le dépôt de production est
[`StephaneSGL/excel-ai-vba-studio`](https://github.com/StephaneSGL/excel-ai-vba-studio).
`package.json` est l’unique source de version. Les paquets distribuables sont
construits depuis ce dépôt et générés sous `output/vsix`.

L’exécutable `bin/win32-x64/excel-ai-vba-writeback.exe` est un composant
d’exécution livré avec l’extension. Sa source maintenue se trouve sous
`native/vba-writeback`; la CI Windows reconstruit et teste cette source.

## Décisions sur les prototypes

| Prototype/capacité | Décision | Destination ou suivi |
| --- | --- | --- |
| Édition XLSM ciblée via Excel COM | Retenue et portée | `scripts/office-ai-apply-edits.ps1` et `src/provider/nativeExcelBridge.ts` |
| Tests de préservation XLSM/VBA/UserForm | Retenus et portés | `test/run-native-edit-tests.ps1` et `test/run-vba-preservation-test.ps1` |
| Réinjection directe du code VBA | Retenue | `native/vba-writeback` et VBA Studio |
| Hôte C# Office Workbench autonome | Non retenu | Le chemin PowerShell/COM couvre le périmètre borné actuel sans second hôte de production |
| Création de UserForms, contrôles et boutons | Différée | [Issue #7](https://github.com/StephaneSGL/excel-ai-vba-studio/issues/7) |
| Dimensions, fusions, structure et règles conditionnelles éditables | Différées | [Issue #10](https://github.com/StephaneSGL/excel-ai-vba-studio/issues/10) |

Les anciens prototypes sont des archives ou références techniques, jamais des
sources de release ni des extensions concurrentes. Aucun remote de prototype
autonome ne doit pousser une version de production. Les branches historiques
ne sont supprimées qu’après fusion et validation de la consolidation.

## Garde-fous de contribution

1. Partir de `main` du dépôt canonique.
2. Porter une capacité bornée avec ses tests et sa décision d’architecture.
3. Ne jamais suivre de classeur personnel, secret ou artefact temporaire.
4. Exécuter `npm run validate`; pour une modification COM, exécuter aussi la
   suite Excel Desktop.
5. Faire relire la PR et attendre ses checks avant fusion.
6. Construire la release depuis la version de `package.json`, jamais depuis un
   prototype ou un chemin local.

La décision de sécurité complète est décrite dans
[`docs/VBA-WRITEBACK-ADR.md`](VBA-WRITEBACK-ADR.md).
