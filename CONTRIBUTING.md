# Contribuer / Contributing

Merci de contribuer à Excel AI & VBA Studio. Les échanges peuvent être rédigés en français ou en anglais.

## Avant de commencer

- Recherchez les [issues](https://github.com/StephaneSGL/excel-ai-vba-studio/issues) et [discussions](https://github.com/StephaneSGL/excel-ai-vba-studio/discussions) existantes.
- Utilisez une discussion pour une question générale et une issue pour un bug reproductible ou une proposition précise.
- Signalez toute vulnérabilité uniquement via le [formulaire privé GitHub](https://github.com/StephaneSGL/excel-ai-vba-studio/security/advisories/new), conformément à [SECURITY.md](SECURITY.md).
- Ne partagez jamais de classeur réel, de données confidentielles, d'identifiants ou de code VBA propriétaire. Créez un exemple synthétique minimal.

## Développement local

Prérequis : Node.js 22, npm, Visual Studio Code 1.95 ou ultérieur. Windows x64 et Microsoft Excel desktop sont nécessaires pour tester l'intégration native Excel, COM et VBA.

1. Forkez puis clonez le dépôt.
2. Créez une branche courte depuis `main`.
3. Installez exactement les dépendances verrouillées avec `npm ci`.
4. Développez et ajoutez les tests ou validations pertinents.
5. Exécutez `npm run validate` avant d'ouvrir une pull request.

Les chemins d'export, l'automatisation Excel, les macros et les formats `.xlsm`, `.xls` et `.xlsb` sont sensibles. Toute modification de ces zones doit préserver les garde-fous de sécurité, ne jamais exécuter de macro pendant l'analyse et utiliser uniquement des fichiers de test synthétiques.

## Pull requests

- Gardez chaque pull request ciblée et expliquez le problème résolu.
- Décrivez les tests manuels et automatiques effectués.
- Mettez à jour la documentation et `CHANGELOG.md` lorsque le comportement visible change.
- Préservez les mentions de licence et d'attribution du projet d'origine.

En proposant une contribution, vous acceptez qu'elle soit distribuée sous la [licence MIT](LICENSE) du projet.

---

Contributions are welcome in French or English. Please keep pull requests focused, use synthetic test workbooks only, run `npm run validate`, document user-visible changes, and report security issues privately. Native Excel/COM/VBA testing requires Windows x64 with Microsoft Excel desktop installed.
