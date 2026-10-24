import { escapeHtml } from '../common/extensionResource';
import { buildSetupChecks, type SetupSnapshot } from './setupDiagnostics';

export function renderStartPage(snapshot: SetupSnapshot | undefined, nonce: string): string {
	if (!/^[a-zA-Z0-9]+$/.test(nonce)) throw new Error('Invalid webview nonce');
	const checks = snapshot ? buildSetupChecks(snapshot).map(check => `
		<li class="check"><div class="check-heading"><strong>${escapeHtml(check.label)}</strong>
		<span class="badge ${check.state}">${escapeHtml(check.status)}</span></div><p>${escapeHtml(check.detail)}</p></li>`).join('')
		: '<li class="check">Lecture des prérequis locaux… Aucun classeur ne sera ouvert.</li>';
	return `<!DOCTYPE html>
<html lang="fr"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'nonce-${nonce}'; script-src 'nonce-${nonce}'; base-uri 'none'; form-action 'none'">
<title>Bienvenue · Excel AI &amp; VBA Studio</title>
<style nonce="${nonce}">
:root{color-scheme:light dark}*{box-sizing:border-box}body{margin:0;background:var(--vscode-editor-background,#171b1a);color:var(--vscode-editor-foreground,#e3ebe6);font:14px/1.6 var(--vscode-font-family,system-ui,sans-serif)}
main{max-width:1100px;margin:0 auto;padding:38px 32px 28px}.topline,.brand,.check-heading,.section-heading,.actions{display:flex;align-items:center;gap:12px}.topline,.section-heading,.check-heading{justify-content:space-between}.brand{font-weight:650;letter-spacing:.02em}.mark{display:grid;place-items:center;width:34px;height:34px;border-radius:9px;background:#107c41;color:#fff;font-size:21px}.preview{border:1px solid var(--vscode-panel-border,#44534b);border-radius:99px;padding:3px 12px;font-size:12px;color:var(--vscode-descriptionForeground,#a7b8ad)}
.hero{position:relative;overflow:hidden;margin:28px 0;padding:38px;border:1px solid var(--vscode-panel-border,#3c5145);border-radius:18px;background:linear-gradient(125deg,#107c4120,transparent 75%)}.eyebrow{margin:0 0 12px;font-size:11px;letter-spacing:.16em;text-transform:uppercase;font-weight:700;color:var(--vscode-descriptionForeground,#a7b8ad)}h1{font-size:clamp(26px,4vw,40px);line-height:1.16;letter-spacing:-.035em;margin:0 0 16px;max-width:760px}h2{font-size:19px;letter-spacing:-.02em;margin:0}h3{font-size:16px;margin:12px 0 8px}p{margin:0 0 12px}.intro{max-width:700px;color:var(--vscode-descriptionForeground,#a7b8ad);font-size:15px}.actions{flex-wrap:wrap;margin-top:22px}
button{cursor:pointer;border:1px solid transparent;border-radius:6px;background:var(--vscode-button-background,#137f46);color:var(--vscode-button-foreground,#fff);font:inherit;font-size:13px;font-weight:600;padding:10px 17px}button:hover{background:var(--vscode-button-hoverBackground,#169d55)}button.secondary{background:var(--vscode-button-secondaryBackground,#303b35);color:var(--vscode-button-secondaryForeground,#e3ebe6);border-color:var(--vscode-panel-border,#44534b)}button.secondary:hover{background:var(--vscode-button-secondaryHoverBackground,#3b4a41)}button:focus-visible,summary:focus-visible{outline:2px solid var(--vscode-focusBorder,#69d69c);outline-offset:4px}button:disabled{opacity:.55;cursor:wait}
.cards{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px;margin:20px 0 34px}.card{border:1px solid var(--vscode-panel-border,#3c5145);border-radius:12px;padding:22px}.number{font-size:12px;font-weight:700;color:var(--vscode-descriptionForeground,#a7b8ad)}.card p,.check p{color:var(--vscode-descriptionForeground,#a7b8ad);font-size:13px}.tag{display:inline-block;margin-top:8px;font-size:11px;border:1px solid var(--vscode-panel-border,#44534b);border-radius:5px;padding:3px 8px}
.checks{list-style:none;margin:16px 0 24px;padding:0;border:1px solid var(--vscode-panel-border,#3c5145);border-radius:12px;overflow:hidden}.check{padding:17px 20px;border-bottom:1px solid var(--vscode-panel-border,#3c5145)}.check:last-child{border:0}.check p{margin:7px 0 0}.badge{font-size:11px;border:1px solid currentColor;border-radius:99px;padding:2px 9px;white-space:nowrap}.ok{color:var(--vscode-testing-iconPassed,#71d69b)}.warning{color:var(--vscode-editorWarning-foreground,#e3bd68)}.blocked{color:var(--vscode-errorForeground,#ff9393)}.unknown{color:var(--vscode-descriptionForeground,#a7b8ad)}
.note{border-left:3px solid var(--vscode-editorWarning-foreground,#d4ad55);padding:5px 16px;margin:16px 0 26px;color:var(--vscode-descriptionForeground,#a7b8ad);font-size:13px}.help{border-top:1px solid var(--vscode-panel-border,#3c5145);padding-top:20px}.help summary{cursor:pointer;font-weight:600}.help ol{padding-left:22px}.help li{margin:10px 0}.help code{font-family:var(--vscode-editor-font-family,monospace);overflow-wrap:anywhere}footer{font-size:11px;color:var(--vscode-descriptionForeground,#a7b8ad);padding-top:24px}#feedback{min-height:24px;font-size:13px}
@media(max-width:700px){main{padding:22px 18px}.hero{padding:25px}.cards{grid-template-columns:1fr;gap:10px}.section-heading{align-items:flex-start;flex-direction:column}.check-heading{align-items:flex-start;flex-wrap:wrap}.topline{flex-wrap:wrap}}@media(forced-colors:active){.hero,.card,.checks,button,.mark{border:1px solid CanvasText}.badge{color:CanvasText}}
</style></head><body><main>
<header class="topline"><div class="brand"><span class="mark" aria-hidden="true">▦</span> Excel AI &amp; VBA Studio</div><span class="preview">${snapshot ? `v${escapeHtml(snapshot.extensionVersion)} · ` : ''}Preview</span></header>
<section class="hero" aria-labelledby="welcome"><p class="eyebrow">Votre espace Excel dans VS Code</p><h1 id="welcome">Du classeur au code.<br>Un point de départ clair.</h1><p class="intro">Ouvrez votre tableur, retrouvez votre projet VBA et préparez votre contexte IA. Commencez avec une copie de test de votre fichier.</p>
<div class="actions"><button data-action="openWorkbook">Ouvrir un classeur…</button><button class="secondary" data-action="refresh">Vérifier ma configuration</button></div></section>
<section aria-labelledby="workflows"><h2 id="workflows">Choisissez votre façon de travailler</h2><div class="cards">
<article class="card"><span class="number">01 / TABLEUR</span><h3>Vos données, au premier plan</h3><p>Consultez et modifiez vos fichiers XLSX, CSV et TSV dans la grille intégrée.</p><span class="tag">Sans Excel pour la grille</span></article>
<article class="card"><span class="number">02 / VBA STUDIO</span><h3>Du code lié au bon classeur</h3><p>Ouvrez un fichier Excel, puis utilisez « Ouvrir le studio VBA » dans la palette de commandes.</p><span class="tag">Fonctions natives : Excel requis</span></article>
<article class="card"><span class="number">03 / ASSISTANCE IA</span><h3>Un contexte choisi par vous</h3><p>Exportez le contexte à la demande. Utilisez #excelVbaWorkbook dans votre chat compatible.</p><span class="tag">Chat IA facultatif</span></article>
</div></section>
<section aria-labelledby="setup"><div class="section-heading"><h2 id="setup">Votre configuration locale</h2><button class="secondary" data-action="copyReport" ${snapshot ? '' : 'disabled'}>Copier le diagnostic</button></div><ul class="checks" aria-live="polite">${checks}</ul>
<p class="note">Ce contrôle ne lance pas Excel et ne valide pas ses autorisations VBA. Il ne modifie aucune politique Office. Les macros s’exécutent dans Microsoft Excel, pas dans l’aperçu VS Code.</p></section>
<details class="help"><summary>Installer sur un autre PC ou résoudre un problème</summary><ol>
<li>Sur Windows x64, utilisez VS Code 1.95 ou ultérieur. Dans Extensions → … → <strong>Installer à partir d’un VSIX</strong>, choisissez le paquet Windows x64 fourni, puis rechargez la fenêtre.</li>
<li>Dans la palette (<strong>Ctrl+Maj+P</strong>), recherchez <strong>Excel AI &amp; VBA Studio : Accueil et diagnostic</strong>. Vérifiez la version et les prérequis ci-dessus.</li>
<li>Testez d’abord un CSV ou XLSX synthétique : ouverture, modification, enregistrement et réouverture. Ne commencez pas avec votre seul exemplaire d’un fichier important.</li>
<li>Pour le VBA et les opérations natives, Microsoft Excel desktop doit être installé. Utilisez le Centre de sécurité du classeur pour comprendre les blocages ; ne baissez pas les protections de votre organisation.</li>
<li>Si une erreur de manifeste mentionne <code>\\p{L}</code> ou <code>\\p{N}</code>, vérifiez que le nouveau VSIX est installé dans le bon profil VS Code et rechargez la fenêtre.</li>
</ol><p>Pour demander de l’aide, joignez le diagnostic, le message d’erreur exact et les étapes de reproduction. N’envoyez pas de classeur confidentiel.</p></details>
<p id="feedback" role="status" aria-live="polite"></p><footer>Diagnostic local uniquement · Aucun envoi réseau · Aucun contenu de classeur dans le rapport</footer>
</main><script nonce="${nonce}">
const api = acquireVsCodeApi();
document.querySelectorAll('button[data-action]').forEach(button => button.addEventListener('click', () => {
  api.postMessage({ action: button.dataset.action });
}));
window.addEventListener('message', event => {
  if (event.data && event.data.type === 'feedback' && typeof event.data.text === 'string') {
    document.getElementById('feedback').textContent = event.data.text;
  }
});
</script></body></html>`;
}
