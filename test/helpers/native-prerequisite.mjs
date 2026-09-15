import { spawnSync } from 'node:child_process';

export function skipNativeTest(reason) {
  if (process.env.EXCEL_AI_REQUIRE_NATIVE_TESTS === '1') {
    throw new Error(`Required native acceptance test cannot run: ${reason}`);
  }
  console.log(`SKIPPED native acceptance: ${reason}`);
  process.exit(0);
}
export function requireExcelRegistration(label) {
  if (!isExcelRegistered()) skipNativeTest(`${label} requires Windows and a registered Microsoft Excel COM installation.`);
}

export function isExcelRegistered() {
  if (process.platform !== 'win32') return false;
  const probe = spawnSync('powershell.exe', ['-NoLogo', '-NoProfile', '-NonInteractive', '-Command',
    "if ($null -eq [type]::GetTypeFromProgID('Excel.Application')) { exit 2 }"],
  { encoding: 'utf8', windowsHide: true, timeout: 30_000 });
  if (probe.error) throw probe.error;
  if (probe.status === 2) return false;
  if (probe.status !== 0) throw new Error(`Excel prerequisite probe failed: ${probe.stderr}`);
  return true;
}
